/**
 * Import/Export Functions
 * 
 * This file contains functions for importing data into the QA platform
 * and exporting reports and data.
 */

/**
 * Import audit queue items from Gmail
 * Looks for emails with CSV attachments containing audit data
 */
function importFromGmail() {
  if (!hasPermission(USER_ROLES.QA_MANAGER) && !hasPermission(USER_ROLES.ADMIN)) {
    SpreadsheetApp.getUi().alert('You do not have permission to import audit data');
    return;
  }
  
  try {
    // Look for emails with subject containing "QA Audit Data"
    const threads = GmailApp.search('subject:"QA Audit Data" has:attachment newer_than:7d', 0, 10);
    let importCount = 0;
    
    if (threads.length === 0) {
      SpreadsheetApp.getUi().alert('No emails with audit data found from the last 7 days');
      return;
    }
    
    // Process each thread
    for (const thread of threads) {
      const messages = thread.getMessages();
      
      for (const message of messages) {
        const attachments = message.getAttachments();
        
        for (const attachment of attachments) {
          // Only process CSV files
          if (attachment.getName().toLowerCase().endsWith('.csv')) {
            const csvData = attachment.getDataAsString();
            const importedCount = processCSVData(csvData);
            importCount += importedCount;
            
            // Mark the email as read
            message.markRead();
          }
        }
      }
      
      // Add a label to the processed thread
      const label = GmailApp.createLabel('Processed-QA-Import');
      thread.addLabel(label);
    }
    
    if (importCount > 0) {
      SpreadsheetApp.getUi().alert(`Successfully imported ${importCount} audit items`);
    } else {
      SpreadsheetApp.getUi().alert('No new audit items were imported');
    }
  } catch (error) {
    Logger.log(`Error importing from Gmail: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error importing data: ${error.message}`);
  }
}

/**
 * Process CSV data and import into audit queue
 * 
 * @param {string} csvData - CSV data as string
 * @return {number} Number of records imported
 */
function processCSVData(csvData) {
  const rows = Utilities.parseCsv(csvData);
  
  // Assume first row is headers
  const headers = rows[0];
  let importCount = 0;
  
  // Define field mappings between CSV headers and sheet columns
  // This allows for flexibility in the CSV format
  const fieldMappings = {
    'agent_email': 'Agent',
    'customer_id': 'Customer',
    'date': 'Date',
    'interaction_type': 'Interaction Type',
    'duration': 'Duration'
  };
  
  // Get existing IDs to avoid duplicates
  const existingAudits = getDataFromSheet(SHEET_NAMES.AUDIT_QUEUE);
  const existingIds = new Set(existingAudits.map(audit => audit.ID));
  
  // Process data rows
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row.length !== headers.length) continue; // Skip malformed rows
    
    // Create object from CSV row
    const rowData = {};
    for (let j = 0; j < headers.length; j++) {
      const header = headers[j].trim().toLowerCase();
      rowData[header] = row[j];
    }
    
    // Generate unique ID
    const auditId = generateUniqueId();
    
    // Skip if this appears to be a duplicate (checking a compound key)
    const compoundKey = `${rowData['agent_email']}_${rowData['date']}_${rowData['customer_id']}`;
    if (existingIds.has(compoundKey)) continue;
    
    // Map CSV data to sheet columns
    const auditItem = {
      ID: auditId,
      Date: new Date(rowData['date']),
      Agent: rowData['agent_email'],
      Customer: rowData['customer_id'],
      'Interaction Type': rowData['interaction_type'],
      Duration: rowData['duration'],
      'Queue Status': 'Pending',
      'Assigned To': '',
      Priority: determinePriority(rowData),
      'Import Date': new Date()
    };
    
    // Add to sheet
    addRowToSheet(SHEET_NAMES.AUDIT_QUEUE, auditItem);
    importCount++;
  }
  
  return importCount;
}

/**
 * Determine priority of an audit item based on business rules
 * 
 * @param {Object} rowData - Data for the audit item
 * @return {string} Priority ('High', 'Medium', or 'Low')
 */
function determinePriority(rowData) {
  // Example business rules for priority:
  // - High: Customer complaints or calls over 30 minutes
  // - Medium: Sales calls or calls between 15-30 minutes
  // - Low: All others
  
  const interactionType = rowData['interaction_type']?.toLowerCase();
  const duration = parseInt(rowData['duration']);
  
  if (interactionType?.includes('complaint') || duration > 30) {
    return 'High';
  } else if (interactionType?.includes('sales') || (duration >= 15 && duration <= 30)) {
    return 'Medium';
  } else {
    return 'Low';
  }
}

/**
 * Export evaluations to CSV
 * 
 * @param {Date} startDate - Start date for export range
 * @param {Date} endDate - End date for export range
 * @return {string} CSV data
 */
function exportEvaluationsToCSV(startDate, endDate) {
  if (!hasPermission(USER_ROLES.QA_MANAGER) && !hasPermission(USER_ROLES.ADMIN)) {
    throw new Error('You do not have permission to export evaluation data');
  }
  
  // Get evaluations in date range
  const evaluations = getDataFromSheet(SHEET_NAMES.EVALUATIONS);
  const filteredEvals = evaluations.filter(eval => {
    const evalDate = new Date(eval['Evaluation Date']);
    return evalDate >= startDate && evalDate <= endDate;
  });
  
  if (filteredEvals.length === 0) {
    return 'No evaluations found in the specified date range';
  }
  
  // Get headers from the first evaluation
  const headers = Object.keys(filteredEvals[0]);
  
  // Build CSV content
  let csvContent = headers.join(',') + '\\n';
  
  for (const eval of filteredEvals) {
    const values = headers.map(header => {
      let value = eval[header];
      
      // Format dates
      if (value instanceof Date) {
        value = value.toISOString();
      }
      
      // Escape commas and quotes
      if (typeof value === 'string' && (value.includes(',') || value.includes('"'))) {
        value = '"' + value.replace(/"/g, '""') + '"';
      }
      
      return value;
    });
    
    csvContent += values.join(',') + '\\n';
  }
  
  return csvContent;
}

/**
 * Generate a performance report
 * 
 * @param {string} reportType - Type of report ('agent', 'team', 'evaluator')
 * @param {Date} startDate - Start date for report
 * @param {Date} endDate - End date for report
 * @return {Object} Report data
 */
function generatePerformanceReport(reportType, startDate, endDate) {
  if (!hasPermission(USER_ROLES.QA_MANAGER) && !hasPermission(USER_ROLES.ADMIN)) {
    throw new Error('You do not have permission to generate performance reports');
  }
  
  const evaluations = getDataFromSheet(SHEET_NAMES.EVALUATIONS);
  
  // Filter evaluations in date range
  const filteredEvals = evaluations.filter(eval => {
    const evalDate = new Date(eval['Evaluation Date']);
    return evalDate >= startDate && evalDate <= endDate;
  });
  
  if (filteredEvals.length === 0) {
    return { message: 'No evaluations found in the specified date range' };
  }
  
  // Generate report based on type
  switch (reportType) {
    case 'agent':
      return generateAgentReport(filteredEvals);
    
    case 'team':
      return generateTeamReport(filteredEvals);
    
    case 'evaluator':
      return generateEvaluatorReport(filteredEvals);
    
    default:
      throw new Error(`Unknown report type: ${reportType}`);
  }
}

/**
 * Generate agent performance report
 * 
 * @param {Array} evaluations - Filtered evaluations
 * @return {Object} Report data
 */
function generateAgentReport(evaluations) {
  // Group by agent
  const agentGroups = {};
  
  for (const eval of evaluations) {
    const agent = eval.Agent;
    
    if (!agentGroups[agent]) {
      agentGroups[agent] = [];
    }
    
    agentGroups[agent].push(eval);
  }
  
  // Calculate stats for each agent
  const agentStats = [];
  const passingThreshold = parseInt(getSetting('passing_score_percentage', '80'));
  
  for (const agent in agentGroups) {
    const evals = agentGroups[agent];
    let totalScore = 0;
    let totalPossible = 0;
    let passingCount = 0;
    
    for (const eval of evals) {
      const score = parseInt(eval.Score);
      const maxPossible = parseInt(eval['Max Possible']);
      const percentage = (score / maxPossible) * 100;
      
      totalScore += score;
      totalPossible += maxPossible;
      
      if (percentage >= passingThreshold) {
        passingCount++;
      }
    }
    
    const avgScore = (totalScore / totalPossible) * 100;
    
    agentStats.push({
      agent: agent,
      evaluationCount: evals.length,
      averageScore: avgScore.toFixed(2),
      passingRate: ((passingCount / evals.length) * 100).toFixed(2),
      evaluations: evals
    });
  }
  
  // Sort by average score descending
  agentStats.sort((a, b) => parseFloat(b.averageScore) - parseFloat(a.averageScore));
  
  return {
    reportType: 'Agent Performance',
    totalEvaluations: evaluations.length,
    agentStats: agentStats
  };
}

/**
 * Generate team performance report
 * 
 * @param {Array} evaluations - Filtered evaluations
 * @return {Object} Report data
 */
function generateTeamReport(evaluations) {
  // Get all users to get department info
  const users = getAllUsers();
  const userMap = {};
  
  for (const user of users) {
    userMap[user.Email] = user;
  }
  
  // Group by department
  const departmentGroups = {};
  
  for (const eval of evaluations) {
    const agent = eval.Agent;
    const user = userMap[agent];
    const department = user ? user.Department : 'Unknown';
    
    if (!departmentGroups[department]) {
      departmentGroups[department] = [];
    }
    
    departmentGroups[department].push(eval);
  }
  
  // Calculate stats for each department
  const departmentStats = [];
  
  for (const department in departmentGroups) {
    const evals = departmentGroups[department];
    let totalScore = 0;
    let totalPossible = 0;
    
    for (const eval of evals) {
      totalScore += parseInt(eval.Score);
      totalPossible += parseInt(eval['Max Possible']);
    }
    
    const avgScore = (totalScore / totalPossible) * 100;
    
    departmentStats.push({
      department: department,
      evaluationCount: evals.length,
      averageScore: avgScore.toFixed(2),
      agentCount: new Set(evals.map(e => e.Agent)).size
    });
  }
  
  // Sort by average score descending
  departmentStats.sort((a, b) => parseFloat(b.averageScore) - parseFloat(a.averageScore));
  
  return {
    reportType: 'Team Performance',
    totalEvaluations: evaluations.length,
    departmentStats: departmentStats
  };
}

/**
 * Generate evaluator performance report
 * 
 * @param {Array} evaluations - Filtered evaluations
 * @return {Object} Report data
 */
function generateEvaluatorReport(evaluations) {
  // Group by evaluator
  const evaluatorGroups = {};
  
  for (const eval of evaluations) {
    const evaluator = eval.Evaluator;
    
    if (!evaluatorGroups[evaluator]) {
      evaluatorGroups[evaluator] = [];
    }
    
    evaluatorGroups[evaluator].push(eval);
  }
  
  // Calculate stats for each evaluator
  const evaluatorStats = [];
  
  for (const evaluator in evaluatorGroups) {
    const evals = evaluatorGroups[evaluator];
    let totalScore = 0;
    let totalPossible = 0;
    const disputedEvals = evals.filter(e => e.Disputed === true);
    
    for (const eval of evals) {
      totalScore += parseInt(eval.Score);
      totalPossible += parseInt(eval['Max Possible']);
    }
    
    const avgScore = (totalScore / totalPossible) * 100;
    
    evaluatorStats.push({
      evaluator: evaluator,
      evaluationCount: evals.length,
      averageScore: avgScore.toFixed(2),
      disputeRate: ((disputedEvals.length / evals.length) * 100).toFixed(2)
    });
  }
  
  // Sort by evaluation count descending
  evaluatorStats.sort((a, b) => b.evaluationCount - a.evaluationCount);
  
  return {
    reportType: 'Evaluator Performance',
    totalEvaluations: evaluations.length,
    evaluatorStats: evaluatorStats
  };
}