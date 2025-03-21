/**
 * ImportExport.gs
 * 
 * This file contains functions for importing and exporting data.
 */

/**
 * Import audit queue items from Gmail attachments
 * 
 * @param {Object} options - Import options
 * @return {Object} Result object with success flag, message, and import stats
 */
function importFromGmail(options = {}) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to import data'
      };
    }
    
    // Default options
    const defaults = {
      searchQuery: 'subject:"QA Audit Data" has:attachment',
      filePattern: /audit.*\.csv$/i,
      maxEmails: 10,
      processedLabel: 'QA-Processed',
      skipProcessed: true
    };
    
    // Merge options with defaults
    const opts = { ...defaults, ...options };
    
    // Search for emails with CSV attachments
    const threads = GmailApp.search(opts.searchQuery, 0, opts.maxEmails);
    
    // Create or get the processed label
    let processedLabel;
    try {
      processedLabel = GmailApp.createLabel(opts.processedLabel);
    } catch (e) {
      processedLabel = GmailApp.getUserLabelByName(opts.processedLabel);
    }
    
    let importStats = {
      emailsProcessed: 0,
      emailsSkipped: 0,
      filesProcessed: 0,
      recordsImported: 0,
      recordsSkipped: 0,
      errors: []
    };
    
    // Process each thread
    threads.forEach(thread => {
      // Skip already processed threads if required
      if (opts.skipProcessed && thread.getLabels().some(label => label.getName() === opts.processedLabel)) {
        importStats.emailsSkipped++;
        return;
      }
      
      const messages = thread.getMessages();
      
      // Process each message in the thread
      messages.forEach(message => {
        const attachments = message.getAttachments();
        
        // Process each attachment
        attachments.forEach(attachment => {
          const fileName = attachment.getName();
          
          // Check if the file matches the pattern
          if (opts.filePattern.test(fileName)) {
            try {
              // Get CSV content
              const csvContent = attachment.getDataAsString();
              
              // Process the CSV
              const importResult = processAuditCsv(csvContent);
              
              // Update stats
              importStats.filesProcessed++;
              importStats.recordsImported += importResult.imported;
              importStats.recordsSkipped += importResult.skipped;
              
              if (importResult.errors.length > 0) {
                importStats.errors = importStats.errors.concat(importResult.errors);
              }
            } catch (err) {
              importStats.errors.push(`Error processing ${fileName}: ${err.message}`);
            }
          }
        });
      });
      
      // Mark thread as processed
      thread.addLabel(processedLabel);
      importStats.emailsProcessed++;
    });
    
    // Log the action
    const currentUser = Session.getActiveUser().getEmail();
    logAction(currentUser, 'Import from Gmail', 
              `Imported ${importStats.recordsImported} records from ${importStats.filesProcessed} files`);
    
    return {
      success: true,
      message: `Imported ${importStats.recordsImported} records from ${importStats.filesProcessed} files`,
      stats: importStats
    };
  } catch (error) {
    Logger.log(`Error in importFromGmail: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Process audit CSV data
 * 
 * @param {string} csvContent - The CSV content as a string
 * @return {Object} Result object with import statistics
 */
function processAuditCsv(csvContent) {
  const result = {
    imported: 0,
    skipped: 0,
    errors: []
  };
  
  try {
    // Parse CSV
    const rows = Utilities.parseCsv(csvContent);
    
    // Check for empty data
    if (rows.length < 2) {
      result.errors.push('CSV file has no data rows');
      return result;
    }
    
    // Get headers
    const headers = rows[0];
    
    // Check required fields
    const requiredFields = ['Date', 'Agent', 'Customer ID', 'Interaction ID', 'Interaction Type'];
    
    // Verify headers contain required fields
    const missingFields = requiredFields.filter(field => !headers.includes(field));
    if (missingFields.length > 0) {
      result.errors.push(`CSV missing required fields: ${missingFields.join(', ')}`);
      return result;
    }
    
    // Get existing audit queue for duplicate checking
    const existingQueue = getDataFromSheet(SHEET_NAMES.AUDIT_QUEUE);
    
    // Process each data row (skip header)
    for (let i = 1; i < rows.length; i++) {
      try {
        const row = rows[i];
        
        // Skip empty rows
        if (row.every(cell => !cell)) {
          continue;
        }
        
        // Create record object
        const record = {};
        headers.forEach((header, index) => {
          record[header] = row[index] || '';
        });
        
        // Generate ID
        record.ID = generateUniqueId();
        
        // Set default values for missing fields
        record.Status = record.Status || STATUS.PENDING;
        record.Priority = record.Priority || 'Normal';
        record['Assigned To'] = record['Assigned To'] || '';
        
        // Convert date if needed
        if (record.Date && !(record.Date instanceof Date)) {
          try {
            record.Date = new Date(record.Date);
          } catch (e) {
            // Leave as is if parsing fails
          }
        }
        
        // Calculate due date if not provided
        if (!record['Due Date']) {
          const dueDate = new Date(record.Date || new Date());
          dueDate.setDate(dueDate.getDate() + 7); // Default: 7 days after interaction
          record['Due Date'] = dueDate;
        }
        
        // Check for duplicate
        const isDuplicate = existingQueue.some(item => 
          item['Interaction ID'] === record['Interaction ID'] &&
          item.Agent === record.Agent
        );
        
        if (isDuplicate) {
          result.skipped++;
          continue;
        }
        
        // Add to sheet
        if (addRowToSheet(SHEET_NAMES.AUDIT_QUEUE, record)) {
          result.imported++;
        } else {
          result.errors.push(`Failed to add row ${i} to sheet`);
          result.skipped++;
        }
      } catch (rowError) {
        result.errors.push(`Error processing row ${i}: ${rowError.message}`);
        result.skipped++;
      }
    }
    
    return result;
  } catch (error) {
    result.errors.push(`Error processing CSV: ${error.message}`);
    return result;
  }
}

/**
 * Export evaluations to CSV
 * 
 * @param {Object} options - Export options
 * @return {Object} Result object with success flag, message, and CSV data
 */
function exportEvaluationsToCSV(options = {}) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.QA_MANAGER) && !hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to export data'
      };
    }
    
    // Default options
    const defaults = {
      startDate: null,
      endDate: null,
      agent: null,
      evaluator: null
    };
    
    // Merge options with defaults
    const opts = { ...defaults, ...options };
    
    // Get all evaluations
    let evaluations = getAllEvaluations();
    
    // Apply filters
    if (opts.startDate) {
      const startDate = opts.startDate instanceof Date ? opts.startDate : new Date(opts.startDate);
      evaluations = evaluations.filter(eval => {
        const evalDate = eval.Date instanceof Date ? eval.Date : new Date(eval.Date);
        return evalDate >= startDate;
      });
    }
    
    if (opts.endDate) {
      const endDate = opts.endDate instanceof Date ? opts.endDate : new Date(opts.endDate);
      evaluations = evaluations.filter(eval => {
        const evalDate = eval.Date instanceof Date ? eval.Date : new Date(eval.Date);
        return evalDate <= endDate;
      });
    }
    
    if (opts.agent) {
      evaluations = evaluations.filter(eval => eval.Agent === opts.agent);
    }
    
    if (opts.evaluator) {
      evaluations = evaluations.filter(eval => eval.Evaluator === opts.evaluator);
    }
    
    // Create CSV content
    let csvRows = [];
    
    // Add headers
    const headers = [
      'ID', 'Date', 'Agent', 'Agent Name', 'Evaluator', 'Evaluator Name',
      'Interaction Type', 'Customer ID', 'Interaction ID',
      'Score', 'Max Possible', 'Percentage', 'Status',
      'Strengths', 'Areas for Improvement', 'Comments'
    ];
    csvRows.push(headers.join(','));
    
    // Add evaluation data
    evaluations.forEach(eval => {
      // Get agent and evaluator names
      const agent = getUserByEmail(eval.Agent);
      const agentName = agent ? agent.Name : '';
      
      const evaluator = getUserByEmail(eval.Evaluator);
      const evaluatorName = evaluator ? evaluator.Name : '';
      
      // Calculate percentage
      const score = parseInt(eval.Score);
      const maxPossible = parseInt(eval['Max Possible']);
      const percentage = maxPossible > 0 ? (score / maxPossible * 100).toFixed(1) : '0.0';
      
      // Format date
      const date = eval.Date instanceof Date ?
        Utilities.formatDate(eval.Date, Session.getScriptTimeZone(), 'yyyy-MM-dd') :
        eval.Date;
      
      // Create row
      const row = [
        eval.ID,
        date,
        eval.Agent,
        csvEscape(agentName),
        eval.Evaluator,
        csvEscape(evaluatorName),
        csvEscape(eval['Interaction Type'] || ''),
        csvEscape(eval['Customer ID'] || ''),
        csvEscape(eval['Interaction ID'] || ''),
        eval.Score,
        eval['Max Possible'],
        percentage,
        eval.Status,
        csvEscape(eval.Strengths || ''),
        csvEscape(eval['Areas for Improvement'] || ''),
        csvEscape(eval.Comments || '')
      ];
      
      csvRows.push(row.join(','));
    });
    
    // Join all rows
    const csvContent = csvRows.join('\n');
    
    // Log the action
    const currentUser = Session.getActiveUser().getEmail();
    logAction(currentUser, 'Export Evaluations', 
              `Exported ${evaluations.length} evaluations to CSV`);
    
    return {
      success: true,
      message: `Exported ${evaluations.length} evaluations to CSV`,
      data: csvContent
    };
  } catch (error) {
    Logger.log(`Error in exportEvaluationsToCSV: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Upload CSV data to the audit queue
 * 
 * @param {string} csvContent - The CSV content as a string
 * @return {Object} Result object with success flag, message, and import stats
 */
function uploadAuditCSV(csvContent) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.QA_MANAGER) && !hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to upload audit data'
      };
    }
    
    // Process the CSV
    const result = processAuditCsv(csvContent);
    
    // Log the action
    const currentUser = Session.getActiveUser().getEmail();
    logAction(currentUser, 'Upload Audit CSV', 
              `Imported ${result.imported} records, skipped ${result.skipped}`);
    
    return {
      success: result.errors.length === 0,
      message: `Imported ${result.imported} records, skipped ${result.skipped}` + 
               (result.errors.length > 0 ? `. Errors: ${result.errors.join('; ')}` : ''),
      stats: result
    };
  } catch (error) {
    Logger.log(`Error in uploadAuditCSV: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Export audit queue to CSV
 * 
 * @param {Object} options - Export options
 * @return {Object} Result object with success flag, message, and CSV data
 */
function exportAuditQueueToCSV(options = {}) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.QA_MANAGER) && !hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to export data'
      };
    }
    
    // Default options
    const defaults = {
      status: null,
      assignedTo: null
    };
    
    // Merge options with defaults
    const opts = { ...defaults, ...options };
    
    // Get audit queue
    let auditQueue = getDataFromSheet(SHEET_NAMES.AUDIT_QUEUE);
    
    // Apply filters
    if (opts.status) {
      auditQueue = auditQueue.filter(item => item.Status === opts.status);
    }
    
    if (opts.assignedTo) {
      auditQueue = auditQueue.filter(item => item['Assigned To'] === opts.assignedTo);
    }
    
    // Create CSV content
    let csvRows = [];
    
    // Add headers
    const headers = [
      'ID', 'Date', 'Agent', 'Customer ID', 'Interaction ID', 'Interaction Type',
      'Assigned To', 'Status', 'Priority', 'Due Date', 'Notes'
    ];
    csvRows.push(headers.join(','));
    
    // Add audit queue data
    auditQueue.forEach(item => {
      // Format dates
      const date = item.Date instanceof Date ?
        Utilities.formatDate(item.Date, Session.getScriptTimeZone(), 'yyyy-MM-dd') :
        item.Date;
      
      const dueDate = item['Due Date'] instanceof Date ?
        Utilities.formatDate(item['Due Date'], Session.getScriptTimeZone(), 'yyyy-MM-dd') :
        item['Due Date'];
      
      // Create row
      const row = [
        item.ID,
        date,
        item.Agent,
        csvEscape(item['Customer ID'] || ''),
        csvEscape(item['Interaction ID'] || ''),
        csvEscape(item['Interaction Type'] || ''),
        item['Assigned To'] || '',
        item.Status || '',
        item.Priority || '',
        dueDate,
        csvEscape(item.Notes || '')
      ];
      
      csvRows.push(row.join(','));
    });
    
    // Join all rows
    const csvContent = csvRows.join('\n');
    
    // Log the action
    const currentUser = Session.getActiveUser().getEmail();
    logAction(currentUser, 'Export Audit Queue', 
              `Exported ${auditQueue.length} audit queue items to CSV`);
    
    return {
      success: true,
      message: `Exported ${auditQueue.length} audit queue items to CSV`,
      data: csvContent
    };
  } catch (error) {
    Logger.log(`Error in exportAuditQueueToCSV: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Escape a string for CSV format
 * 
 * @param {string} str - The string to escape
 * @return {string} The escaped string
 */
function csvEscape(str) {
  if (!str) return '';
  
  // Check if the string needs to be quoted
  if (str.includes(',') || str.includes('"') || str.includes('\n')) {
    // Replace double quotes with two double quotes
    str = str.replace(/"/g, '""');
    // Wrap in quotes
    return `"${str}"`;
  }
  
  return str;
}

/**
 * Get available export formats
 * 
 * @return {Array} Array of export format objects
 */
function getExportFormats() {
  return [
    {
      id: 'evaluations_csv',
      name: 'Evaluations (CSV)',
      description: 'Export evaluations data as a CSV file'
    },
    {
      id: 'audit_queue_csv',
      name: 'Audit Queue (CSV)',
      description: 'Export audit queue data as a CSV file'
    },
    {
      id: 'disputes_csv',
      name: 'Disputes (CSV)',
      description: 'Export disputes data as a CSV file'
    },
    {
      id: 'users_csv',
      name: 'Users (CSV)',
      description: 'Export users data as a CSV file'
    }
  ];
}

/**
 * Export data in the requested format
 * 
 * @param {string} format - The export format ID
 * @param {Object} options - Export options
 * @return {Object} Result object with success flag, message, and data
 */
function exportData(format, options = {}) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.QA_MANAGER) && !hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to export data'
      };
    }
    
    // Handle different export formats
    switch (format) {
      case 'evaluations_csv':
        return exportEvaluationsToCSV(options);
        
      case 'audit_queue_csv':
        return exportAuditQueueToCSV(options);
        
      case 'disputes_csv':
        return exportDisputesToCSV(options);
        
      case 'users_csv':
        return exportUsersToCSV(options);
        
      default:
        return {
          success: false,
          message: `Unsupported export format: ${format}`
        };
    }
  } catch (error) {
    Logger.log(`Error in exportData: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Export disputes to CSV
 * 
 * @param {Object} options - Export options
 * @return {Object} Result object with success flag, message, and CSV data
 */
function exportDisputesToCSV(options = {}) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.QA_MANAGER) && !hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to export data'
      };
    }
    
    // Default options
    const defaults = {
      status: null,
      startDate: null,
      endDate: null
    };
    
    // Merge options with defaults
    const opts = { ...defaults, ...options };
    
    // Get all disputes
    let disputes = getAllDisputes();
    
    // Apply filters
    if (opts.status) {
      disputes = disputes.filter(dispute => dispute.Status === opts.status);
    }
    
    if (opts.startDate) {
      const startDate = opts.startDate instanceof Date ? opts.startDate : new Date(opts.startDate);
      disputes = disputes.filter(dispute => {
        const disputeDate = dispute['Submission Date'] instanceof Date ? 
          dispute['Submission Date'] : new Date(dispute['Submission Date']);
        return disputeDate >= startDate;
      });
    }
    
    if (opts.endDate) {
      const endDate = opts.endDate instanceof Date ? opts.endDate : new Date(opts.endDate);
      disputes = disputes.filter(dispute => {
        const disputeDate = dispute['Submission Date'] instanceof Date ? 
          dispute['Submission Date'] : new Date(dispute['Submission Date']);
        return disputeDate <= endDate;
      });
    }
    
    // Create CSV content
    let csvRows = [];
    
    // Add headers
    const headers = [
      'ID', 'Evaluation ID', 'Agent', 'Submitted By', 'Submission Date',
      'Reason', 'Status', 'Reviewed By', 'Review Date',
      'Score Adjustment', 'Details', 'Additional Evidence', 'Review Notes'
    ];
    csvRows.push(headers.join(','));
    
    // Add dispute data
    disputes.forEach(dispute => {
      // Get associated evaluation to get agent
      const evaluation = getEvaluationById(dispute['Evaluation ID']);
      const agent = evaluation ? evaluation.Agent : '';
      
      // Format dates
      const submissionDate = dispute['Submission Date'] instanceof Date ?
        Utilities.formatDate(dispute['Submission Date'], Session.getScriptTimeZone(), 'yyyy-MM-dd') :
        dispute['Submission Date'];
      
      const reviewDate = dispute['Review Date'] instanceof Date ?
        Utilities.formatDate(dispute['Review Date'], Session.getScriptTimeZone(), 'yyyy-MM-dd') :
        dispute['Review Date'] || '';
      
      // Create row
      const row = [
        dispute.ID,
        dispute['Evaluation ID'],
        agent,
        dispute['Submitted By'],
        submissionDate,
        csvEscape(dispute.Reason || ''),
        dispute.Status,
        dispute['Reviewed By'] || '',
        reviewDate,
        dispute['Score Adjustment'] || '0',
        csvEscape(dispute.Details || ''),
        csvEscape(dispute['Additional Evidence'] || ''),
        csvEscape(dispute['Review Notes'] || '')
      ];
      
      csvRows.push(row.join(','));
    });
    
    // Join all rows
    const csvContent = csvRows.join('\n');
    
    // Log the action
    const currentUser = Session.getActiveUser().getEmail();
    logAction(currentUser, 'Export Disputes', 
              `Exported ${disputes.length} disputes to CSV`);
    
    return {
      success: true,
      message: `Exported ${disputes.length} disputes to CSV`,
      data: csvContent
    };
  } catch (error) {
    Logger.log(`Error in exportDisputesToCSV: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Export users to CSV
 * 
 * @param {Object} options - Export options
 * @return {Object} Result object with success flag, message, and CSV data
 */
function exportUsersToCSV(options = {}) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to export user data'
      };
    }
    
    // Default options
    const defaults = {
      role: null,
      department: null
    };
    
    // Merge options with defaults
    const opts = { ...defaults, ...options };
    
    // Get all users
    let users = getAllUsers();
    
    // Apply filters
    if (opts.role) {
      users = users.filter(user => user.Role === opts.role);
    }
    
    if (opts.department) {
      users = users.filter(user => user.Department === opts.department);
    }
    
    // Create CSV content
    let csvRows = [];
    
    // Add headers
    const headers = [
      'ID', 'Name', 'Email', 'Role', 'Department', 'Manager', 'Created', 'Last Login'
    ];
    csvRows.push(headers.join(','));
    
    // Add user data
    users.forEach(user => {
      // Format dates
      const created = user.Created instanceof Date ?
        Utilities.formatDate(user.Created, Session.getScriptTimeZone(), 'yyyy-MM-dd') :
        user.Created;
      
      const lastLogin = user['Last Login'] instanceof Date ?
        Utilities.formatDate(user['Last Login'], Session.getScriptTimeZone(), 'yyyy-MM-dd') :
        user['Last Login'] || '';
      
      // Create row
      const row = [
        user.ID,
        csvEscape(user.Name),
        user.Email,
        user.Role,
        csvEscape(user.Department || ''),
        user.Manager || '',
        created,
        lastLogin
      ];
      
      csvRows.push(row.join(','));
    });
    
    // Join all rows
    const csvContent = csvRows.join('\n');
    
    // Log the action
    const currentUser = Session.getActiveUser().getEmail();
    logAction(currentUser, 'Export Users', 
              `Exported ${users.length} users to CSV`);
    
    return {
      success: true,
      message: `Exported ${users.length} users to CSV`,
      data: csvContent
    };
  } catch (error) {
    Logger.log(`Error in exportUsersToCSV: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}