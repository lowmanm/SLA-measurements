/**
 * Email Service
 * 
 * This file contains all functions related to sending email notifications
 * and importing data from Gmail.
 */

/**
 * Send an email notification
 * 
 * @param {string} recipient - Email address of recipient
 * @param {string} subject - Email subject
 * @param {string} body - Email body (HTML)
 * @param {Object} options - Additional options
 */
function sendEmailNotification(recipient, subject, body, options = {}) {
  // Check if email notifications are enabled
  const notificationsEnabled = getSetting('notification_email_enabled', 'true');
  if (notificationsEnabled !== 'true') {
    Logger.log('Email notifications are disabled');
    return;
  }
  
  try {
    const emailOptions = {
      htmlBody: body,
      name: 'QA Platform'
    };
    
    // Add optional parameters if provided
    if (options.cc) emailOptions.cc = options.cc;
    if (options.bcc) emailOptions.bcc = options.bcc;
    if (options.attachments) emailOptions.attachments = options.attachments;
    
    GmailApp.sendEmail(recipient, subject, body, emailOptions);
    Logger.log(`Email sent to ${recipient}`);
  } catch (error) {
    Logger.log(`Error sending email: ${error.message}`);
    throw error;
  }
}

/**
 * Send notification about a new evaluation
 * 
 * @param {Object} evaluationData - Data about the evaluation
 */
function sendNewEvaluationNotification(evaluationData) {
  // Get agent's email
  const usersData = getDataFromSheet(SHEET_NAMES.USERS);
  const agent = usersData.find(user => user.Name === evaluationData.Agent);
  
  if (!agent) {
    Logger.log(`Agent ${evaluationData.Agent} not found in users list`);
    return;
  }
  
  const scorePct = Math.round((evaluationData.Score / evaluationData.MaxPossible) * 100);
  const passingScore = getSetting('passing_score_percentage', '80');
  const passed = scorePct >= parseInt(passingScore);
  
  const subject = `New QA Evaluation - ${passed ? 'Passed' : 'Needs Improvement'} - Score: ${scorePct}%`;
  
  const body = `
    <h2>Quality Assurance Evaluation</h2>
    <p>Dear ${evaluationData.Agent},</p>
    <p>A new quality assurance evaluation has been completed for you.</p>
    
    <div style="margin: 20px 0; padding: 15px; border-radius: 5px; background-color: ${passed ? '#e6f4ea' : '#fce8e6'};">
      <h3 style="margin-top: 0;">Evaluation Results</h3>
      <p><strong>Score:</strong> ${evaluationData.Score} out of ${evaluationData.MaxPossible} (${scorePct}%)</p>
      <p><strong>Status:</strong> ${passed ? 'Passed' : 'Needs Improvement'}</p>
      <p><strong>Evaluated By:</strong> ${evaluationData.Evaluator}</p>
      <p><strong>Date:</strong> ${evaluationData.EvaluationDate.toDateString()}</p>
    </div>
    
    <p>You can view the detailed results in the QA Platform.</p>
    
    <p>If you disagree with this evaluation, you can file a dispute within 
    ${getSetting('dispute_window_days', '7')} days by accessing the QA Platform and selecting "File Dispute".</p>
    
    <p>Thank you,<br>QA Team</p>
  `;
  
  sendEmailNotification(agent.Email, subject, body);
  
  // Also notify the agent's manager if configured
  const notifyManager = getSetting('notify_manager_on_evaluation', 'true');
  if (notifyManager === 'true') {
    const managers = usersData.filter(user => user.Role === USER_ROLES.AGENT_MANAGER);
    if (managers.length > 0) {
      const managerEmails = managers.map(m => m.Email).join(',');
      sendEmailNotification(managerEmails, `Agent Evaluation: ${evaluationData.Agent}`, body);
    }
  }
}

/**
 * Send notification about a dispute being filed
 * 
 * @param {Object} disputeData - Data about the dispute
 * @param {Object} evaluationData - Data about the evaluation being disputed
 */
function sendDisputeFiledNotification(disputeData, evaluationData) {
  // Get QA managers' emails
  const usersData = getDataFromSheet(SHEET_NAMES.USERS);
  const qaManagers = usersData.filter(user => user.Role === USER_ROLES.QA_MANAGER);
  
  if (qaManagers.length === 0) {
    Logger.log('No QA managers found to notify about dispute');
    return;
  }
  
  const subject = `Dispute Filed - Evaluation ID: ${disputeData.EvaluationID}`;
  
  const body = `
    <h2>Quality Assurance Dispute</h2>
    <p>A new dispute has been filed:</p>
    
    <div style="margin: 20px 0; padding: 15px; border-radius: 5px; background-color: #f1f3f4;">
      <h3 style="margin-top: 0;">Dispute Details</h3>
      <p><strong>Evaluation ID:</strong> ${disputeData.EvaluationID}</p>
      <p><strong>Agent:</strong> ${evaluationData.Agent}</p>
      <p><strong>Filed By:</strong> ${disputeData.FiledBy}</p>
      <p><strong>Date Filed:</strong> ${disputeData.DateFiled.toDateString()}</p>
      <p><strong>Original Score:</strong> ${disputeData.OriginalScore}</p>
    </div>
    
    <p>Please review this dispute in the QA Platform.</p>
    
    <p>Thank you,<br>QA Platform</p>
  `;
  
  // Send to all QA managers
  const managerEmails = qaManagers.map(manager => manager.Email).join(',');
  sendEmailNotification(managerEmails, subject, body);
  
  // Also notify the QA analyst who performed the evaluation
  const evaluator = evaluationData.Evaluator;
  const evaluatorData = usersData.find(user => user.Name === evaluator || user.Email === evaluator);
  
  if (evaluatorData) {
    sendEmailNotification(evaluatorData.Email, `Your evaluation has been disputed - ID: ${disputeData.EvaluationID}`, body);
  }
}

/**
 * Send notification about a dispute resolution
 * 
 * @param {Object} disputeData - Data about the resolved dispute
 * @param {Object} evaluationData - Data about the evaluation
 */
function sendDisputeResolutionNotification(disputeData, evaluationData) {
  // Get agent's email
  const usersData = getDataFromSheet(SHEET_NAMES.USERS);
  const agent = usersData.find(user => user.Name === evaluationData.Agent);
  
  if (!agent) {
    Logger.log(`Agent ${evaluationData.Agent} not found in users list`);
    return;
  }
  
  const scoreChanged = disputeData.OriginalScore !== disputeData.AdjustedScore;
  const subject = `Dispute Resolution - ${scoreChanged ? 'Score Adjusted' : 'Original Score Maintained'}`;
  
  const body = `
    <h2>Dispute Resolution</h2>
    <p>Dear ${evaluationData.Agent},</p>
    <p>Your dispute for evaluation ID ${disputeData.EvaluationID} has been reviewed and resolved.</p>
    
    <div style="margin: 20px 0; padding: 15px; border-radius: 5px; background-color: #f1f3f4;">
      <h3 style="margin-top: 0;">Resolution Details</h3>
      <p><strong>Original Score:</strong> ${disputeData.OriginalScore}</p>
      <p><strong>Adjusted Score:</strong> ${disputeData.AdjustedScore}</p>
      <p><strong>Reviewed By:</strong> ${disputeData.Reviewer}</p>
      <p><strong>Resolution Date:</strong> ${disputeData.ResolutionDate.toDateString()}</p>
      <p><strong>Notes:</strong> ${disputeData.ResolutionNotes}</p>
    </div>
    
    <p>You can view the updated evaluation in the QA Platform.</p>
    
    <p>Thank you,<br>QA Team</p>
  `;
  
  sendEmailNotification(agent.Email, subject, body);
  
  // Also notify the person who filed the dispute if different from the agent
  if (disputeData.FiledBy !== agent.Email && disputeData.FiledBy !== agent.Name) {
    const filer = usersData.find(user => user.Email === disputeData.FiledBy || user.Name === disputeData.FiledBy);
    if (filer) {
      sendEmailNotification(filer.Email, `Dispute Resolution - Evaluation ID: ${disputeData.EvaluationID}`, body);
    }
  }
}

/**
 * Import audit queue data from Gmail
 * Searches for emails with CSV attachments and imports them
 */
function importFromGmail() {
  const searchQuery = 'has:attachment filename:csv';
  const searchLabel = getSetting('import_gmail_label', 'QA-Import');
  const processedLabel = getSetting('processed_gmail_label', 'QA-Processed');
  
  // Add label search if configured
  const fullQuery = searchLabel ? `${searchQuery} label:${searchLabel}` : searchQuery;
  
  try {
    // Search for emails with CSV attachments
    const threads = GmailApp.search(fullQuery, 0, 10); // Limit to 10 threads
    
    if (threads.length === 0) {
      Logger.log('No emails found matching criteria');
      return;
    }
    
    let importCount = 0;
    
    for (const thread of threads) {
      const messages = thread.getMessages();
      
      for (const message of messages) {
        const attachments = message.getAttachments();
        
        for (const attachment of attachments) {
          // Check if it's a CSV file
          if (attachment.getName().toLowerCase().endsWith('.csv')) {
            // Process the CSV file
            importCount += processCSVAttachment(attachment);
            
            // Mark message as processed
            if (processedLabel) {
              const label = getOrCreateLabel(processedLabel);
              thread.addLabel(label);
              
              // Remove the import label if it exists
              if (searchLabel) {
                const importLabel = GmailApp.getUserLabelByName(searchLabel);
                if (importLabel) {
                  thread.removeLabel(importLabel);
                }
              }
            }
          }
        }
      }
    }
    
    if (importCount > 0) {
      SpreadsheetApp.getUi().alert(`Import complete. ${importCount} records added to audit queue.`);
    } else {
      SpreadsheetApp.getUi().alert('Import complete. No new records were added.');
    }
    
  } catch (error) {
    Logger.log(`Error importing from Gmail: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error during import: ${error.message}`);
  }
}

/**
 * Process a CSV attachment and import data to audit queue
 * 
 * @param {GmailAttachment} attachment - The Gmail attachment to process
 * @return {number} Number of records imported
 */
function processCSVAttachment(attachment) {
  const csvData = Utilities.parseCsv(attachment.getDataAsString());
  const auditSheet = spreadsheet.getSheetByName(SHEET_NAMES.AUDIT_QUEUE);
  
  // Skip if there's no data or just headers
  if (!csvData || csvData.length <= 1) {
    return 0;
  }
  
  const headers = csvData[0];
  const expectedHeaders = ['ID', 'Date', 'Agent', 'Customer', 'Interaction Type', 'Duration'];
  
  // Check required headers
  for (const header of expectedHeaders) {
    if (!headers.includes(header)) {
      throw new Error(`Required header "${header}" missing in CSV file`);
    }
  }
  
  // Process each row
  let importCount = 0;
  
  for (let i = 1; i < csvData.length; i++) {
    const row = csvData[i];
    if (row.length !== headers.length) continue; // Skip malformed rows
    
    // Create record object
    const record = {};
    for (let j = 0; j < headers.length; j++) {
      record[headers[j]] = row[j];
    }
    
    // Check if this record already exists
    const existingData = findRowsInSheet(SHEET_NAMES.AUDIT_QUEUE, { ID: record.ID });
    if (existingData.length > 0) {
      continue; // Skip existing records
    }
    
    // Add default values for queue-specific fields
    record['Queue Status'] = 'New';
    record['Assigned To'] = '';
    record['Priority'] = 'Normal';
    record['Import Date'] = new Date();
    
    // Add to audit queue
    addRowToSheet(SHEET_NAMES.AUDIT_QUEUE, record);
    importCount++;
  }
  
  return importCount;
}

/**
 * Get or create a Gmail label
 * 
 * @param {string} labelName - Name of the label
 * @return {GmailLabel} The Gmail label
 */
function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  
  return label;
}