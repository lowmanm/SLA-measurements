/**
 * Admin.gs
 * 
 * This file contains administrative functions for the QA Platform.
 */

/**
 * Get all question sets
 * 
 * @return {Array} Array of question set objects
 */
function getAllQuestionSets() {
  return getDataFromSheet(SHEET_NAMES.QUESTION_SETS);
}

/**
 * Get a question set by ID
 * 
 * @param {string} questionSetId - The question set ID
 * @return {Object} Question set object or null if not found
 */
function getQuestionSetById(questionSetId) {
  return getRowById(SHEET_NAMES.QUESTION_SETS, questionSetId);
}

/**
 * Create or update a question set
 * 
 * @param {Object} questionSetData - The question set data
 * @return {Object} Result object with success flag, message, and question set ID
 */
function saveQuestionSet(questionSetData) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to manage question sets'
      };
    }
    
    // Validate required fields
    if (!questionSetData.Name || !questionSetData.Description || !questionSetData['Interaction Type']) {
      return {
        success: false,
        message: 'Required fields missing (Name, Description, or Interaction Type)'
      };
    }
    
    const currentUser = Session.getActiveUser().getEmail();
    const now = new Date();
    
    // Check if this is a new question set or an update
    if (!questionSetData.ID) {
      // New question set
      
      // Generate ID
      const questionSetId = generateUniqueId();
      
      // Prepare question set record
      const questionSetRecord = {
        ID: questionSetId,
        Name: questionSetData.Name,
        Description: questionSetData.Description,
        Category: questionSetData.Category || '',
        'Interaction Type': questionSetData['Interaction Type'],
        'Created By': currentUser,
        'Last Modified': now
      };
      
      // Save question set
      if (!addRowToSheet(SHEET_NAMES.QUESTION_SETS, questionSetRecord)) {
        return {
          success: false,
          message: 'Failed to save question set'
        };
      }
      
      // Log the action
      logAction(currentUser, 'Create Question Set', 
                `Created question set: ${questionSetData.Name}`);
      
      return {
        success: true,
        message: 'Question set created successfully',
        questionSetId: questionSetId
      };
    } else {
      // Updating existing question set
      
      // Get the existing question set
      const existingQuestionSet = getQuestionSetById(questionSetData.ID);
      if (!existingQuestionSet) {
        return {
          success: false,
          message: `Question set not found: ${questionSetData.ID}`
        };
      }
      
      // Prepare update data
      const updateData = {
        Name: questionSetData.Name,
        Description: questionSetData.Description,
        Category: questionSetData.Category || '',
        'Interaction Type': questionSetData['Interaction Type'],
        'Last Modified': now
      };
      
      // Update the question set
      if (!updateRowInSheet(SHEET_NAMES.QUESTION_SETS, questionSetData.ID, updateData)) {
        return {
          success: false,
          message: 'Failed to update question set'
        };
      }
      
      // Log the action
      logAction(currentUser, 'Update Question Set', 
                `Updated question set: ${questionSetData.Name}`);
      
      return {
        success: true,
        message: 'Question set updated successfully',
        questionSetId: questionSetData.ID
      };
    }
  } catch (error) {
    Logger.log(`Error in saveQuestionSet: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Delete a question set
 * 
 * @param {string} questionSetId - The ID of the question set to delete
 * @return {Object} Result object with success flag and message
 */
function deleteQuestionSet(questionSetId) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to delete question sets'
      };
    }
    
    // Get the existing question set
    const existingQuestionSet = getQuestionSetById(questionSetId);
    if (!existingQuestionSet) {
      return {
        success: false,
        message: `Question set not found: ${questionSetId}`
      };
    }
    
    // Check if this question set is used in any evaluations
    const evaluations = getFilteredData(SHEET_NAMES.EVALUATIONS, { 'Question Set ID': questionSetId });
    if (evaluations.length > 0) {
      return {
        success: false,
        message: `Cannot delete question set because it is used in ${evaluations.length} evaluation(s)`
      };
    }
    
    // Delete associated questions first
    const questions = getQuestionsForQuestionSet(questionSetId);
    questions.forEach(question => {
      deleteRowFromSheet(SHEET_NAMES.QUESTIONS, question.ID);
    });
    
    // Delete the question set
    if (!deleteRowFromSheet(SHEET_NAMES.QUESTION_SETS, questionSetId)) {
      return {
        success: false,
        message: 'Failed to delete question set'
      };
    }
    
    // Log the action
    const currentUser = Session.getActiveUser().getEmail();
    logAction(currentUser, 'Delete Question Set', 
              `Deleted question set: ${existingQuestionSet.Name}`);
    
    return {
      success: true,
      message: `Question set "${existingQuestionSet.Name}" and its questions deleted successfully`
    };
  } catch (error) {
    Logger.log(`Error in deleteQuestionSet: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Get all questions
 * 
 * @return {Array} Array of question objects
 */
function getAllQuestions() {
  return getDataFromSheet(SHEET_NAMES.QUESTIONS);
}

/**
 * Get a question by ID
 * 
 * @param {string} questionId - The question ID
 * @return {Object} Question object or null if not found
 */
function getQuestionById(questionId) {
  return getRowById(SHEET_NAMES.QUESTIONS, questionId);
}

/**
 * Create or update a question
 * 
 * @param {Object} questionData - The question data
 * @return {Object} Result object with success flag, message, and question ID
 */
function saveQuestion(questionData) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to manage questions'
      };
    }
    
    // Validate required fields
    if (!questionData['Question Set ID'] || !questionData.Question || 
        !questionData.Type || questionData['Possible Score'] === undefined) {
      return {
        success: false,
        message: 'Required fields missing (Question Set ID, Question, Type, or Possible Score)'
      };
    }
    
    // Verify the question set exists
    const questionSet = getQuestionSetById(questionData['Question Set ID']);
    if (!questionSet) {
      return {
        success: false,
        message: `Question set not found: ${questionData['Question Set ID']}`
      };
    }
    
    // Validate question type
    const validTypes = ['Yes/No', 'Multiple Choice', 'Numeric', 'Text'];
    if (!validTypes.includes(questionData.Type)) {
      return {
        success: false,
        message: `Invalid question type. Must be one of: ${validTypes.join(', ')}`
      };
    }
    
    // Validate possible score (must be a non-negative number)
    const possibleScore = parseInt(questionData['Possible Score']);
    if (isNaN(possibleScore) || possibleScore < 0) {
      return {
        success: false,
        message: 'Possible Score must be a non-negative number'
      };
    }
    
    // Check if this is a new question or an update
    if (!questionData.ID) {
      // New question
      
      // Generate ID
      const questionId = generateUniqueId();
      
      // Prepare question record
      const questionRecord = {
        ID: questionId,
        'Question Set ID': questionData['Question Set ID'],
        Question: questionData.Question,
        Type: questionData.Type,
        Weight: questionData.Weight || 1,
        'Possible Score': questionData['Possible Score'],
        Critical: questionData.Critical === true,
        Options: questionData.Options || '',
        'Help Text': questionData['Help Text'] || ''
      };
      
      // Save question
      if (!addRowToSheet(SHEET_NAMES.QUESTIONS, questionRecord)) {
        return {
          success: false,
          message: 'Failed to save question'
        };
      }
      
      // Log the action
      const currentUser = Session.getActiveUser().getEmail();
      logAction(currentUser, 'Create Question', 
                `Created question for set: ${questionSet.Name}`);
      
      return {
        success: true,
        message: 'Question created successfully',
        questionId: questionId
      };
    } else {
      // Updating existing question
      
      // Get the existing question
      const existingQuestion = getQuestionById(questionData.ID);
      if (!existingQuestion) {
        return {
          success: false,
          message: `Question not found: ${questionData.ID}`
        };
      }
      
      // Prepare update data
      const updateData = {
        'Question Set ID': questionData['Question Set ID'],
        Question: questionData.Question,
        Type: questionData.Type,
        Weight: questionData.Weight || 1,
        'Possible Score': questionData['Possible Score'],
        Critical: questionData.Critical === true,
        Options: questionData.Options || '',
        'Help Text': questionData['Help Text'] || ''
      };
      
      // Update the question
      if (!updateRowInSheet(SHEET_NAMES.QUESTIONS, questionData.ID, updateData)) {
        return {
          success: false,
          message: 'Failed to update question'
        };
      }
      
      // Log the action
      const currentUser = Session.getActiveUser().getEmail();
      logAction(currentUser, 'Update Question', 
                `Updated question for set: ${questionSet.Name}`);
      
      return {
        success: true,
        message: 'Question updated successfully',
        questionId: questionData.ID
      };
    }
  } catch (error) {
    Logger.log(`Error in saveQuestion: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Delete a question
 * 
 * @param {string} questionId - The ID of the question to delete
 * @return {Object} Result object with success flag and message
 */
function deleteQuestion(questionId) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to delete questions'
      };
    }
    
    // Get the existing question
    const existingQuestion = getQuestionById(questionId);
    if (!existingQuestion) {
      return {
        success: false,
        message: `Question not found: ${questionId}`
      };
    }
    
    // Delete the question
    if (!deleteRowFromSheet(SHEET_NAMES.QUESTIONS, questionId)) {
      return {
        success: false,
        message: 'Failed to delete question'
      };
    }
    
    // Log the action
    const currentUser = Session.getActiveUser().getEmail();
    logAction(currentUser, 'Delete Question', 
              `Deleted question: ${existingQuestion.Question}`);
    
    return {
      success: true,
      message: 'Question deleted successfully'
    };
  } catch (error) {
    Logger.log(`Error in deleteQuestion: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Get all settings
 * 
 * @return {Array} Array of setting objects
 */
function getAllSettingsData() {
  return getDataFromSheet(SHEET_NAMES.SETTINGS);
}

/**
 * Save a setting
 * 
 * @param {string} key - The setting key
 * @param {string} value - The setting value
 * @param {string} description - The setting description (optional)
 * @return {Object} Result object with success flag and message
 */
function saveSetting(key, value, description) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to modify settings'
      };
    }
    
    // Validate key
    if (!key) {
      return {
        success: false,
        message: 'Setting key is required'
      };
    }
    
    // Get settings
    const settings = getAllSettingsData();
    const existingSetting = settings.find(s => s.Key === key);
    
    if (existingSetting) {
      // Update existing setting
      if (!updateRowInSheet(SHEET_NAMES.SETTINGS, key, { 
        Value: value,
        Description: description || existingSetting.Description
      })) {
        return {
          success: false,
          message: 'Failed to update setting'
        };
      }
    } else {
      // Add new setting
      if (!addRowToSheet(SHEET_NAMES.SETTINGS, {
        Key: key,
        Value: value,
        Description: description || ''
      })) {
        return {
          success: false,
          message: 'Failed to add setting'
        };
      }
    }
    
    // Log the action
    const currentUser = Session.getActiveUser().getEmail();
    logAction(currentUser, 'Update Setting', 
              `Updated setting: ${key} = ${value}`);
    
    return {
      success: true,
      message: 'Setting saved successfully'
    };
  } catch (error) {
    Logger.log(`Error in saveSetting: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Delete a setting
 * 
 * @param {string} key - The setting key to delete
 * @return {Object} Result object with success flag and message
 */
function deleteSetting(key) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to modify settings'
      };
    }
    
    // Get the setting
    const settings = getAllSettingsData();
    const existingSetting = settings.find(s => s.Key === key);
    
    if (!existingSetting) {
      return {
        success: false,
        message: `Setting not found: ${key}`
      };
    }
    
    // Check if this is a protected setting
    const protectedSettings = [
      'version', 'company_name', 'platform_name', 
      'passing_score_percentage', 'dispute_time_limit_days'
    ];
    
    if (protectedSettings.includes(key)) {
      return {
        success: false,
        message: `Cannot delete protected setting: ${key}`
      };
    }
    
    // Delete the setting
    if (!deleteRowFromSheet(SHEET_NAMES.SETTINGS, key)) {
      return {
        success: false,
        message: 'Failed to delete setting'
      };
    }
    
    // Log the action
    const currentUser = Session.getActiveUser().getEmail();
    logAction(currentUser, 'Delete Setting', 
              `Deleted setting: ${key}`);
    
    return {
      success: true,
      message: 'Setting deleted successfully'
    };
  } catch (error) {
    Logger.log(`Error in deleteSetting: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Get system information
 * 
 * @return {Object} System information object
 */
function getSystemInfo() {
  try {
    // Get settings
    const settings = getSettings();
    
    // Get sheet counts
    const usersCount = getDataFromSheet(SHEET_NAMES.USERS).length;
    const questionSetsCount = getDataFromSheet(SHEET_NAMES.QUESTION_SETS).length;
    const questionsCount = getDataFromSheet(SHEET_NAMES.QUESTIONS).length;
    const evaluationsCount = getDataFromSheet(SHEET_NAMES.EVALUATIONS).length;
    const disputesCount = getDataFromSheet(SHEET_NAMES.DISPUTES).length;
    const logsCount = getDataFromSheet(SHEET_NAMES.LOGS).length;
    
    // Get sheet sizes (approximation based on row count)
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSizes = {};
    
    for (const sheetName in SHEET_NAMES) {
      const sheet = spreadsheet.getSheetByName(SHEET_NAMES[sheetName]);
      if (sheet) {
        const lastRow = sheet.getLastRow();
        const lastColumn = sheet.getLastColumn();
        
        if (lastRow > 0 && lastColumn > 0) {
          const data = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
          let sizeEstimate = 0;
          
          for (const row of data) {
            for (const cell of row) {
              // Estimate size in bytes (rough approximation)
              sizeEstimate += String(cell).length * 2; // Unicode characters can be 2 bytes
            }
          }
          
          sheetSizes[SHEET_NAMES[sheetName]] = formatBytes(sizeEstimate);
        } else {
          sheetSizes[SHEET_NAMES[sheetName]] = formatBytes(0);
        }
      }
    }
    
    // Get recent activity
    const logs = getRecentLogs(10);
    
    return {
      version: settings.version || '1.0.0',
      companyName: settings.company_name || 'My Company',
      platformName: settings.platform_name || 'QA Platform',
      passingThreshold: settings.passing_score_percentage || '80',
      disputeTimeLimit: settings.dispute_time_limit_days || '5',
      counts: {
        users: usersCount,
        questionSets: questionSetsCount,
        questions: questionsCount,
        evaluations: evaluationsCount,
        disputes: disputesCount,
        logs: logsCount
      },
      sheetSizes: sheetSizes,
      recentActivity: logs,
      systemTime: new Date()
    };
  } catch (error) {
    Logger.log(`Error in getSystemInfo: ${error.message}`);
    return {
      error: error.message
    };
  }
}

/**
 * Format bytes to a human-readable format
 * 
 * @param {number} bytes - The number of bytes
 * @param {number} decimals - The number of decimal places
 * @return {string} Formatted string
 */
function formatBytes(bytes, decimals = 2) {
  if (bytes === 0) return '0 Bytes';
  
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
  
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  
  return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

/**
 * Backup the database to a separate spreadsheet
 * 
 * @return {Object} Result object with success flag and message
 */
function backupDatabase() {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to backup the database'
      };
    }
    
    // Create a new spreadsheet
    const currentUser = Session.getActiveUser().getEmail();
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const backupName = `QA_Platform_Backup_${formattedDate}`;
    
    const backup = SpreadsheetApp.create(backupName);
    const backupId = backup.getId();
    
    // Get source spreadsheet
    const source = SpreadsheetApp.getActiveSpreadsheet();
    
    // Copy each sheet
    for (const sheetName in SHEET_NAMES) {
      const sourceSheet = source.getSheetByName(SHEET_NAMES[sheetName]);
      
      if (sourceSheet) {
        // Delete the backup's initial sheet if this is the first copy
        if (sheetName === 'USERS' && backup.getSheets().length === 1) {
          backup.deleteSheet(backup.getSheets()[0]);
        }
        
        // Copy the sheet content
        const destSheet = backup.insertSheet(SHEET_NAMES[sheetName]);
        
        const lastRow = sourceSheet.getLastRow();
        const lastCol = sourceSheet.getLastColumn();
        
        if (lastRow > 0 && lastCol > 0) {
          // Copy data
          const data = sourceSheet.getRange(1, 1, lastRow, lastCol).getValues();
          destSheet.getRange(1, 1, lastRow, lastCol).setValues(data);
          
          // Copy formatting
          if (sourceSheet.getFrozenRows() > 0) {
            destSheet.setFrozenRows(sourceSheet.getFrozenRows());
          }
          
          // Copy column widths
          for (let i = 1; i <= lastCol; i++) {
            destSheet.setColumnWidth(i, sourceSheet.getColumnWidth(i));
          }
        }
      }
    }
    
    // Add metadata sheet
    const metaSheet = backup.insertSheet('Metadata');
    metaSheet.getRange('A1:B1').setValues([['Backup Information', '']]);
    metaSheet.getRange('A2:B2').setValues([['Backup Date', now]]);
    metaSheet.getRange('A3:B3').setValues([['Backup By', currentUser]]);
    metaSheet.getRange('A4:B4').setValues([['Source Spreadsheet', source.getName()]]);
    metaSheet.getRange('A5:B5').setValues([['Backup Spreadsheet ID', backupId]]);
    
    // Format the metadata sheet
    metaSheet.getRange('A1:B1').merge();
    metaSheet.getRange('A1:A5').setFontWeight('bold');
    metaSheet.autoResizeColumns(1, 2);
    
    // Make the backup accessible to the current user
    const file = DriveApp.getFileById(backupId);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Get the URL
    const url = file.getUrl();
    
    // Log the action
    logAction(currentUser, 'Backup Database', 
              `Created backup: ${backupName}`);
    
    return {
      success: true,
      message: 'Database backup created successfully',
      backupId: backupId,
      backupUrl: url,
      backupName: backupName
    };
  } catch (error) {
    Logger.log(`Error in backupDatabase: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Clear old data from the database
 * 
 * @param {Object} options - Options for data clearing
 * @return {Object} Result object with success flag and message
 */
function clearOldData(options = {}) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to clear data'
      };
    }
    
    // Default options
    const defaults = {
      olderThan: 365, // days
      evaluations: true,
      disputes: true,
      logs: true,
      backup: true
    };
    
    // Merge options with defaults
    const opts = { ...defaults, ...options };
    
    // Calculate cutoff date
    const now = new Date();
    const cutoffDate = new Date(now);
    cutoffDate.setDate(cutoffDate.getDate() - opts.olderThan);
    
    // Create backup first if requested
    let backupResult = null;
    if (opts.backup) {
      backupResult = backupDatabase();
      if (!backupResult.success) {
        return {
          success: false,
          message: `Failed to create backup: ${backupResult.message}`
        };
      }
    }
    
    let totalDeleted = 0;
    
    // Clear old logs
    if (opts.logs) {
      const logs = getDataFromSheet(SHEET_NAMES.LOGS);
      let logsDeleted = 0;
      
      for (let i = logs.length - 1; i >= 0; i--) {
        const log = logs[i];
        const logDate = log.Timestamp instanceof Date ? log.Timestamp : new Date(log.Timestamp);
        
        if (logDate < cutoffDate) {
          if (deleteRowFromSheet(SHEET_NAMES.LOGS, i + 2)) { // +2 for header row and 0-indexing
            logsDeleted++;
          }
        }
      }
      
      totalDeleted += logsDeleted;
    }
    
    // Clear old disputes and associated data
    if (opts.disputes) {
      const disputes = getDataFromSheet(SHEET_NAMES.DISPUTES);
      let disputesDeleted = 0;
      
      for (let i = disputes.length - 1; i >= 0; i--) {
        const dispute = disputes[i];
        const disputeDate = dispute['Submission Date'] instanceof Date ? 
          dispute['Submission Date'] : new Date(dispute['Submission Date']);
        
        if (disputeDate < cutoffDate) {
          // Delete associated dispute resolutions
          const resolutions = getFilteredData(SHEET_NAMES.DISPUTE_RESOLUTIONS, { 'Dispute ID': dispute.ID });
          resolutions.forEach(resolution => {
            deleteRowFromSheet(SHEET_NAMES.DISPUTE_RESOLUTIONS, resolution.ID);
          });
          
          // Delete the dispute
          if (deleteRowFromSheet(SHEET_NAMES.DISPUTES, dispute.ID)) {
            disputesDeleted++;
          }
        }
      }
      
      totalDeleted += disputesDeleted;
    }
    
    // Clear old evaluations and associated data
    if (opts.evaluations) {
      const evaluations = getDataFromSheet(SHEET_NAMES.EVALUATIONS);
      let evaluationsDeleted = 0;
      
      for (let i = evaluations.length - 1; i >= 0; i--) {
        const evaluation = evaluations[i];
        const evalDate = evaluation.Date instanceof Date ? evaluation.Date : new Date(evaluation.Date);
        
        if (evalDate < cutoffDate) {
          // Check if there are any pending disputes for this evaluation
          const pendingDisputes = getFilteredData(SHEET_NAMES.DISPUTES, {
            'Evaluation ID': evaluation.ID,
            Status: STATUS.PENDING
          });
          
          if (pendingDisputes.length === 0) {
            // Delete associated answers
            const answers = getFilteredData(SHEET_NAMES.EVALUATION_ANSWERS, { 'Evaluation ID': evaluation.ID });
            answers.forEach(answer => {
              deleteRowFromSheet(SHEET_NAMES.EVALUATION_ANSWERS, answer.ID);
            });
            
            // Delete the evaluation
            if (deleteRowFromSheet(SHEET_NAMES.EVALUATIONS, evaluation.ID)) {
              evaluationsDeleted++;
            }
          }
        }
      }
      
      totalDeleted += evaluationsDeleted;
    }
    
    // Log the action
    const currentUser = Session.getActiveUser().getEmail();
    logAction(currentUser, 'Clear Old Data', 
              `Cleared ${totalDeleted} records older than ${opts.olderThan} days`);
    
    return {
      success: true,
      message: `Cleared ${totalDeleted} records older than ${opts.olderThan} days`,
      backupResult: backupResult
    };
  } catch (error) {
    Logger.log(`Error in clearOldData: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}