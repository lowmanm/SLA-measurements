/**
 * Dispute Functions
 * 
 * This file contains functions for managing disputes to evaluations,
 * including filing, reviewing, and resolving disputes.
 */

/**
 * Show the dispute form UI
 */
function showDisputeForm() {
  if (!hasPermission(USER_ROLES.AGENT_MANAGER)) {
    SpreadsheetApp.getUi().alert('You do not have permission to file disputes');
    return;
  }
  
  const html = HtmlService.createTemplateFromFile('UI/DisputeForm')
    .evaluate()
    .setTitle('File Dispute')
    .setWidth(800)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'File Dispute');
}

/**
 * Show the dispute review UI
 */
function showDisputeReview() {
  if (!hasPermission(USER_ROLES.QA_MANAGER)) {
    SpreadsheetApp.getUi().alert('You do not have permission to review disputes');
    return;
  }
  
  const html = HtmlService.createTemplateFromFile('UI/DisputeReview')
    .evaluate()
    .setTitle('Review Disputes')
    .setWidth(800)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Review Disputes');
}

/**
 * Get evaluations available for dispute
 * 
 * @return {Array} Array of evaluations that can be disputed
 */
function getEvaluationsForDispute() {
  if (!hasPermission(USER_ROLES.AGENT_MANAGER)) {
    throw new Error('You do not have permission to view evaluations for dispute');
  }
  
  const evaluations = getDataFromSheet(SHEET_NAMES.EVALUATIONS);
  const disputes = getDataFromSheet(SHEET_NAMES.DISPUTES);
  
  // Get the dispute window in days
  const disputeWindowDays = parseInt(getSetting('dispute_window_days', '7'));
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - disputeWindowDays);
  
  // Get already disputed evaluation IDs
  const disputedEvalIds = disputes.map(dispute => dispute['Evaluation ID']);
  
  // Filter evaluations that are not already disputed and within the dispute window
  return evaluations.filter(eval => {
    const evalDate = new Date(eval['Evaluation Date']);
    return !disputedEvalIds.includes(eval.ID) && 
           evalDate >= cutoffDate &&
           eval.Status === 'Completed';
  });
}

/**
 * File a new dispute
 * 
 * @param {string} evaluationId - ID of the evaluation being disputed
 * @param {string} reason - Reason for the dispute
 * @param {Array} questionDisputes - Array of disputed questions with reasoning
 * @return {Object} Newly created dispute data
 */
function fileDispute(evaluationId, reason, questionDisputes) {
  if (!hasPermission(USER_ROLES.AGENT_MANAGER)) {
    throw new Error('You do not have permission to file disputes');
  }
  
  // Get the evaluation
  const evaluation = findRowById(SHEET_NAMES.EVALUATIONS, evaluationId);
  if (!evaluation) {
    throw new Error(`Evaluation with ID ${evaluationId} not found`);
  }
  
  // Check if this evaluation already has a dispute
  const existingDisputes = findRowsInSheet(SHEET_NAMES.DISPUTES, { 'Evaluation ID': evaluationId });
  if (existingDisputes.length > 0) {
    throw new Error('This evaluation already has a dispute filed');
  }
  
  // Get the dispute window in days
  const disputeWindowDays = parseInt(getSetting('dispute_window_days', '7'));
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - disputeWindowDays);
  
  // Check if the evaluation is still within the dispute window
  const evalDate = new Date(evaluation['Evaluation Date']);
  if (evalDate < cutoffDate) {
    throw new Error('This evaluation is outside the dispute window');
  }
  
  // Create the dispute record
  const disputeId = generateUniqueId();
  const dispute = {
    ID: disputeId,
    'Evaluation ID': evaluationId,
    'Date Filed': new Date(),
    'Filed By': currentUser,
    Status: 'Pending',
    Reviewer: '',
    'Resolution Date': '',
    'Resolution Notes': '',
    'Original Score': evaluation.Score,
    'Adjusted Score': ''
  };
  
  const rowIndex = addRowToSheet(SHEET_NAMES.DISPUTES, dispute);
  
  // Update the evaluation to mark it as disputed
  updateRowInSheet(SHEET_NAMES.EVALUATIONS, evaluation.rowIndex, {
    Disputed: true,
    'Last Updated': new Date()
  });
  
  // Store the detailed dispute information in the properties service
  const disputeDetail = {
    disputeId: disputeId,
    evaluationId: evaluationId,
    reason: reason,
    questionDisputes: questionDisputes,
    dateFiled: new Date().toISOString()
  };
  
  const userProperties = PropertiesService.getScriptProperties();
  userProperties.setProperty(`dispute_${disputeId}`, JSON.stringify(disputeDetail));
  
  // Send notification to QA Managers
  const manager = getUserByEmail(currentUser);
  const evaluationData = getEvaluationWithDetails(evaluationId);
  
  notifyQAManagerAboutDispute(dispute, evaluation, manager);
  
  return {
    dispute: dispute,
    detail: disputeDetail
  };
}

/**
 * Get pending disputes for review
 * 
 * @return {Array} Array of pending disputes
 */
function getPendingDisputes() {
  if (!hasPermission(USER_ROLES.QA_MANAGER)) {
    throw new Error('You do not have permission to view pending disputes');
  }
  
  const disputes = getDataFromSheet(SHEET_NAMES.DISPUTES);
  
  return disputes.filter(dispute => dispute.Status === 'Pending');
}

/**
 * Get a dispute by ID with details
 * 
 * @param {string} disputeId - ID of the dispute
 * @return {Object} Dispute with details
 */
function getDisputeWithDetails(disputeId) {
  const dispute = findRowById(SHEET_NAMES.DISPUTES, disputeId);
  if (!dispute) {
    throw new Error(`Dispute with ID ${disputeId} not found`);
  }
  
  // Get the detailed dispute information
  const userProperties = PropertiesService.getScriptProperties();
  const detailJson = userProperties.getProperty(`dispute_${disputeId}`);
  
  if (!detailJson) {
    throw new Error(`Detailed information for dispute ${disputeId} not found`);
  }
  
  const detail = JSON.parse(detailJson);
  
  // Get the associated evaluation
  const evaluation = getEvaluationWithDetails(dispute['Evaluation ID']);
  
  return {
    dispute: dispute,
    detail: detail,
    evaluation: evaluation
  };
}

/**
 * Resolve a dispute
 * 
 * @param {string} disputeId - ID of the dispute
 * @param {string} resolution - Resolution status ('Approved' or 'Denied')
 * @param {string} notes - Resolution notes
 * @param {number} adjustedScore - Adjusted score (if approved)
 * @return {Object} Updated dispute data
 */
function resolveDispute(disputeId, resolution, notes, adjustedScore) {
  if (!hasPermission(USER_ROLES.QA_MANAGER)) {
    throw new Error('You do not have permission to resolve disputes');
  }
  
  // Get the dispute
  const dispute = findRowById(SHEET_NAMES.DISPUTES, disputeId);
  if (!dispute) {
    throw new Error(`Dispute with ID ${disputeId} not found`);
  }
  
  // Check if the dispute is still pending
  if (dispute.Status !== 'Pending') {
    throw new Error('This dispute has already been resolved');
  }
  
  // Update the dispute record
  const updatedDispute = {
    Status: resolution,
    Reviewer: currentUser,
    'Resolution Date': new Date(),
    'Resolution Notes': notes,
    'Adjusted Score': resolution === 'Approved' ? adjustedScore : dispute['Original Score']
  };
  
  updateRowInSheet(SHEET_NAMES.DISPUTES, dispute.rowIndex, updatedDispute);
  
  // If approved, update the evaluation score
  if (resolution === 'Approved') {
    const evaluation = findRowById(SHEET_NAMES.EVALUATIONS, dispute['Evaluation ID']);
    
    if (evaluation) {
      updateRowInSheet(SHEET_NAMES.EVALUATIONS, evaluation.rowIndex, {
        Score: adjustedScore,
        'Last Updated': new Date()
      });
    }
  }
  
  // Send notifications
  const qaManager = getUserByEmail(currentUser);
  const manager = getUserByEmail(dispute['Filed By']);
  
  const evaluation = findRowById(SHEET_NAMES.EVALUATIONS, dispute['Evaluation ID']);
  const agent = getUserByEmail(evaluation.Agent);
  
  if (agent && manager) {
    notifyAboutDisputeResolution(
      { ...dispute, ...updatedDispute },
      evaluation,
      agent,
      manager,
      qaManager
    );
  }
  
  return {
    dispute: { ...dispute, ...updatedDispute },
    resolution: resolution
  };
}

/**
 * Get dispute statistics
 * 
 * @param {Date} startDate - Start date for the period
 * @param {Date} endDate - End date for the period
 * @return {Object} Statistics object
 */
function getDisputeStats(startDate, endDate) {
  if (!hasPermission(USER_ROLES.QA_MANAGER)) {
    throw new Error('You do not have permission to view dispute statistics');
  }
  
  const disputes = getDataFromSheet(SHEET_NAMES.DISPUTES);
  
  // Filter disputes by date range
  const periodDisputes = disputes.filter(dispute => {
    const disputeDate = new Date(dispute['Date Filed']);
    return disputeDate >= startDate && disputeDate <= endDate;
  });
  
  if (periodDisputes.length === 0) {
    return {
      count: 0,
      approvalRate: 0,
      pendingCount: 0,
      avgResolutionDays: 0
    };
  }
  
  // Calculate statistics
  const approved = periodDisputes.filter(d => d.Status === 'Approved').length;
  const pending = periodDisputes.filter(d => d.Status === 'Pending').length;
  const resolved = periodDisputes.filter(d => d.Status !== 'Pending');
  
  // Average resolution time in days
  let totalResolutionDays = 0;
  for (const dispute of resolved) {
    const filedDate = new Date(dispute['Date Filed']);
    const resolvedDate = new Date(dispute['Resolution Date']);
    const days = (resolvedDate - filedDate) / (1000 * 60 * 60 * 24);
    totalResolutionDays += days;
  }
  
  const avgResolutionDays = resolved.length > 0 ? 
    (totalResolutionDays / resolved.length).toFixed(1) : 0;
  
  return {
    count: periodDisputes.length,
    approvalRate: ((approved / (periodDisputes.length - pending)) * 100).toFixed(1),
    pendingCount: pending,
    avgResolutionDays: avgResolutionDays
  };
}