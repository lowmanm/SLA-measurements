/**
 * Dispute.gs
 * 
 * This file contains functions for handling evaluation disputes.
 */

/**
 * Get all disputes
 * 
 * @return {Array} Array of dispute objects
 */
function getAllDisputes() {
  return getDataFromSheet(SHEET_NAMES.DISPUTES);
}

/**
 * Get disputes based on filter criteria
 * 
 * @param {Object} filters - Object with filter criteria
 * @return {Array} Array of filtered dispute objects
 */
function getFilteredDisputes(filters) {
  return getFilteredData(SHEET_NAMES.DISPUTES, filters);
}

/**
 * Get a dispute by ID
 * 
 * @param {string} disputeId - The dispute ID
 * @return {Object} Dispute object or null if not found
 */
function getDisputeById(disputeId) {
  return getRowById(SHEET_NAMES.DISPUTES, disputeId);
}

/**
 * Get disputes for a specific evaluation
 * 
 * @param {string} evaluationId - The evaluation ID
 * @return {Array} Array of dispute objects for the evaluation
 */
function getDisputesForEvaluation(evaluationId) {
  return getFilteredData(SHEET_NAMES.DISPUTES, { 'Evaluation ID': evaluationId });
}

/**
 * Get disputes submitted by a specific user
 * 
 * @param {string} userEmail - The user's email
 * @return {Array} Array of dispute objects submitted by the user
 */
function getDisputesBySubmitter(userEmail) {
  return getFilteredData(SHEET_NAMES.DISPUTES, { 'Submitted By': userEmail });
}

/**
 * Get disputes for a specific agent
 * 
 * @param {string} agentEmail - The agent's email
 * @return {Array} Array of dispute objects for the agent
 */
function getDisputesForAgent(agentEmail) {
  // First get all evaluations for the agent
  const agentEvaluations = getEvaluationsForAgent(agentEmail);
  const evaluationIds = agentEvaluations.map(eval => eval.ID);
  
  // Then get all disputes for those evaluations
  const allDisputes = getAllDisputes();
  return allDisputes.filter(dispute => evaluationIds.includes(dispute['Evaluation ID']));
}

/**
 * Get pending disputes for review
 * 
 * @return {Array} Array of dispute objects with status 'Pending'
 */
function getPendingDisputes() {
  return getFilteredData(SHEET_NAMES.DISPUTES, { Status: STATUS.PENDING });
}

/**
 * File a new dispute
 * 
 * @param {Object} disputeData - The dispute data
 * @return {Object} Result object with success flag, message, and dispute ID
 */
function fileDispute(disputeData) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.AGENT_MANAGER) && !hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to file disputes'
      };
    }
    
    // Validate required fields
    if (!disputeData['Evaluation ID'] || !disputeData.Reason || !disputeData.Details) {
      return {
        success: false,
        message: 'Required fields missing (Evaluation ID, Reason, or Details)'
      };
    }
    
    // Verify the evaluation exists
    const evaluation = getEvaluationById(disputeData['Evaluation ID']);
    if (!evaluation) {
      return {
        success: false,
        message: `Evaluation not found: ${disputeData['Evaluation ID']}`
      };
    }
    
    // Check if this evaluation already has an active dispute
    const existingDisputes = getDisputesForEvaluation(disputeData['Evaluation ID']);
    const activeDispute = existingDisputes.find(d => 
      d.Status === STATUS.PENDING || d.Status === STATUS.IN_PROGRESS);
    
    if (activeDispute) {
      return {
        success: false,
        message: 'This evaluation already has an active dispute'
      };
    }
    
    // Check if the dispute time limit has passed
    const evaluationDate = evaluation.Date instanceof Date ? 
      evaluation.Date : new Date(evaluation.Date);
    const timeLimit = parseInt(getSetting('dispute_time_limit_days', '5'));
    const now = new Date();
    const limitDate = new Date(evaluationDate);
    limitDate.setDate(limitDate.getDate() + timeLimit);
    
    if (now > limitDate) {
      return {
        success: false,
        message: `The time limit of ${timeLimit} days for filing a dispute has passed`
      };
    }
    
    // Generate ID
    const disputeId = generateUniqueId();
    
    // Get current user as submitter if not specified
    const currentUser = Session.getActiveUser().getEmail();
    if (!disputeData['Submitted By']) {
      disputeData['Submitted By'] = currentUser;
    }
    
    // Prepare dispute record
    const disputeRecord = {
      ID: disputeId,
      'Evaluation ID': disputeData['Evaluation ID'],
      'Submitted By': disputeData['Submitted By'],
      'Submission Date': now,
      Reason: disputeData.Reason,
      Details: disputeData.Details,
      'Additional Evidence': disputeData['Additional Evidence'] || '',
      'Requested Score Change': disputeData['Requested Score Change'] || '',
      Status: STATUS.PENDING,
      'Reviewed By': '',
      'Review Date': '',
      'Score Adjustment': 0
    };
    
    // Save dispute
    if (!addRowToSheet(SHEET_NAMES.DISPUTES, disputeRecord)) {
      return {
        success: false,
        message: 'Failed to save dispute'
      };
    }
    
    // Update evaluation status
    if (!updateRowInSheet(SHEET_NAMES.EVALUATIONS, disputeData['Evaluation ID'], {
      Status: STATUS.DISPUTED
    })) {
      // Log warning but don't fail the operation
      Logger.log(`Warning: Failed to update evaluation status to Disputed`);
    }
    
    // Log the action
    logAction(currentUser, 'File Dispute', 
              `Filed dispute for evaluation ${disputeData['Evaluation ID']}`);
    
    // Send notification to the evaluator
    sendDisputeNotificationToEvaluator(disputeRecord, evaluation);
    
    return {
      success: true,
      message: 'Dispute filed successfully',
      disputeId: disputeId
    };
  } catch (error) {
    Logger.log(`Error in fileDispute: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Update an existing dispute
 * 
 * @param {string} disputeId - The ID of the dispute to update
 * @param {Object} updateData - The dispute data to update
 * @return {Object} Result object with success flag and message
 */
function updateDispute(disputeId, updateData) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.AGENT_MANAGER) && !hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to update disputes'
      };
    }
    
    // Get the existing dispute
    const existingDispute = getDisputeById(disputeId);
    if (!existingDispute) {
      return {
        success: false,
        message: `Dispute not found: ${disputeId}`
      };
    }
    
    // Check if the dispute is still pending
    if (existingDispute.Status !== STATUS.PENDING) {
      return {
        success: false,
        message: 'Cannot update a dispute that is already being reviewed or has been resolved'
      };
    }
    
    // Check if the current user is the original submitter or has admin permission
    const currentUser = Session.getActiveUser().getEmail();
    const isOriginalSubmitter = existingDispute['Submitted By'] === currentUser;
    const isAdmin = hasPermission(USER_ROLES.ADMIN);
    
    if (!isOriginalSubmitter && !isAdmin) {
      return {
        success: false,
        message: 'You can only update disputes that you submitted'
      };
    }
    
    // Prepare update data
    const fieldsToUpdate = {};
    
    // Only update allowed fields
    if (updateData.Reason !== undefined) {
      fieldsToUpdate.Reason = updateData.Reason;
    }
    
    if (updateData.Details !== undefined) {
      fieldsToUpdate.Details = updateData.Details;
    }
    
    if (updateData['Additional Evidence'] !== undefined) {
      fieldsToUpdate['Additional Evidence'] = updateData['Additional Evidence'];
    }
    
    if (updateData['Requested Score Change'] !== undefined) {
      fieldsToUpdate['Requested Score Change'] = updateData['Requested Score Change'];
    }
    
    // Update the dispute
    if (Object.keys(fieldsToUpdate).length > 0) {
      if (!updateRowInSheet(SHEET_NAMES.DISPUTES, disputeId, fieldsToUpdate)) {
        return {
          success: false,
          message: 'Failed to update dispute'
        };
      }
    }
    
    // Log the action
    logAction(currentUser, 'Update Dispute', `Updated dispute ${disputeId}`);
    
    return {
      success: true,
      message: 'Dispute updated successfully'
    };
  } catch (error) {
    Logger.log(`Error in updateDispute: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Review a dispute
 * 
 * @param {string} disputeId - The ID of the dispute to review
 * @param {Object} reviewData - The review data
 * @return {Object} Result object with success flag and message
 */
function reviewDispute(disputeId, reviewData) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.QA_MANAGER) && !hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to review disputes'
      };
    }
    
    // Get the existing dispute
    const existingDispute = getDisputeById(disputeId);
    if (!existingDispute) {
      return {
        success: false,
        message: `Dispute not found: ${disputeId}`
      };
    }
    
    // Check if the dispute is still pending or in progress
    if (existingDispute.Status !== STATUS.PENDING && existingDispute.Status !== STATUS.IN_PROGRESS) {
      return {
        success: false,
        message: 'This dispute has already been resolved'
      };
    }
    
    // Validate required fields
    if (!reviewData.Status || !reviewData['Review Notes']) {
      return {
        success: false,
        message: 'Required fields missing (Status or Review Notes)'
      };
    }
    
    // Validate status
    const validStatuses = [STATUS.APPROVED, STATUS.PARTIALLY_APPROVED, STATUS.REJECTED];
    if (!validStatuses.includes(reviewData.Status)) {
      return {
        success: false,
        message: `Invalid status. Must be one of: ${validStatuses.join(', ')}`
      };
    }
    
    // Get current user as reviewer if not specified
    const currentUser = Session.getActiveUser().getEmail();
    if (!reviewData['Reviewed By']) {
      reviewData['Reviewed By'] = currentUser;
    }
    
    // Calculate score adjustment
    let scoreAdjustment = 0;
    if (reviewData.Status === STATUS.APPROVED || reviewData.Status === STATUS.PARTIALLY_APPROVED) {
      if (reviewData['Score Adjustment'] !== undefined) {
        scoreAdjustment = parseInt(reviewData['Score Adjustment']) || 0;
      }
    }
    
    // Get the evaluation
    const evaluation = getEvaluationById(existingDispute['Evaluation ID']);
    if (!evaluation) {
      return {
        success: false,
        message: `Evaluation not found: ${existingDispute['Evaluation ID']}`
      };
    }
    
    // Prepare update data for dispute
    const now = new Date();
    const disputeUpdateData = {
      Status: reviewData.Status,
      'Reviewed By': reviewData['Reviewed By'],
      'Review Date': now,
      'Score Adjustment': scoreAdjustment
    };
    
    // Update the dispute
    if (!updateRowInSheet(SHEET_NAMES.DISPUTES, disputeId, disputeUpdateData)) {
      return {
        success: false,
        message: 'Failed to update dispute'
      };
    }
    
    // Create dispute resolution record
    const resolutionRecord = {
      ID: generateUniqueId(),
      'Dispute ID': disputeId,
      'Resolution Date': now,
      'Resolution By': reviewData['Reviewed By'],
      Decision: reviewData.Status,
      'Review Notes': reviewData['Review Notes']
    };
    
    // Save resolution
    if (!addRowToSheet(SHEET_NAMES.DISPUTE_RESOLUTIONS, resolutionRecord)) {
      Logger.log('Warning: Failed to save dispute resolution record');
    }
    
    // Update evaluation if approved
    if (reviewData.Status === STATUS.APPROVED || reviewData.Status === STATUS.PARTIALLY_APPROVED) {
      // Calculate new score
      const currentScore = parseInt(evaluation.Score);
      const newScore = currentScore + scoreAdjustment;
      
      const evalUpdateData = {
        Score: newScore,
        Status: STATUS.COMPLETED
      };
      
      if (!updateRowInSheet(SHEET_NAMES.EVALUATIONS, existingDispute['Evaluation ID'], evalUpdateData)) {
        Logger.log('Warning: Failed to update evaluation score');
      }
    } else {
      // If rejected, just update status back to completed
      updateRowInSheet(SHEET_NAMES.EVALUATIONS, existingDispute['Evaluation ID'], {
        Status: STATUS.COMPLETED
      });
    }
    
    // Log the action
    logAction(currentUser, 'Review Dispute', 
              `Reviewed dispute ${disputeId} with decision: ${reviewData.Status}`);
    
    // Send notifications
    const updatedDispute = { ...existingDispute, ...disputeUpdateData };
    
    // Build list of recipients
    const recipients = [];
    
    // Add agent
    recipients.push(evaluation.Agent);
    
    // Add submitter if different from agent (e.g., manager)
    if (existingDispute['Submitted By'] !== evaluation.Agent) {
      recipients.push(existingDispute['Submitted By']);
    }
    
    // Add evaluator
    recipients.push(evaluation.Evaluator);
    
    // Send resolution notification
    sendDisputeResolutionNotification(updatedDispute, evaluation, recipients);
    
    return {
      success: true,
      message: 'Dispute reviewed successfully'
    };
  } catch (error) {
    Logger.log(`Error in reviewDispute: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Cancel a dispute
 * 
 * @param {string} disputeId - The ID of the dispute to cancel
 * @return {Object} Result object with success flag and message
 */
function cancelDispute(disputeId) {
  try {
    // Get the existing dispute
    const existingDispute = getDisputeById(disputeId);
    if (!existingDispute) {
      return {
        success: false,
        message: `Dispute not found: ${disputeId}`
      };
    }
    
    // Check if the dispute is still pending
    if (existingDispute.Status !== STATUS.PENDING) {
      return {
        success: false,
        message: 'Cannot cancel a dispute that is already being reviewed or has been resolved'
      };
    }
    
    // Check if the current user is the original submitter or has higher permission
    const currentUser = Session.getActiveUser().getEmail();
    const isOriginalSubmitter = existingDispute['Submitted By'] === currentUser;
    const isAdmin = hasPermission(USER_ROLES.ADMIN);
    
    if (!isOriginalSubmitter && !isAdmin) {
      return {
        success: false,
        message: 'You can only cancel disputes that you submitted'
      };
    }
    
    // Delete the dispute
    if (!deleteRowFromSheet(SHEET_NAMES.DISPUTES, disputeId)) {
      return {
        success: false,
        message: 'Failed to cancel dispute'
      };
    }
    
    // Update the evaluation status back to completed
    updateRowInSheet(SHEET_NAMES.EVALUATIONS, existingDispute['Evaluation ID'], {
      Status: STATUS.COMPLETED
    });
    
    // Log the action
    logAction(currentUser, 'Cancel Dispute', `Cancelled dispute ${disputeId}`);
    
    return {
      success: true,
      message: 'Dispute cancelled successfully'
    };
  } catch (error) {
    Logger.log(`Error in cancelDispute: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Get dispute statistics
 * 
 * @param {Date} startDate - Start date for the statistics (optional)
 * @param {Date} endDate - End date for the statistics (optional)
 * @return {Object} Object with dispute statistics
 */
function getDisputeStatistics(startDate = null, endDate = null) {
  try {
    // Get all disputes
    let disputes = getAllDisputes();
    
    // Filter by date range if provided
    if (startDate) {
      const start = startDate instanceof Date ? startDate : new Date(startDate);
      disputes = disputes.filter(dispute => {
        const disputeDate = dispute['Submission Date'] instanceof Date ? 
          dispute['Submission Date'] : new Date(dispute['Submission Date']);
        return disputeDate >= start;
      });
    }
    
    if (endDate) {
      const end = endDate instanceof Date ? endDate : new Date(endDate);
      disputes = disputes.filter(dispute => {
        const disputeDate = dispute['Submission Date'] instanceof Date ? 
          dispute['Submission Date'] : new Date(dispute['Submission Date']);
        return disputeDate <= end;
      });
    }
    
    // Calculate stats
    const totalDisputes = disputes.length;
    let approvedCount = 0;
    let partiallyApprovedCount = 0;
    let rejectedCount = 0;
    let pendingCount = 0;
    
    // Reason counts
    const reasonCounts = {};
    
    // Submitter counts
    const submitterCounts = {};
    
    // Process disputes
    disputes.forEach(dispute => {
      // Count by status
      if (dispute.Status === STATUS.APPROVED) {
        approvedCount++;
      } else if (dispute.Status === STATUS.PARTIALLY_APPROVED) {
        partiallyApprovedCount++;
      } else if (dispute.Status === STATUS.REJECTED) {
        rejectedCount++;
      } else if (dispute.Status === STATUS.PENDING || dispute.Status === STATUS.IN_PROGRESS) {
        pendingCount++;
      }
      
      // Count by reason
      const reason = dispute.Reason || 'Unknown';
      reasonCounts[reason] = (reasonCounts[reason] || 0) + 1;
      
      // Count by submitter
      const submitter = dispute['Submitted By'] || 'Unknown';
      submitterCounts[submitter] = (submitterCounts[submitter] || 0) + 1;
    });
    
    // Calculate percentages
    const resolvedDisputes = approvedCount + partiallyApprovedCount + rejectedCount;
    const approvalRate = resolvedDisputes > 0 ? 
      ((approvedCount + partiallyApprovedCount) / resolvedDisputes) * 100 : 0;
    
    // Return statistics
    return {
      totalDisputes: totalDisputes,
      pendingCount: pendingCount,
      approvedCount: approvedCount,
      partiallyApprovedCount: partiallyApprovedCount,
      rejectedCount: rejectedCount,
      approvalRate: approvalRate.toFixed(1),
      reasonCounts: reasonCounts,
      submitterCounts: submitterCounts
    };
  } catch (error) {
    Logger.log(`Error in getDisputeStatistics: ${error.message}`);
    return {
      totalDisputes: 0,
      pendingCount: 0,
      approvedCount: 0,
      partiallyApprovedCount: 0,
      rejectedCount: 0,
      approvalRate: '0.0',
      reasonCounts: {},
      submitterCounts: {}
    };
  }
}