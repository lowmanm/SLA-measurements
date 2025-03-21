/**
 * Evaluation.gs
 * 
 * This file contains functions for creating and managing evaluations.
 */

/**
 * Get all evaluations
 * 
 * @return {Array} Array of evaluation objects
 */
function getAllEvaluations() {
  return getDataFromSheet(SHEET_NAMES.EVALUATIONS);
}

/**
 * Get evaluations based on filter criteria
 * 
 * @param {Object} filters - Object with filter criteria (e.g., {Agent: 'email@example.com'})
 * @return {Array} Array of filtered evaluation objects
 */
function getFilteredEvaluations(filters) {
  return getFilteredData(SHEET_NAMES.EVALUATIONS, filters);
}

/**
 * Get an evaluation by ID
 * 
 * @param {string} evaluationId - The evaluation ID
 * @return {Object} Evaluation object or null if not found
 */
function getEvaluationById(evaluationId) {
  return getRowById(SHEET_NAMES.EVALUATIONS, evaluationId);
}

/**
 * Get evaluations for a specific agent
 * 
 * @param {string} agentEmail - The agent's email
 * @return {Array} Array of evaluation objects for the agent
 */
function getEvaluationsForAgent(agentEmail) {
  return getFilteredData(SHEET_NAMES.EVALUATIONS, { Agent: agentEmail });
}

/**
 * Get evaluations by an evaluator
 * 
 * @param {string} evaluatorEmail - The evaluator's email
 * @return {Array} Array of evaluation objects by the evaluator
 */
function getEvaluationsByEvaluator(evaluatorEmail) {
  return getFilteredData(SHEET_NAMES.EVALUATIONS, { Evaluator: evaluatorEmail });
}

/**
 * Get evaluation answers for a specific evaluation
 * 
 * @param {string} evaluationId - The evaluation ID
 * @return {Array} Array of evaluation answer objects
 */
function getEvaluationAnswers(evaluationId) {
  return getFilteredData(SHEET_NAMES.EVALUATION_ANSWERS, { 'Evaluation ID': evaluationId });
}

/**
 * Get question sets suitable for a specific interaction type
 * 
 * @param {string} interactionType - The interaction type
 * @return {Array} Array of question set objects
 */
function getQuestionSetsForInteractionType(interactionType) {
  return getFilteredData(SHEET_NAMES.QUESTION_SETS, { 'Interaction Type': interactionType });
}

/**
 * Get questions for a specific question set
 * 
 * @param {string} questionSetId - The question set ID
 * @return {Array} Array of question objects
 */
function getQuestionsForQuestionSet(questionSetId) {
  return getFilteredData(SHEET_NAMES.QUESTIONS, { 'Question Set ID': questionSetId });
}

/**
 * Create a new evaluation
 * 
 * @param {Object} evaluationData - Basic evaluation data
 * @param {Array} answerData - Array of answer data objects
 * @return {Object} Result object with success flag, message, and evaluation ID
 */
function createEvaluation(evaluationData, answerData) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.QA_ANALYST) && 
        !hasPermission(USER_ROLES.QA_MANAGER) && 
        !hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to create evaluations'
      };
    }
    
    // Validate required fields
    if (!evaluationData.Agent || !evaluationData['Question Set ID'] || !evaluationData['Interaction Type']) {
      return {
        success: false,
        message: 'Required fields missing (Agent, Question Set ID, or Interaction Type)'
      };
    }
    
    // Verify the agent exists
    const agent = getUserByEmail(evaluationData.Agent);
    if (!agent) {
      return {
        success: false,
        message: `Agent not found: ${evaluationData.Agent}`
      };
    }
    
    // Verify the question set exists
    const questionSet = getRowById(SHEET_NAMES.QUESTION_SETS, evaluationData['Question Set ID']);
    if (!questionSet) {
      return {
        success: false,
        message: `Question set not found: ${evaluationData['Question Set ID']}`
      };
    }
    
    // Get current user as evaluator if not specified
    const currentUser = Session.getActiveUser().getEmail();
    if (!evaluationData.Evaluator) {
      evaluationData.Evaluator = currentUser;
    }
    
    // Generate ID
    const evaluationId = generateUniqueId();
    
    // Calculate total score and max possible
    let totalScore = 0;
    let maxPossible = 0;
    
    // Validate answerData
    if (!answerData || !Array.isArray(answerData) || answerData.length === 0) {
      return {
        success: false,
        message: 'No answers provided'
      };
    }
    
    // Process each answer
    answerData.forEach(answer => {
      totalScore += parseInt(answer.Score) || 0;
      maxPossible += parseInt(answer['Max Score']) || 0;
    });
    
    // Prepare evaluation record
    const now = new Date();
    const evaluationRecord = {
      ID: evaluationId,
      Date: now,
      Agent: evaluationData.Agent,
      Evaluator: evaluationData.Evaluator,
      'Question Set ID': evaluationData['Question Set ID'],
      'Interaction Type': evaluationData['Interaction Type'],
      'Customer ID': evaluationData['Customer ID'] || '',
      'Interaction ID': evaluationData['Interaction ID'] || '',
      Score: totalScore,
      'Max Possible': maxPossible,
      Status: STATUS.COMPLETED,
      Strengths: evaluationData.Strengths || '',
      'Areas for Improvement': evaluationData['Areas for Improvement'] || '',
      Comments: evaluationData.Comments || ''
    };
    
    // Save evaluation
    if (!addRowToSheet(SHEET_NAMES.EVALUATIONS, evaluationRecord)) {
      return {
        success: false,
        message: 'Failed to save evaluation'
      };
    }
    
    // Save answers
    let allAnswersSaved = true;
    answerData.forEach(answer => {
      const answerRecord = {
        ID: generateUniqueId(),
        'Evaluation ID': evaluationId,
        'Question ID': answer['Question ID'] || '',
        Question: answer.Question,
        Answer: answer.Answer,
        Score: answer.Score,
        'Max Score': answer['Max Score'],
        Comments: answer.Comments || ''
      };
      
      if (!addRowToSheet(SHEET_NAMES.EVALUATION_ANSWERS, answerRecord)) {
        allAnswersSaved = false;
      }
    });
    
    if (!allAnswersSaved) {
      Logger.log('Some answers failed to save');
    }
    
    // Log the action
    logAction(currentUser, 'Create Evaluation', 
              `Created evaluation for ${evaluationData.Agent} with score ${totalScore}/${maxPossible}`);
    
    // Check if email notifications should be sent
    const sendEmails = evaluationData.sendNotifications !== false;
    
    // Send notifications if requested
    if (sendEmails) {
      // Send to agent
      sendEvaluationNotificationToAgent(evaluationRecord);
      
      // Send to manager
      sendEvaluationNotificationToManager(evaluationRecord);
    }
    
    return {
      success: true,
      message: 'Evaluation created successfully',
      evaluationId: evaluationId
    };
  } catch (error) {
    Logger.log(`Error in createEvaluation: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Update an existing evaluation
 * 
 * @param {string} evaluationId - The ID of the evaluation to update
 * @param {Object} evaluationData - The evaluation data to update
 * @param {Array} answerData - Array of answer data objects (optional)
 * @return {Object} Result object with success flag and message
 */
function updateEvaluation(evaluationId, evaluationData, answerData) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.QA_ANALYST) && 
        !hasPermission(USER_ROLES.QA_MANAGER) && 
        !hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to update evaluations'
      };
    }
    
    // Get the existing evaluation
    const existingEvaluation = getEvaluationById(evaluationId);
    if (!existingEvaluation) {
      return {
        success: false,
        message: `Evaluation not found: ${evaluationId}`
      };
    }
    
    // Check if the evaluation is disputed
    if (existingEvaluation.Status === STATUS.DISPUTED) {
      return {
        success: false,
        message: 'Cannot update a disputed evaluation. Resolve the dispute first.'
      };
    }
    
    // Check if the current user is the original evaluator or has higher permission
    const currentUser = Session.getActiveUser().getEmail();
    const isOriginalEvaluator = existingEvaluation.Evaluator === currentUser;
    const hasHigherPermission = hasPermission(USER_ROLES.QA_MANAGER) || hasPermission(USER_ROLES.ADMIN);
    
    if (!isOriginalEvaluator && !hasHigherPermission) {
      return {
        success: false,
        message: 'You can only update evaluations that you created'
      };
    }
    
    // Prepare update data
    const updateData = {};
    
    // Only update the fields provided
    if (evaluationData.Strengths !== undefined) {
      updateData.Strengths = evaluationData.Strengths;
    }
    
    if (evaluationData['Areas for Improvement'] !== undefined) {
      updateData['Areas for Improvement'] = evaluationData['Areas for Improvement'];
    }
    
    if (evaluationData.Comments !== undefined) {
      updateData.Comments = evaluationData.Comments;
    }
    
    if (evaluationData.Status !== undefined) {
      updateData.Status = evaluationData.Status;
    }
    
    // Update answers if provided
    if (answerData && Array.isArray(answerData) && answerData.length > 0) {
      // Get existing answers
      const existingAnswers = getEvaluationAnswers(evaluationId);
      
      // Calculate new score
      let totalScore = 0;
      let maxPossible = 0;
      
      // Process each answer
      answerData.forEach(answer => {
        // Find existing answer if any
        const existingAnswer = existingAnswers.find(a => a['Question ID'] === answer['Question ID']);
        
        if (existingAnswer) {
          // Update existing answer
          updateRowInSheet(SHEET_NAMES.EVALUATION_ANSWERS, existingAnswer.ID, {
            Answer: answer.Answer,
            Score: answer.Score,
            Comments: answer.Comments || ''
          });
        } else {
          // Add new answer
          const answerRecord = {
            ID: generateUniqueId(),
            'Evaluation ID': evaluationId,
            'Question ID': answer['Question ID'] || '',
            Question: answer.Question,
            Answer: answer.Answer,
            Score: answer.Score,
            'Max Score': answer['Max Score'],
            Comments: answer.Comments || ''
          };
          
          addRowToSheet(SHEET_NAMES.EVALUATION_ANSWERS, answerRecord);
        }
        
        // Add to totals
        totalScore += parseInt(answer.Score) || 0;
        maxPossible += parseInt(answer['Max Score']) || 0;
      });
      
      // Update score in evaluation
      updateData.Score = totalScore;
      updateData['Max Possible'] = maxPossible;
    }
    
    // Update the evaluation
    if (Object.keys(updateData).length > 0) {
      if (!updateRowInSheet(SHEET_NAMES.EVALUATIONS, evaluationId, updateData)) {
        return {
          success: false,
          message: 'Failed to update evaluation'
        };
      }
    }
    
    // Log the action
    logAction(currentUser, 'Update Evaluation', `Updated evaluation ${evaluationId} for ${existingEvaluation.Agent}`);
    
    return {
      success: true,
      message: 'Evaluation updated successfully'
    };
  } catch (error) {
    Logger.log(`Error in updateEvaluation: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Delete an evaluation
 * 
 * @param {string} evaluationId - The ID of the evaluation to delete
 * @return {Object} Result object with success flag and message
 */
function deleteEvaluation(evaluationId) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.QA_MANAGER) && 
        !hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to delete evaluations'
      };
    }
    
    // Get the existing evaluation
    const existingEvaluation = getEvaluationById(evaluationId);
    if (!existingEvaluation) {
      return {
        success: false,
        message: `Evaluation not found: ${evaluationId}`
      };
    }
    
    // Get associated answers
    const answers = getEvaluationAnswers(evaluationId);
    
    // Delete all answers first
    answers.forEach(answer => {
      deleteRowFromSheet(SHEET_NAMES.EVALUATION_ANSWERS, answer.ID);
    });
    
    // Delete disputes if any
    const disputes = getFilteredData(SHEET_NAMES.DISPUTES, { 'Evaluation ID': evaluationId });
    disputes.forEach(dispute => {
      // Delete dispute resolutions
      const resolutions = getFilteredData(SHEET_NAMES.DISPUTE_RESOLUTIONS, { 'Dispute ID': dispute.ID });
      resolutions.forEach(resolution => {
        deleteRowFromSheet(SHEET_NAMES.DISPUTE_RESOLUTIONS, resolution.ID);
      });
      
      // Delete the dispute
      deleteRowFromSheet(SHEET_NAMES.DISPUTES, dispute.ID);
    });
    
    // Delete the evaluation
    if (!deleteRowFromSheet(SHEET_NAMES.EVALUATIONS, evaluationId)) {
      return {
        success: false,
        message: 'Failed to delete evaluation'
      };
    }
    
    // Log the action
    const currentUser = Session.getActiveUser().getEmail();
    logAction(currentUser, 'Delete Evaluation', 
              `Deleted evaluation ${evaluationId} for ${existingEvaluation.Agent}`);
    
    return {
      success: true,
      message: 'Evaluation and related data deleted successfully'
    };
  } catch (error) {
    Logger.log(`Error in deleteEvaluation: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Get recent evaluations for a dashboard
 * 
 * @param {number} limit - Maximum number of evaluations to return
 * @param {string} userEmail - Email of the user to filter evaluations for (optional)
 * @return {Array} Array of recent evaluation objects
 */
function getRecentEvaluations(limit = 10, userEmail = '') {
  try {
    // Get all evaluations
    const evaluations = getAllEvaluations();
    
    // Filter by user if provided
    let filteredEvals = evaluations;
    if (userEmail) {
      // Check user role
      const userInfo = getUserInfo();
      
      if (userInfo && userInfo.role) {
        if (userInfo.role === USER_ROLES.AGENT) {
          // Agents can only see their own evaluations
          filteredEvals = evaluations.filter(eval => eval.Agent === userEmail);
        } else if (userInfo.role === USER_ROLES.AGENT_MANAGER) {
          // Managers can see evaluations for their reports
          const reports = getUsersByManager(userEmail);
          const reportEmails = reports.map(user => user.Email);
          filteredEvals = evaluations.filter(eval => reportEmails.includes(eval.Agent));
        } else if (userInfo.role === USER_ROLES.QA_ANALYST) {
          // QA Analysts can see evaluations they created
          filteredEvals = evaluations.filter(eval => eval.Evaluator === userEmail);
        }
        // QA Managers and Admins can see all evaluations, which happens by default
      }
    }
    
    // Sort by date in descending order
    filteredEvals.sort((a, b) => {
      const dateA = a.Date instanceof Date ? a.Date : new Date(a.Date);
      const dateB = b.Date instanceof Date ? b.Date : new Date(b.Date);
      return dateB - dateA;
    });
    
    // Return limited number of evaluations
    return filteredEvals.slice(0, limit);
  } catch (error) {
    Logger.log(`Error in getRecentEvaluations: ${error.message}`);
    return [];
  }
}

/**
 * Get evaluation statistics for a dashboard
 * 
 * @param {string} userEmail - Email of the user to filter evaluations for (optional)
 * @param {Date} startDate - Start date for the statistics (optional)
 * @param {Date} endDate - End date for the statistics (optional)
 * @return {Object} Object with various statistics
 */
function getEvaluationStatistics(userEmail = '', startDate = null, endDate = null) {
  try {
    // Get all evaluations
    let evaluations = getAllEvaluations();
    
    // Filter by user if provided
    if (userEmail) {
      // Check user role
      const userInfo = getUserInfo();
      
      if (userInfo && userInfo.role) {
        if (userInfo.role === USER_ROLES.AGENT) {
          // Agents can only see their own evaluations
          evaluations = evaluations.filter(eval => eval.Agent === userEmail);
        } else if (userInfo.role === USER_ROLES.AGENT_MANAGER) {
          // Managers can see evaluations for their reports
          const reports = getUsersByManager(userEmail);
          const reportEmails = reports.map(user => user.Email);
          evaluations = evaluations.filter(eval => reportEmails.includes(eval.Agent));
        } else if (userInfo.role === USER_ROLES.QA_ANALYST) {
          // QA Analysts can see evaluations they created
          evaluations = evaluations.filter(eval => eval.Evaluator === userEmail);
        }
        // QA Managers and Admins can see all evaluations, which happens by default
      }
    }
    
    // Filter by date range if provided
    if (startDate) {
      const start = startDate instanceof Date ? startDate : new Date(startDate);
      evaluations = evaluations.filter(eval => {
        const evalDate = eval.Date instanceof Date ? eval.Date : new Date(eval.Date);
        return evalDate >= start;
      });
    }
    
    if (endDate) {
      const end = endDate instanceof Date ? endDate : new Date(endDate);
      evaluations = evaluations.filter(eval => {
        const evalDate = eval.Date instanceof Date ? eval.Date : new Date(eval.Date);
        return evalDate <= end;
      });
    }
    
    // Calculate stats
    const passingThreshold = parseInt(getSetting('passing_score_percentage', '80'));
    
    let totalEvaluations = evaluations.length;
    let totalScore = 0;
    let totalMaxPossible = 0;
    let totalPassed = 0;
    let totalDisputed = 0;
    
    // Interaction type counts
    const interactionTypeCounts = {};
    
    // Evaluator counts
    const evaluatorCounts = {};
    
    // Agent scores
    const agentScores = {};
    
    // Calculate statistics
    evaluations.forEach(eval => {
      // Calculate scores
      const score = parseInt(eval.Score);
      const maxPossible = parseInt(eval['Max Possible']);
      const percentage = (score / maxPossible) * 100;
      
      totalScore += score;
      totalMaxPossible += maxPossible;
      
      if (percentage >= passingThreshold) {
        totalPassed++;
      }
      
      if (eval.Status === STATUS.DISPUTED) {
        totalDisputed++;
      }
      
      // Count interaction types
      const interactionType = eval['Interaction Type'] || 'Unknown';
      interactionTypeCounts[interactionType] = (interactionTypeCounts[interactionType] || 0) + 1;
      
      // Count evaluators
      const evaluator = eval.Evaluator || 'Unknown';
      evaluatorCounts[evaluator] = (evaluatorCounts[evaluator] || 0) + 1;
      
      // Aggregate agent scores
      const agent = eval.Agent || 'Unknown';
      if (!agentScores[agent]) {
        agentScores[agent] = {
          totalScore: 0,
          totalMaxPossible: 0,
          count: 0
        };
      }
      
      agentScores[agent].totalScore += score;
      agentScores[agent].totalMaxPossible += maxPossible;
      agentScores[agent].count++;
    });
    
    // Calculate average score percentage
    const overallAveragePercentage = totalMaxPossible > 0 ? 
      (totalScore / totalMaxPossible) * 100 : 0;
    
    // Calculate agent average scores
    const agentAverages = {};
    for (const agent in agentScores) {
      const data = agentScores[agent];
      if (data.totalMaxPossible > 0) {
        agentAverages[agent] = (data.totalScore / data.totalMaxPossible) * 100;
      }
    }
    
    // Return compiled statistics
    return {
      totalEvaluations: totalEvaluations,
      overallAveragePercentage: overallAveragePercentage.toFixed(1),
      passRate: totalEvaluations > 0 ? ((totalPassed / totalEvaluations) * 100).toFixed(1) : '0.0',
      disputeRate: totalEvaluations > 0 ? ((totalDisputed / totalEvaluations) * 100).toFixed(1) : '0.0',
      interactionTypeCounts: interactionTypeCounts,
      evaluatorCounts: evaluatorCounts,
      agentAverages: agentAverages
    };
  } catch (error) {
    Logger.log(`Error in getEvaluationStatistics: ${error.message}`);
    return {
      totalEvaluations: 0,
      overallAveragePercentage: '0.0',
      passRate: '0.0',
      disputeRate: '0.0',
      interactionTypeCounts: {},
      evaluatorCounts: {},
      agentAverages: {}
    };
  }
}