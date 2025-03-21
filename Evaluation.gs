/**
 * Evaluation Functions
 * 
 * This file contains functions for creating, managing, and analyzing
 * quality evaluations for agents.
 */

/**
 * Show the evaluation form UI
 */
function showEvaluationForm() {
  if (!hasPermission(USER_ROLES.QA_ANALYST)) {
    SpreadsheetApp.getUi().alert('You do not have permission to perform evaluations');
    return;
  }
  
  const html = HtmlService.createTemplateFromFile('UI/EvaluationForm')
    .evaluate()
    .setTitle('New Evaluation')
    .setWidth(800)
    .setHeight(700);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'New Evaluation');
}

/**
 * Get available audit items for evaluation
 * 
 * @return {Array} Array of audit items that need evaluation
 */
function getAvailableAudits() {
  if (!hasPermission(USER_ROLES.QA_ANALYST)) {
    throw new Error('You do not have permission to view audits');
  }
  
  const audits = getDataFromSheet(SHEET_NAMES.AUDIT_QUEUE);
  
  // Filter to audits that are not yet assigned or assigned to the current user
  const availableAudits = audits.filter(audit => 
    audit['Queue Status'] === 'Pending' || 
    audit['Assigned To'] === currentUser
  );
  
  return availableAudits;
}

/**
 * Get evaluation templates
 * 
 * @return {Array} Array of evaluation templates
 */
function getEvaluationTemplates() {
  return getDataFromSheet(SHEET_NAMES.TEMPLATES).filter(template => 
    template.Active === true
  );
}

/**
 * Get questions for a specific template
 * 
 * @param {string} templateId - Template ID
 * @return {Array} Array of question objects
 */
function getQuestionsForTemplate(templateId) {
  const questions = getDataFromSheet(SHEET_NAMES.QUESTIONS);
  
  return questions.filter(question => 
    question['Template ID'] === templateId && 
    question.Active === true
  );
}

/**
 * Save a new evaluation
 * 
 * @param {string} auditId - ID of the audit being evaluated
 * @param {string} agentEmail - Email of the agent being evaluated
 * @param {Array} responses - Array of question responses: {questionId, score, comment}
 * @param {string} overallComment - Overall comment for the evaluation
 * @return {Object} Newly created evaluation data
 */
function saveEvaluation(auditId, agentEmail, responses, overallComment) {
  if (!hasPermission(USER_ROLES.QA_ANALYST)) {
    throw new Error('You do not have permission to create evaluations');
  }
  
  // Get the audit details
  const audit = findRowById(SHEET_NAMES.AUDIT_QUEUE, auditId);
  if (!audit) {
    throw new Error(`Audit with ID ${auditId} not found`);
  }
  
  // Calculate scores
  let totalScore = 0;
  let maxPossible = 0;
  
  // Get all questions to calculate weights properly
  const questionResponses = [];
  for (const response of responses) {
    const question = findRowById(SHEET_NAMES.QUESTIONS, response.questionId);
    if (!question) {
      throw new Error(`Question with ID ${response.questionId} not found`);
    }
    
    const weight = parseInt(question.Weight);
    const score = parseInt(response.score);
    
    totalScore += score * weight;
    maxPossible += 100 * weight; // Assuming max score per question is 100
    
    questionResponses.push({
      questionId: response.questionId,
      question: question.Question,
      category: question.Category,
      weight: weight,
      score: score,
      comment: response.comment || '',
      critical: question.Critical === 'TRUE' || question.Critical === true
    });
  }
  
  // Create the evaluation record
  const evaluationId = generateUniqueId();
  const evaluation = {
    ID: evaluationId,
    'Audit ID': auditId,
    'Evaluation Date': new Date(),
    Evaluator: currentUser,
    Agent: agentEmail,
    Score: totalScore,
    'Max Possible': maxPossible,
    Status: 'Completed',
    Disputed: false,
    'Last Updated': new Date()
  };
  
  const rowIndex = addRowToSheet(SHEET_NAMES.EVALUATIONS, evaluation);
  
  // Update the audit queue status
  updateRowInSheet(SHEET_NAMES.AUDIT_QUEUE, audit.rowIndex, {
    'Queue Status': 'Evaluated',
    'Assigned To': currentUser
  });
  
  // Store the detailed responses in the properties service
  // (since Google Sheets isn't ideal for nested data)
  const evaluationDetail = {
    evaluationId: evaluationId,
    responses: questionResponses,
    overallComment: overallComment,
    evaluationDate: new Date().toISOString()
  };
  
  const userProperties = PropertiesService.getScriptProperties();
  userProperties.setProperty(`evaluation_${evaluationId}`, JSON.stringify(evaluationDetail));
  
  // Send notifications
  const agent = getUserByEmail(agentEmail);
  const evaluator = getUserByEmail(currentUser);
  
  if (agent) {
    // Find the agent's manager (assuming the first agent manager we find)
    const managers = getAllAgentManagers();
    if (managers.length > 0) {
      notifyManagerAboutEvaluation(evaluation, agent, managers[0]);
    }
    
    // Notify the agent
    notifyAgentAboutEvaluation(evaluation, agent, evaluator);
  }
  
  return {
    evaluation: evaluation,
    detail: evaluationDetail
  };
}

/**
 * Get an evaluation by ID with details
 * 
 * @param {string} evaluationId - ID of the evaluation
 * @return {Object} Evaluation with details
 */
function getEvaluationWithDetails(evaluationId) {
  // Get the basic evaluation
  const evaluation = findRowById(SHEET_NAMES.EVALUATIONS, evaluationId);
  if (!evaluation) {
    throw new Error(`Evaluation with ID ${evaluationId} not found`);
  }
  
  // Get the detailed responses from properties
  const userProperties = PropertiesService.getScriptProperties();
  const detailJson = userProperties.getProperty(`evaluation_${evaluationId}`);
  
  if (!detailJson) {
    throw new Error(`Detailed responses for evaluation ${evaluationId} not found`);
  }
  
  const detail = JSON.parse(detailJson);
  
  // Check if this evaluation has any disputes
  const disputes = findRowsInSheet(SHEET_NAMES.DISPUTES, { 'Evaluation ID': evaluationId });
  
  return {
    evaluation: evaluation,
    detail: detail,
    hasDispute: disputes.length > 0,
    disputes: disputes
  };
}

/**
 * Get evaluations for the current user
 * 
 * @param {string} type - 'evaluated' for evaluations created by current user, 
 *                       'received' for evaluations received by current user
 * @return {Array} Array of evaluations
 */
function getMyEvaluations(type) {
  const evaluations = getDataFromSheet(SHEET_NAMES.EVALUATIONS);
  
  if (type === 'evaluated') {
    return evaluations.filter(eval => eval.Evaluator === currentUser);
  } else if (type === 'received') {
    return evaluations.filter(eval => eval.Agent === currentUser);
  } else {
    return evaluations;
  }
}

/**
 * Show the 'My Evaluations' UI
 */
function showMyEvaluations() {
  const html = HtmlService.createTemplateFromFile('UI/MyEvaluations')
    .evaluate()
    .setTitle('My Evaluations')
    .setWidth(800)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'My Evaluations');
}

/**
 * Calculate evaluation statistics for a given agent
 * 
 * @param {string} agentEmail - Email of the agent
 * @param {Date} startDate - Start date for the period
 * @param {Date} endDate - End date for the period
 * @return {Object} Statistics object
 */
function getAgentEvaluationStats(agentEmail, startDate, endDate) {
  const evaluations = getDataFromSheet(SHEET_NAMES.EVALUATIONS);
  
  // Filter evaluations by agent and date range
  const agentEvals = evaluations.filter(eval => {
    const evalDate = new Date(eval['Evaluation Date']);
    return eval.Agent === agentEmail && 
           evalDate >= startDate && 
           evalDate <= endDate;
  });
  
  if (agentEvals.length === 0) {
    return {
      count: 0,
      averageScore: 0,
      passingRate: 0,
      evaluations: []
    };
  }
  
  // Calculate statistics
  let totalScore = 0;
  let totalPossible = 0;
  let passingCount = 0;
  const passingThreshold = parseInt(getSetting('passing_score_percentage', '80'));
  
  for (const eval of agentEvals) {
    const score = parseInt(eval.Score);
    const maxPossible = parseInt(eval['Max Possible']);
    const percentage = (score / maxPossible) * 100;
    
    totalScore += score;
    totalPossible += maxPossible;
    
    if (percentage >= passingThreshold) {
      passingCount++;
    }
  }
  
  const averagePercentage = (totalScore / totalPossible) * 100;
  
  return {
    count: agentEvals.length,
    averageScore: averagePercentage.toFixed(2),
    passingRate: ((passingCount / agentEvals.length) * 100).toFixed(2),
    evaluations: agentEvals
  };
}