/**
 * Dashboard Functions
 * 
 * This file contains functions for the QA Platform dashboard,
 * including data aggregation and visualization.
 */

/**
 * Show the dashboard UI
 */
function showDashboard() {
  const html = HtmlService.createTemplateFromFile('UI/Dashboard')
    .evaluate()
    .setTitle('QA Dashboard')
    .setWidth(900)
    .setHeight(700);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'QA Dashboard');
}

/**
 * Get dashboard data
 * 
 * @return {Object} Dashboard data object
 */
function getDashboardData() {
  // Get user's role to determine what data to show
  const userRole = getUserRole(currentUser);
  
  // Basic dashboard data for all roles
  const dashboardData = {
    queueSummary: getQueueSummary(),
    roleAccess: userRole,
    currentDate: new Date().toISOString()
  };
  
  // Additional data based on role
  if ([USER_ROLES.QA_MANAGER, USER_ROLES.ADMIN].includes(userRole)) {
    // Add QA Manager specific data
    dashboardData.qaManagerData = {
      pendingDisputes: getPendingDisputeCount(),
      evaluationSummary: getEvaluationSummary(),
      monthlyTrends: getMonthlyEvaluationTrends()
    };
  }
  
  if ([USER_ROLES.AGENT_MANAGER, USER_ROLES.ADMIN].includes(userRole)) {
    // Add Agent Manager specific data
    dashboardData.agentManagerData = {
      agentPerformance: getAgentPerformanceSummary(),
      disputeSummary: getDisputeSummary()
    };
  }
  
  if (userRole === USER_ROLES.QA_ANALYST) {
    // Add QA Analyst specific data
    dashboardData.qaAnalystData = {
      myEvaluations: getMyEvaluationCount(),
      pendingQueue: getPendingAuditsCount()
    };
  }
  
  return dashboardData;
}

/**
 * Get audit queue summary
 * 
 * @return {Object} Queue summary data
 */
function getQueueSummary() {
  const audits = getDataFromSheet(SHEET_NAMES.AUDIT_QUEUE);
  
  // Count by status
  const pending = audits.filter(a => a['Queue Status'] === 'Pending').length;
  const assigned = audits.filter(a => a['Queue Status'] === 'Assigned').length;
  const evaluated = audits.filter(a => a['Queue Status'] === 'Evaluated').length;
  
  // Count by priority
  const highPriority = audits.filter(a => a['Priority'] === 'High').length;
  const mediumPriority = audits.filter(a => a['Priority'] === 'Medium').length;
  const lowPriority = audits.filter(a => a['Priority'] === 'Low').length;
  
  // Calculate age of oldest pending audit
  let oldestPendingDays = 0;
  const pendingAudits = audits.filter(a => a['Queue Status'] === 'Pending');
  
  if (pendingAudits.length > 0) {
    const dates = pendingAudits.map(a => new Date(a['Import Date']));
    const oldestDate = new Date(Math.min.apply(null, dates));
    const daysDiff = Math.floor((new Date() - oldestDate) / (1000 * 60 * 60 * 24));
    oldestPendingDays = daysDiff;
  }
  
  return {
    total: audits.length,
    pending: pending,
    assigned: assigned,
    evaluated: evaluated,
    highPriority: highPriority,
    mediumPriority: mediumPriority,
    lowPriority: lowPriority,
    oldestPendingDays: oldestPendingDays
  };
}

/**
 * Get number of pending audits
 * 
 * @return {number} Count of pending audits
 */
function getPendingAuditsCount() {
  const audits = getDataFromSheet(SHEET_NAMES.AUDIT_QUEUE);
  return audits.filter(a => a['Queue Status'] === 'Pending').length;
}

/**
 * Get number of pending disputes
 * 
 * @return {number} Count of pending disputes
 */
function getPendingDisputeCount() {
  const disputes = getDataFromSheet(SHEET_NAMES.DISPUTES);
  return disputes.filter(d => d.Status === 'Pending').length;
}

/**
 * Get evaluation summary
 * 
 * @return {Object} Evaluation summary data
 */
function getEvaluationSummary() {
  const evaluations = getDataFromSheet(SHEET_NAMES.EVALUATIONS);
  
  if (evaluations.length === 0) {
    return {
      count: 0,
      averageScore: 0,
      passingRate: 0,
      disputeRate: 0
    };
  }
  
  // Calculate average score
  let totalScore = 0;
  let totalPossible = 0;
  let passingCount = 0;
  const passingThreshold = parseInt(getSetting('passing_score_percentage', '80'));
  
  for (const eval of evaluations) {
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
  const disputedCount = evaluations.filter(e => e.Disputed === true).length;
  
  return {
    count: evaluations.length,
    averageScore: averagePercentage.toFixed(2),
    passingRate: ((passingCount / evaluations.length) * 100).toFixed(2),
    disputeRate: ((disputedCount / evaluations.length) * 100).toFixed(2)
  };
}

/**
 * Get monthly evaluation trends
 * 
 * @return {Object} Monthly trend data for the last 6 months
 */
function getMonthlyEvaluationTrends() {
  const evaluations = getDataFromSheet(SHEET_NAMES.EVALUATIONS);
  const disputes = getDataFromSheet(SHEET_NAMES.DISPUTES);
  
  // Get last 6 months
  const months = [];
  const today = new Date();
  
  for (let i = 5; i >= 0; i--) {
    const d = new Date(today.getFullYear(), today.getMonth() - i, 1);
    months.push({
      month: d.toLocaleString('default', { month: 'short' }),
      year: d.getFullYear(),
      evaluationCount: 0,
      avgScore: 0,
      disputeCount: 0
    });
  }
  
  // Group evaluations by month
  for (const eval of evaluations) {
    const evalDate = new Date(eval['Evaluation Date']);
    const monthYear = evalDate.toLocaleString('default', { month: 'short' }) + 
                      evalDate.getFullYear();
    
    for (const month of months) {
      if (month.month + month.year === monthYear) {
        month.evaluationCount++;
        
        // Add score to calculate average later
        if (!month.totalScore) month.totalScore = 0;
        if (!month.totalPossible) month.totalPossible = 0;
        
        month.totalScore += parseInt(eval.Score);
        month.totalPossible += parseInt(eval['Max Possible']);
      }
    }
  }
  
  // Group disputes by month
  for (const dispute of disputes) {
    const disputeDate = new Date(dispute['Date Filed']);
    const monthYear = disputeDate.toLocaleString('default', { month: 'short' }) + 
                      disputeDate.getFullYear();
    
    for (const month of months) {
      if (month.month + month.year === monthYear) {
        month.disputeCount++;
      }
    }
  }
  
  // Calculate average scores
  for (const month of months) {
    if (month.evaluationCount > 0) {
      month.avgScore = ((month.totalScore / month.totalPossible) * 100).toFixed(2);
    }
    // Remove calculation fields
    delete month.totalScore;
    delete month.totalPossible;
  }
  
  return months;
}

/**
 * Get agent performance summary
 * 
 * @return {Array} Agent performance data
 */
function getAgentPerformanceSummary() {
  const evaluations = getDataFromSheet(SHEET_NAMES.EVALUATIONS);
  const users = getAllUsers();
  
  // Get agents (filter users with non-manager roles)
  const agents = users.filter(user => 
    user.Role !== USER_ROLES.QA_MANAGER && 
    user.Role !== USER_ROLES.ADMIN &&
    user.Active === true
  );
  
  // Current month range
  const today = new Date();
  const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  const endOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  
  // Calculate performance for each agent
  const agentPerformance = [];
  
  for (const agent of agents) {
    const stats = getAgentEvaluationStats(agent.Email, startOfMonth, endOfMonth);
    
    agentPerformance.push({
      name: agent.Name,
      email: agent.Email,
      department: agent.Department,
      evaluationCount: stats.count,
      averageScore: stats.averageScore,
      passingRate: stats.passingRate
    });
  }
  
  return agentPerformance;
}

/**
 * Get dispute summary
 * 
 * @return {Object} Dispute summary data
 */
function getDisputeSummary() {
  const disputes = getDataFromSheet(SHEET_NAMES.DISPUTES);
  
  if (disputes.length === 0) {
    return {
      total: 0,
      pending: 0,
      approved: 0,
      denied: 0,
      approvalRate: 0
    };
  }
  
  const pending = disputes.filter(d => d.Status === 'Pending').length;
  const approved = disputes.filter(d => d.Status === 'Approved').length;
  const denied = disputes.filter(d => d.Status === 'Denied').length;
  
  // Calculate approval rate (excluding pending)
  const resolved = approved + denied;
  const approvalRate = resolved > 0 ? ((approved / resolved) * 100).toFixed(2) : 0;
  
  return {
    total: disputes.length,
    pending: pending,
    approved: approved,
    denied: denied,
    approvalRate: approvalRate
  };
}

/**
 * Get count of evaluations done by current user
 * 
 * @return {number} Count of evaluations
 */
function getMyEvaluationCount() {
  const evaluations = getDataFromSheet(SHEET_NAMES.EVALUATIONS);
  return evaluations.filter(e => e.Evaluator === currentUser).length;
}