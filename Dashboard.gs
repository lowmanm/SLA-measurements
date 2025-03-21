/**
 * Dashboard.gs
 * 
 * This file contains functions for generating dashboard data and reports.
 */

/**
 * Get dashboard data for the current user
 * 
 * @param {Object} options - Options for dashboard data
 * @return {Object} Dashboard data object
 */
function getDashboardData(options = {}) {
  try {
    // Default options
    const defaults = {
      limit: 10,
      period: 'month' // 'week', 'month', 'quarter', 'year'
    };
    
    // Merge options with defaults
    const opts = { ...defaults, ...options };
    
    // Get current user
    const currentUser = Session.getActiveUser().getEmail();
    const userInfo = getUserInfo();
    
    // Calculate date range based on period
    const now = new Date();
    let startDate = new Date();
    
    switch (opts.period) {
      case 'week':
        startDate.setDate(now.getDate() - 7);
        break;
      case 'month':
        startDate.setMonth(now.getMonth() - 1);
        break;
      case 'quarter':
        startDate.setMonth(now.getMonth() - 3);
        break;
      case 'year':
        startDate.setFullYear(now.getFullYear() - 1);
        break;
      default:
        startDate.setMonth(now.getMonth() - 1); // Default to month
    }
    
    // Get recent evaluations
    const recentEvaluations = getRecentEvaluations(opts.limit, currentUser);
    
    // Get statistics
    const evaluationStats = getEvaluationStatistics(currentUser, startDate, now);
    
    // Get dispute statistics
    const disputeStats = getDisputeStatistics(startDate, now);
    
    // Get user-specific data based on role
    let roleSpecificData = {};
    
    if (userInfo && userInfo.role) {
      switch (userInfo.role) {
        case USER_ROLES.AGENT:
          roleSpecificData = getAgentDashboardData(currentUser, startDate, now);
          break;
        case USER_ROLES.AGENT_MANAGER:
          roleSpecificData = getManagerDashboardData(currentUser, startDate, now);
          break;
        case USER_ROLES.QA_ANALYST:
          roleSpecificData = getQAAnalystDashboardData(currentUser, startDate, now);
          break;
        case USER_ROLES.QA_MANAGER:
        case USER_ROLES.ADMIN:
          roleSpecificData = getQAManagerDashboardData(startDate, now);
          break;
      }
    }
    
    // Compile full dashboard data
    return {
      user: {
        email: currentUser,
        role: userInfo ? userInfo.role : null,
        name: userInfo ? userInfo.name : null
      },
      period: opts.period,
      recentEvaluations: recentEvaluations,
      evaluationStats: evaluationStats,
      disputeStats: disputeStats,
      roleSpecificData: roleSpecificData
    };
  } catch (error) {
    Logger.log(`Error in getDashboardData: ${error.message}`);
    return {
      user: {
        email: Session.getActiveUser().getEmail(),
        role: null,
        name: null
      },
      error: error.message
    };
  }
}

/**
 * Get dashboard data specific to agents
 * 
 * @param {string} agentEmail - The agent's email
 * @param {Date} startDate - Start date for the period
 * @param {Date} endDate - End date for the period
 * @return {Object} Agent-specific dashboard data
 */
function getAgentDashboardData(agentEmail, startDate, endDate) {
  try {
    // Get all evaluations for this agent
    const agentEvaluations = getEvaluationsForAgent(agentEmail);
    
    // Filter by date range
    const filteredEvaluations = agentEvaluations.filter(eval => {
      const evalDate = eval.Date instanceof Date ? eval.Date : new Date(eval.Date);
      return evalDate >= startDate && evalDate <= endDate;
    });
    
    // Calculate trend data (group by week)
    const weeklyScores = {};
    filteredEvaluations.forEach(eval => {
      const evalDate = eval.Date instanceof Date ? eval.Date : new Date(eval.Date);
      const weekStart = new Date(evalDate);
      weekStart.setDate(evalDate.getDate() - evalDate.getDay()); // Set to beginning of week (Sunday)
      weekStart.setHours(0, 0, 0, 0);
      
      const weekKey = weekStart.toISOString().split('T')[0];
      
      if (!weeklyScores[weekKey]) {
        weeklyScores[weekKey] = {
          totalScore: 0,
          totalMaxPossible: 0,
          count: 0,
          week: weekKey
        };
      }
      
      weeklyScores[weekKey].totalScore += parseInt(eval.Score);
      weeklyScores[weekKey].totalMaxPossible += parseInt(eval['Max Possible']);
      weeklyScores[weekKey].count += 1;
    });
    
    // Calculate percentages for each week
    const scoreTrend = Object.values(weeklyScores).map(week => {
      const percentage = week.totalMaxPossible > 0 ? 
        (week.totalScore / week.totalMaxPossible) * 100 : 0;
      
      return {
        week: week.week,
        percentage: parseFloat(percentage.toFixed(1)),
        count: week.count
      };
    });
    
    // Sort by week
    scoreTrend.sort((a, b) => a.week.localeCompare(b.week));
    
    // Get strengths and areas for improvement
    const strengthsMap = {};
    const areasForImprovementMap = {};
    
    filteredEvaluations.forEach(eval => {
      // Process strengths
      if (eval.Strengths) {
        const strengths = eval.Strengths.split(/[.,;]/)
          .map(s => s.trim())
          .filter(s => s.length > 0);
        
        strengths.forEach(strength => {
          strengthsMap[strength] = (strengthsMap[strength] || 0) + 1;
        });
      }
      
      // Process areas for improvement
      if (eval['Areas for Improvement']) {
        const areas = eval['Areas for Improvement'].split(/[.,;]/)
          .map(s => s.trim())
          .filter(s => s.length > 0);
        
        areas.forEach(area => {
          areasForImprovementMap[area] = (areasForImprovementMap[area] || 0) + 1;
        });
      }
    });
    
    // Convert to arrays and sort by frequency
    const topStrengths = Object.entries(strengthsMap)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([strength, count]) => ({ strength, count }));
    
    const topAreasForImprovement = Object.entries(areasForImprovementMap)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([area, count]) => ({ area, count }));
    
    // Calculate pass rate
    const passingThreshold = parseInt(getSetting('passing_score_percentage', '80'));
    let passedCount = 0;
    
    filteredEvaluations.forEach(eval => {
      const score = parseInt(eval.Score);
      const maxPossible = parseInt(eval['Max Possible']);
      const percentage = (score / maxPossible) * 100;
      
      if (percentage >= passingThreshold) {
        passedCount++;
      }
    });
    
    const passRate = filteredEvaluations.length > 0 ? 
      (passedCount / filteredEvaluations.length) * 100 : 0;
    
    // Compile agent dashboard data
    return {
      totalEvaluations: filteredEvaluations.length,
      passRate: parseFloat(passRate.toFixed(1)),
      scoreTrend: scoreTrend,
      topStrengths: topStrengths,
      topAreasForImprovement: topAreasForImprovement
    };
  } catch (error) {
    Logger.log(`Error in getAgentDashboardData: ${error.message}`);
    return {
      error: error.message
    };
  }
}

/**
 * Get dashboard data specific to managers
 * 
 * @param {string} managerEmail - The manager's email
 * @param {Date} startDate - Start date for the period
 * @param {Date} endDate - End date for the period
 * @return {Object} Manager-specific dashboard data
 */
function getManagerDashboardData(managerEmail, startDate, endDate) {
  try {
    // Get all reports (agents) for this manager
    const reports = getUsersByManager(managerEmail);
    const reportEmails = reports.map(user => user.Email);
    
    // Get all evaluations
    const allEvaluations = getAllEvaluations();
    
    // Filter evaluations for reports and date range
    const reportEvaluations = allEvaluations.filter(eval => {
      const evalDate = eval.Date instanceof Date ? eval.Date : new Date(eval.Date);
      return reportEmails.includes(eval.Agent) && 
             evalDate >= startDate && 
             evalDate <= endDate;
    });
    
    // Calculate agent performance data
    const agentPerformance = {};
    reportEvaluations.forEach(eval => {
      const agent = eval.Agent;
      
      if (!agentPerformance[agent]) {
        agentPerformance[agent] = {
          agent: agent,
          name: '',
          totalScore: 0,
          totalMaxPossible: 0,
          count: 0,
          disputes: 0
        };
        
        // Get agent name
        const agentUser = getUserByEmail(agent);
        if (agentUser) {
          agentPerformance[agent].name = agentUser.Name;
        }
      }
      
      agentPerformance[agent].totalScore += parseInt(eval.Score);
      agentPerformance[agent].totalMaxPossible += parseInt(eval['Max Possible']);
      agentPerformance[agent].count += 1;
      
      if (eval.Status === STATUS.DISPUTED) {
        agentPerformance[agent].disputes += 1;
      }
    });
    
    // Calculate percentages and add to array
    const agentPerformanceArray = Object.values(agentPerformance).map(agent => {
      const percentage = agent.totalMaxPossible > 0 ? 
        (agent.totalScore / agent.totalMaxPossible) * 100 : 0;
      
      return {
        agent: agent.agent,
        name: agent.name,
        percentage: parseFloat(percentage.toFixed(1)),
        count: agent.count,
        disputes: agent.disputes
      };
    });
    
    // Sort by percentage (descending)
    agentPerformanceArray.sort((a, b) => b.percentage - a.percentage);
    
    // Get all disputes for reports
    const allDisputes = getAllDisputes();
    const reportDisputes = allDisputes.filter(dispute => {
      // Get the evaluation to check if it's for one of our reports
      const evaluation = getEvaluationById(dispute['Evaluation ID']);
      if (!evaluation) return false;
      
      const disputeDate = dispute['Submission Date'] instanceof Date ? 
        dispute['Submission Date'] : new Date(dispute['Submission Date']);
      
      return reportEmails.includes(evaluation.Agent) &&
             disputeDate >= startDate &&
             disputeDate <= endDate;
    });
    
    // Calculate dispute statistics
    const disputeByStatus = {
      [STATUS.PENDING]: 0,
      [STATUS.APPROVED]: 0,
      [STATUS.PARTIALLY_APPROVED]: 0,
      [STATUS.REJECTED]: 0
    };
    
    reportDisputes.forEach(dispute => {
      if (disputeByStatus[dispute.Status] !== undefined) {
        disputeByStatus[dispute.Status]++;
      }
    });
    
    // Get pending evaluations
    const pendingEvaluations = reportEvaluations.filter(eval => eval.Status === STATUS.PENDING);
    
    // Compile manager dashboard data
    return {
      totalReports: reports.length,
      totalEvaluations: reportEvaluations.length,
      pendingEvaluations: pendingEvaluations.length,
      agentPerformance: agentPerformanceArray,
      disputeStats: {
        total: reportDisputes.length,
        pending: disputeByStatus[STATUS.PENDING],
        approved: disputeByStatus[STATUS.APPROVED],
        partiallyApproved: disputeByStatus[STATUS.PARTIALLY_APPROVED],
        rejected: disputeByStatus[STATUS.REJECTED]
      }
    };
  } catch (error) {
    Logger.log(`Error in getManagerDashboardData: ${error.message}`);
    return {
      error: error.message
    };
  }
}

/**
 * Get dashboard data specific to QA Analysts
 * 
 * @param {string} analystEmail - The QA Analyst's email
 * @param {Date} startDate - Start date for the period
 * @param {Date} endDate - End date for the period
 * @return {Object} QA Analyst-specific dashboard data
 */
function getQAAnalystDashboardData(analystEmail, startDate, endDate) {
  try {
    // Get evaluations created by this analyst
    const analystEvaluations = getEvaluationsByEvaluator(analystEmail);
    
    // Filter by date range
    const filteredEvaluations = analystEvaluations.filter(eval => {
      const evalDate = eval.Date instanceof Date ? eval.Date : new Date(eval.Date);
      return evalDate >= startDate && evalDate <= endDate;
    });
    
    // Calculate activity metrics (evaluations per day)
    const dailyActivity = {};
    filteredEvaluations.forEach(eval => {
      const evalDate = eval.Date instanceof Date ? eval.Date : new Date(eval.Date);
      const dateKey = evalDate.toISOString().split('T')[0];
      
      if (!dailyActivity[dateKey]) {
        dailyActivity[dateKey] = {
          date: dateKey,
          count: 0
        };
      }
      
      dailyActivity[dateKey].count += 1;
    });
    
    // Convert to array and sort by date
    const activityTrend = Object.values(dailyActivity);
    activityTrend.sort((a, b) => a.date.localeCompare(b.date));
    
    // Get evaluation types
    const evaluationTypes = {};
    filteredEvaluations.forEach(eval => {
      const type = eval['Interaction Type'] || 'Unknown';
      evaluationTypes[type] = (evaluationTypes[type] || 0) + 1;
    });
    
    // Convert to array
    const evaluationTypeArray = Object.entries(evaluationTypes).map(([type, count]) => ({
      type,
      count
    }));
    
    // Sort by count (descending)
    evaluationTypeArray.sort((a, b) => b.count - a.count);
    
    // Get disputes for this evaluator's evaluations
    const allDisputes = getAllDisputes();
    const analystDisputes = allDisputes.filter(dispute => {
      // Get the evaluation to check if it was created by this analyst
      const evaluation = getEvaluationById(dispute['Evaluation ID']);
      if (!evaluation) return false;
      
      const disputeDate = dispute['Submission Date'] instanceof Date ? 
        dispute['Submission Date'] : new Date(dispute['Submission Date']);
      
      return evaluation.Evaluator === analystEmail &&
             disputeDate >= startDate &&
             disputeDate <= endDate;
    });
    
    // Calculate dispute rate
    const disputeRate = filteredEvaluations.length > 0 ? 
      (analystDisputes.length / filteredEvaluations.length) * 100 : 0;
    
    // Compile QA Analyst dashboard data
    return {
      totalEvaluations: filteredEvaluations.length,
      activityTrend: activityTrend,
      evaluationTypes: evaluationTypeArray,
      disputes: {
        total: analystDisputes.length,
        rate: parseFloat(disputeRate.toFixed(1))
      }
    };
  } catch (error) {
    Logger.log(`Error in getQAAnalystDashboardData: ${error.message}`);
    return {
      error: error.message
    };
  }
}

/**
 * Get dashboard data specific to QA Managers and Admins
 * 
 * @param {Date} startDate - Start date for the period
 * @param {Date} endDate - End date for the period
 * @return {Object} QA Manager-specific dashboard data
 */
function getQAManagerDashboardData(startDate, endDate) {
  try {
    // Get all evaluations within date range
    const allEvaluations = getAllEvaluations();
    const filteredEvaluations = allEvaluations.filter(eval => {
      const evalDate = eval.Date instanceof Date ? eval.Date : new Date(eval.Date);
      return evalDate >= startDate && evalDate <= endDate;
    });
    
    // Get all disputes within date range
    const allDisputes = getAllDisputes();
    const filteredDisputes = allDisputes.filter(dispute => {
      const disputeDate = dispute['Submission Date'] instanceof Date ? 
        dispute['Submission Date'] : new Date(dispute['Submission Date']);
      return disputeDate >= startDate && disputeDate <= endDate;
    });
    
    // Calculate evaluator performance
    const evaluatorPerformance = {};
    filteredEvaluations.forEach(eval => {
      const evaluator = eval.Evaluator;
      
      if (!evaluatorPerformance[evaluator]) {
        evaluatorPerformance[evaluator] = {
          evaluator: evaluator,
          name: '',
          count: 0,
          disputes: 0
        };
        
        // Get evaluator name
        const evaluatorUser = getUserByEmail(evaluator);
        if (evaluatorUser) {
          evaluatorPerformance[evaluator].name = evaluatorUser.Name;
        }
      }
      
      evaluatorPerformance[evaluator].count += 1;
    });
    
    // Add dispute counts
    filteredDisputes.forEach(dispute => {
      const evaluation = getEvaluationById(dispute['Evaluation ID']);
      if (!evaluation) return;
      
      const evaluator = evaluation.Evaluator;
      if (evaluatorPerformance[evaluator]) {
        evaluatorPerformance[evaluator].disputes += 1;
      }
    });
    
    // Calculate dispute rates and add to array
    const evaluatorPerformanceArray = Object.values(evaluatorPerformance).map(evaluator => {
      const disputeRate = evaluator.count > 0 ? 
        (evaluator.disputes / evaluator.count) * 100 : 0;
      
      return {
        evaluator: evaluator.evaluator,
        name: evaluator.name,
        count: evaluator.count,
        disputes: evaluator.disputes,
        disputeRate: parseFloat(disputeRate.toFixed(1))
      };
    });
    
    // Sort by count (descending)
    evaluatorPerformanceArray.sort((a, b) => b.count - a.count);
    
    // Calculate department performance
    const departmentPerformance = {};
    
    // Get all users
    const users = getAllUsers();
    
    filteredEvaluations.forEach(eval => {
      // Get agent's department
      const agent = users.find(u => u.Email === eval.Agent);
      if (!agent || !agent.Department) return;
      
      const department = agent.Department;
      
      if (!departmentPerformance[department]) {
        departmentPerformance[department] = {
          department: department,
          totalScore: 0,
          totalMaxPossible: 0,
          count: 0
        };
      }
      
      departmentPerformance[department].totalScore += parseInt(eval.Score);
      departmentPerformance[department].totalMaxPossible += parseInt(eval['Max Possible']);
      departmentPerformance[department].count += 1;
    });
    
    // Calculate percentages and add to array
    const departmentPerformanceArray = Object.values(departmentPerformance).map(dept => {
      const percentage = dept.totalMaxPossible > 0 ? 
        (dept.totalScore / dept.totalMaxPossible) * 100 : 0;
      
      return {
        department: dept.department,
        percentage: parseFloat(percentage.toFixed(1)),
        count: dept.count
      };
    });
    
    // Sort by percentage (descending)
    departmentPerformanceArray.sort((a, b) => b.percentage - a.percentage);
    
    // Calculate pending items
    const pendingDisputes = filteredDisputes.filter(dispute => dispute.Status === STATUS.PENDING).length;
    
    // Compile QA Manager dashboard data
    return {
      totalEvaluations: filteredEvaluations.length,
      totalDisputes: filteredDisputes.length,
      pendingDisputes: pendingDisputes,
      evaluatorPerformance: evaluatorPerformanceArray,
      departmentPerformance: departmentPerformanceArray
    };
  } catch (error) {
    Logger.log(`Error in getQAManagerDashboardData: ${error.message}`);
    return {
      error: error.message
    };
  }
}

/**
 * Generate a CSV report of evaluations
 * 
 * @param {Object} options - Report options
 * @return {Object} Result object with success flag, message, and data
 */
function generateEvaluationsReport(options = {}) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.QA_MANAGER) && !hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to generate reports'
      };
    }
    
    // Default options
    const defaults = {
      startDate: null,
      endDate: null,
      agent: null,
      evaluator: null,
      interactionType: null,
      includeAnswers: false
    };
    
    // Merge options with defaults
    const opts = { ...defaults, ...options };
    
    // Get all evaluations
    let evaluations = getAllEvaluations();
    
    // Filter by date range if provided
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
    
    // Filter by agent if provided
    if (opts.agent) {
      evaluations = evaluations.filter(eval => eval.Agent === opts.agent);
    }
    
    // Filter by evaluator if provided
    if (opts.evaluator) {
      evaluations = evaluations.filter(eval => eval.Evaluator === opts.evaluator);
    }
    
    // Filter by interaction type if provided
    if (opts.interactionType) {
      evaluations = evaluations.filter(eval => eval['Interaction Type'] === opts.interactionType);
    }
    
    // Prepare CSV data
    let csvData = [];
    
    // Add headers
    const headers = [
      'ID', 'Date', 'Agent', 'Evaluator', 'Interaction Type', 
      'Customer ID', 'Score', 'Max Possible', 'Status', 
      'Strengths', 'Areas for Improvement', 'Comments'
    ];
    
    csvData.push(headers.join(','));
    
    // Process each evaluation
    evaluations.forEach(eval => {
      // Format the data
      const rowData = [
        eval.ID,
        formatDate(eval.Date),
        eval.Agent,
        eval.Evaluator,
        eval['Interaction Type'] || '',
        eval['Customer ID'] || '',
        eval.Score,
        eval['Max Possible'],
        eval.Status,
        csvEscape(eval.Strengths || ''),
        csvEscape(eval['Areas for Improvement'] || ''),
        csvEscape(eval.Comments || '')
      ];
      
      csvData.push(rowData.join(','));
      
      // Add answers if requested
      if (opts.includeAnswers) {
        const answers = getEvaluationAnswers(eval.ID);
        
        if (answers.length > 0) {
          // Add blank line and answer headers
          csvData.push('');
          csvData.push('Question,Answer,Score,Max Score,Comments');
          
          // Add each answer
          answers.forEach(answer => {
            const answerRow = [
              csvEscape(answer.Question),
              csvEscape(answer.Answer),
              answer.Score,
              answer['Max Score'],
              csvEscape(answer.Comments || '')
            ];
            
            csvData.push(answerRow.join(','));
          });
          
          // Add blank line after answers
          csvData.push('');
        }
      }
    });
    
    // Join CSV data
    const csvString = csvData.join('\n');
    
    // Log the action
    const currentUser = Session.getActiveUser().getEmail();
    logAction(currentUser, 'Generate Report', 
              `Generated evaluations report with ${evaluations.length} evaluations`);
    
    return {
      success: true,
      message: `Report generated with ${evaluations.length} evaluations`,
      data: csvString
    };
  } catch (error) {
    Logger.log(`Error in generateEvaluationsReport: ${error.message}`);
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
}

/**
 * Generate a CSV report of disputes
 * 
 * @param {Object} options - Report options
 * @return {Object} Result object with success flag, message, and data
 */
function generateDisputesReport(options = {}) {
  try {
    // Check permissions
    if (!hasPermission(USER_ROLES.QA_MANAGER) && !hasPermission(USER_ROLES.ADMIN)) {
      return {
        success: false,
        message: 'You do not have permission to generate reports'
      };
    }
    
    // Default options
    const defaults = {
      startDate: null,
      endDate: null,
      status: null,
      submitter: null
    };
    
    // Merge options with defaults
    const opts = { ...defaults, ...options };
    
    // Get all disputes
    let disputes = getAllDisputes();
    
    // Filter by date range if provided
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
    
    // Filter by status if provided
    if (opts.status) {
      disputes = disputes.filter(dispute => dispute.Status === opts.status);
    }
    
    // Filter by submitter if provided
    if (opts.submitter) {
      disputes = disputes.filter(dispute => dispute['Submitted By'] === opts.submitter);
    }
    
    // Prepare CSV data
    let csvData = [];
    
    // Add headers
    const headers = [
      'ID', 'Evaluation ID', 'Submitted By', 'Submission Date', 
      'Reason', 'Status', 'Reviewed By', 'Review Date', 
      'Score Adjustment', 'Details'
    ];
    
    csvData.push(headers.join(','));
    
    // Process each dispute
    disputes.forEach(dispute => {
      // Format the data
      const rowData = [
        dispute.ID,
        dispute['Evaluation ID'],
        dispute['Submitted By'],
        formatDate(dispute['Submission Date']),
        csvEscape(dispute.Reason),
        dispute.Status,
        dispute['Reviewed By'] || '',
        formatDate(dispute['Review Date'] || ''),
        dispute['Score Adjustment'] || '0',
        csvEscape(dispute.Details || '')
      ];
      
      csvData.push(rowData.join(','));
    });
    
    // Join CSV data
    const csvString = csvData.join('\n');
    
    // Log the action
    const currentUser = Session.getActiveUser().getEmail();
    logAction(currentUser, 'Generate Report', 
              `Generated disputes report with ${disputes.length} disputes`);
    
    return {
      success: true,
      message: `Report generated with ${disputes.length} disputes`,
      data: csvString
    };
  } catch (error) {
    Logger.log(`Error in generateDisputesReport: ${error.message}`);
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
 * Format a date for display
 * 
 * @param {Date|string} date - The date to format
 * @return {string} The formatted date string
 */
function formatDate(date) {
  if (!date) return '';
  
  if (typeof date === 'string') {
    date = new Date(date);
  }
  
  if (isNaN(date.getTime())) {
    return '';
  }
  
  return date.toISOString().split('T')[0];
}