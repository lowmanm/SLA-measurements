/**
 * QA Platform - Main Code File
 *
 * This is the main entry point for the QA Platform application.
 * It contains menu creation, global constants, and initialization.
 */

/**
 * doGet function - Required entry point for web app deployment
 * This function is called when the web app is accessed via a GET request
 * 
 * @param {Object} e - Event object containing request parameters
 * @return {HtmlOutput} HTML content to display
 */
function doGet(e) {
  // Determine which page to display based on parameters
  const params = e.parameter || {};
  const page = params.page || 'main';
  
  // Initialize database if needed (only run by admin)
  const isInitialized = isDatabaseInitialized();
  if (!isInitialized) {
    const userInfo = getUserInfo();
    if (userInfo && userInfo.role === USER_ROLES.ADMIN) {
      // Auto-initialize for admin users
      initializeDatabase();
    } else if (page !== 'login') {
      // Redirect non-admins to login page if database is not initialized
      return HtmlService.createTemplateFromFile('UI/LoginPage')
        .evaluate()
        .setTitle('QA Platform - Login Required')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }
  
  // Check if user is authenticated (except for login page)
  if (page !== 'login') {
    const userInfo = getUserInfo();
    if (!userInfo || !userInfo.role) {
      // User is not authenticated, redirect to login
      return HtmlService.createTemplateFromFile('UI/LoginPage')
        .evaluate()
        .setTitle('QA Platform - Login Required')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }
  
  // Serve the appropriate page
  let template;
  let title = 'QA Platform';
  
  switch (page) {
    case 'login':
      template = HtmlService.createTemplateFromFile('UI/LoginPage');
      title = 'QA Platform - Login';
      break;
      
    case 'dashboard':
      template = HtmlService.createTemplateFromFile('UI/Dashboard');
      title = 'QA Platform - Dashboard';
      break;
      
    case 'evaluation':
      template = HtmlService.createTemplateFromFile('UI/EvaluationForm');
      title = 'QA Platform - Evaluation';
      break;
      
    case 'dispute':
      template = HtmlService.createTemplateFromFile('UI/DisputeForm');
      title = 'QA Platform - File Dispute';
      break;
      
    case 'review':
      template = HtmlService.createTemplateFromFile('UI/DisputeReview');
      title = 'QA Platform - Review Disputes';
      break;
      
    case 'users':
      template = HtmlService.createTemplateFromFile('UI/UserManagement');
      title = 'QA Platform - User Management';
      break;
      
    case 'admin':
      template = HtmlService.createTemplateFromFile('UI/AdminPanel');
      title = 'QA Platform - Administration';
      break;
      
    default:
      // Main application page with navigation
      template = HtmlService.createTemplateFromFile('UI/MainApp');
      title = 'QA Platform';
  }
  
  // Render the template and return the HTML output
  return template
    .evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Global Constants
const SHEET_NAMES = {
  USERS: 'Users',
  QUESTION_SETS: 'Question Sets',
  QUESTIONS: 'Questions',
  AUDIT_QUEUE: 'Audit Queue',
  EVALUATIONS: 'Evaluations',
  EVALUATION_ANSWERS: 'Evaluation Answers',
  DISPUTES: 'Disputes',
  DISPUTE_RESOLUTIONS: 'Dispute Resolutions',
  SETTINGS: 'Settings',
  LOGS: 'Logs'
};

// User role constants
const USER_ROLES = {
  AGENT: 'agent',
  AGENT_MANAGER: 'agent_manager',
  QA_ANALYST: 'qa_analyst',
  QA_MANAGER: 'qa_manager',
  ADMIN: 'admin'
};

// Status constants
const STATUS = {
  PENDING: 'Pending',
  IN_PROGRESS: 'In Progress',
  COMPLETED: 'Completed',
  FAILED: 'Failed',
  DISPUTED: 'Disputed',
  APPROVED: 'Approved',
  REJECTED: 'Rejected',
  PARTIALLY_APPROVED: 'Partially Approved'
};

/**
 * Runs when the spreadsheet is opened
 * Creates the custom menu
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Create the main menu
  const menu = ui.createMenu('QA Platform');
  
  // Add menu items
  menu.addItem('Open QA Platform', 'showMainSidebar');
  menu.addSeparator();
  
  // Get current user info
  const userInfo = getUserInfo();
  
  // Add role-specific menu items
  if (userInfo && userInfo.role) {
    // QA Analyst specific menu items
    if (hasPermission(USER_ROLES.QA_ANALYST) || hasPermission(USER_ROLES.QA_MANAGER) || hasPermission(USER_ROLES.ADMIN)) {
      menu.addItem('New Evaluation', 'showEvaluationForm');
    }
    
    // Agent Manager specific menu items
    if (hasPermission(USER_ROLES.AGENT_MANAGER) || hasPermission(USER_ROLES.ADMIN)) {
      menu.addItem('File Dispute', 'showDisputeForm');
    }
    
    // QA Manager specific menu items
    if (hasPermission(USER_ROLES.QA_MANAGER) || hasPermission(USER_ROLES.ADMIN)) {
      menu.addItem('Review Disputes', 'showDisputeReview');
    }
    
    // Common menu items for authenticated users
    menu.addSeparator();
    menu.addItem('Dashboard', 'showDashboard');
  }
  
  // Admin specific menu items
  if (hasPermission(USER_ROLES.ADMIN)) {
    menu.addSeparator();
    menu.addSubMenu(ui.createMenu('Admin')
      .addItem('Manage Question Sets', 'showQuestionManager')
      .addItem('Manage Users', 'showUserManager')
      .addItem('Settings', 'showSettingsPanel')
      .addSeparator()
      .addItem('Initialize Database', 'initializeDatabase')
      .addItem('Import From Gmail', 'importFromGmail')
      .addItem('Export Data', 'showExportPanel'));
  }
  
  // Add help menu
  menu.addSeparator();
  menu.addItem('Help', 'showHelp');
  menu.addItem('About', 'showAbout');
  
  // Add the menu to the UI
  menu.addToUi();
}

/**
 * Show the main sidebar
 */
function showMainSidebar() {
  // Check if database is initialized
  const isInitialized = isDatabaseInitialized();
  
  // If not initialized and user is admin, show initialization dialog
  if (!isInitialized && hasPermission(USER_ROLES.ADMIN)) {
    showInitializationDialog();
    return;
  } else if (!isInitialized) {
    // If not initialized and user is not admin, show message
    SpreadsheetApp.getUi().alert(
      'QA Platform Not Initialized', 
      'The QA Platform has not been initialized. Please contact the system administrator.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Get the sidebar content
  const html = HtmlService.createTemplateFromFile('UI/MainSidebar')
    .evaluate()
    .setTitle('QA Platform')
    .setWidth(300);
  
  // Show the sidebar
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Show the evaluation form
 */
function showEvaluationForm() {
  // Check permissions
  if (!hasPermission(USER_ROLES.QA_ANALYST) && 
      !hasPermission(USER_ROLES.QA_MANAGER) && 
      !hasPermission(USER_ROLES.ADMIN)) {
    SpreadsheetApp.getUi().alert('You do not have permission to perform evaluations.');
    return;
  }
  
  // Create and show the dialog
  const html = HtmlService.createTemplateFromFile('UI/EvaluationForm')
    .evaluate()
    .setWidth(800)
    .setHeight(600)
    .setTitle('QA Evaluation');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'QA Evaluation');
}

/**
 * Show the dispute form
 */
function showDisputeForm() {
  // Check permissions
  if (!hasPermission(USER_ROLES.AGENT_MANAGER) && !hasPermission(USER_ROLES.ADMIN)) {
    SpreadsheetApp.getUi().alert('You do not have permission to file disputes.');
    return;
  }
  
  // Create and show the dialog
  const html = HtmlService.createTemplateFromFile('UI/DisputeForm')
    .evaluate()
    .setWidth(800)
    .setHeight(600)
    .setTitle('File Dispute');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'File Dispute');
}

/**
 * Show the dispute review interface
 */
function showDisputeReview() {
  // Check permissions
  if (!hasPermission(USER_ROLES.QA_MANAGER) && !hasPermission(USER_ROLES.ADMIN)) {
    SpreadsheetApp.getUi().alert('You do not have permission to review disputes.');
    return;
  }
  
  // Create and show the dialog
  const html = HtmlService.createTemplateFromFile('UI/DisputeReview')
    .evaluate()
    .setWidth(800)
    .setHeight(600)
    .setTitle('Dispute Review');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Dispute Review');
}

/**
 * Show the dashboard
 */
function showDashboard() {
  // Create and show the dialog
  const html = HtmlService.createTemplateFromFile('UI/Dashboard')
    .evaluate()
    .setWidth(800)
    .setHeight(600)
    .setTitle('QA Dashboard');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'QA Dashboard');
}

/**
 * Show the question set manager
 */
function showQuestionManager() {
  // Check permissions
  if (!hasPermission(USER_ROLES.ADMIN)) {
    SpreadsheetApp.getUi().alert('You do not have permission to manage question sets.');
    return;
  }
  
  // Create and show the dialog
  const html = HtmlService.createTemplateFromFile('UI/AdminPanel')
    .evaluate()
    .setWidth(800)
    .setHeight(600)
    .setTitle('Question Set Manager');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Question Set Manager');
}

/**
 * Show the user manager
 */
function showUserManager() {
  // Check permissions
  if (!hasPermission(USER_ROLES.ADMIN)) {
    SpreadsheetApp.getUi().alert('You do not have permission to manage users.');
    return;
  }
  
  // Create and show the dialog
  const html = HtmlService.createTemplateFromFile('UI/UserManagement')
    .evaluate()
    .setWidth(800)
    .setHeight(600)
    .setTitle('User Management');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'User Management');
}

/**
 * Show the settings panel
 */
function showSettingsPanel() {
  // Check permissions
  if (!hasPermission(USER_ROLES.ADMIN)) {
    SpreadsheetApp.getUi().alert('You do not have permission to modify settings.');
    return;
  }
  
  // Create and show the dialog
  const html = HtmlService.createTemplateFromFile('UI/AdminPanel')
    .evaluate()
    .setWidth(800)
    .setHeight(600)
    .setTitle('Settings');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
}

/**
 * Show the export panel
 */
function showExportPanel() {
  // Check permissions
  if (!hasPermission(USER_ROLES.QA_MANAGER) && !hasPermission(USER_ROLES.ADMIN)) {
    SpreadsheetApp.getUi().alert('You do not have permission to export data.');
    return;
  }
  
  // Create and show the dialog
  const html = HtmlService.createHtmlOutput('<p>Export functionality would be implemented here.</p>')
    .setWidth(400)
    .setHeight(300)
    .setTitle('Export Data');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Export Data');
}

/**
 * Show the help dialog
 */
function showHelp() {
  // Create and show the dialog
  const html = HtmlService.createHtmlOutput(`
    <h2>QA Platform Help</h2>
    <p>The QA Platform helps you evaluate agent performance and manage quality assurance processes.</p>
    <h3>Key Features</h3>
    <ul>
      <li>Conduct evaluations with customizable question sets</li>
      <li>File and review disputes for evaluations</li>
      <li>Generate performance reports and metrics</li>
      <li>Import and export data</li>
    </ul>
    <p>For more information, please contact the system administrator.</p>
  `)
    .setWidth(400)
    .setHeight(300)
    .setTitle('Help');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Help');
}

/**
 * Show the about dialog
 */
function showAbout() {
  // Get the version from settings
  const settings = getSettings();
  const version = settings.version || '1.0.0';
  
  // Create and show the dialog
  const html = HtmlService.createHtmlOutput(`
    <div style="text-align: center; padding: 20px;">
      <h2>QA Platform</h2>
      <p>Version ${version}</p>
      <p>A quality assurance platform for evaluating agent performance.</p>
      <p>&copy; 2025 All rights reserved.</p>
    </div>
  `)
    .setWidth(300)
    .setHeight(200)
    .setTitle('About QA Platform');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'About');
}

/**
 * Show the initialization dialog
 */
function showInitializationDialog() {
  const ui = SpreadsheetApp.getUi();
  
  // Show dialog
  const response = ui.alert(
    'Initialize QA Platform', 
    'The QA Platform needs to be initialized. This will create all required sheets and initial data. Do you want to proceed?', 
    ui.ButtonSet.YES_NO);
  
  // Process response
  if (response === ui.Button.YES) {
    initializeDatabase();
  }
}

/**
 * Include HTML file
 * Used by HTML templates to include other HTML files
 * 
 * @param {string} filename - The name of the file to include
 * @return {string} The contents of the file
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get user info for the current user
 * 
 * @return {Object} User info object
 */
function getUserInfo() {
  try {
    // Get the current user's email
    const email = Session.getActiveUser().getEmail();
    
    // Check if the Users sheet exists
    if (!sheetExists(SHEET_NAMES.USERS)) {
      return {
        email: email,
        role: isSystemAdmin(email) ? USER_ROLES.ADMIN : null,
        isInitialized: false
      };
    }
    
    // Look up the user in the database
    const users = getDataFromSheet(SHEET_NAMES.USERS);
    const user = users.find(u => u.Email === email);
    
    // If user found, return their info
    if (user) {
      return {
        email: email,
        role: user.Role,
        name: user.Name,
        department: user.Department,
        manager: user.Manager,
        isInitialized: true
      };
    }
    
    // If sheet exists but user not found, check if they are a system admin
    if (isSystemAdmin(email)) {
      return {
        email: email,
        role: USER_ROLES.ADMIN,
        isInitialized: true
      };
    }
    
    // User not found and not admin
    return {
      email: email,
      role: null,
      isInitialized: true
    };
  } catch (error) {
    Logger.log(`Error in getUserInfo: ${error.message}`);
    return {
      email: Session.getActiveUser().getEmail(),
      role: null,
      error: error.message,
      isInitialized: false
    };
  }
}

/**
 * Check if a user has a specific permission based on their role
 * 
 * @param {string} requiredRole - The role required for the permission
 * @return {boolean} True if the user has the required permission
 */
function hasPermission(requiredRole) {
  const userInfo = getUserInfo();
  
  // If user info couldn't be retrieved, deny permission
  if (!userInfo || !userInfo.role) {
    return false;
  }
  
  // Admin has all permissions
  if (userInfo.role === USER_ROLES.ADMIN) {
    return true;
  }
  
  // QA Manager has QA Analyst permissions
  if (requiredRole === USER_ROLES.QA_ANALYST && userInfo.role === USER_ROLES.QA_MANAGER) {
    return true;
  }
  
  // Exact role match
  return userInfo.role === requiredRole;
}

/**
 * Check if a user is a system administrator
 * 
 * @param {string} email - The email to check
 * @return {boolean} True if the user is a system administrator
 */
function isSystemAdmin(email) {
  // In a real implementation, this might check against a configured list of admins
  // For simplicity, we'll consider the sheet owner as admin
  const owner = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
  return email === owner;
}

/**
 * Check if the database has been initialized
 * 
 * @return {boolean} True if the database has been initialized
 */
function isDatabaseInitialized() {
  // Check if all required sheets exist
  for (const sheetName in SHEET_NAMES) {
    if (!sheetExists(SHEET_NAMES[sheetName])) {
      return false;
    }
  }
  
  return true;
}

/**
 * Check if a sheet exists in the spreadsheet
 * 
 * @param {string} name - The name of the sheet to check
 * @return {boolean} True if the sheet exists
 */
function sheetExists(name) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(name);
  return sheet !== null;
}

/**
 * Generate a unique ID
 * 
 * @return {string} A unique ID
 */
function generateUniqueId() {
  const timestamp = new Date().getTime().toString();
  const random = Math.random().toString(36).substring(2, 8);
  return `${timestamp}_${random}`;
}