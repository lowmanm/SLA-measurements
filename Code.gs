/**
 * QA Platform - Main Application Entry Point
 * This file contains the main menu setup and global configurations
 */

// Global constants
const SHEET_NAMES = {
  AUDIT_QUEUE: 'auditQueue',
  EVALUATIONS: 'evaluations',
  QUESTIONS: 'questions',
  USERS: 'users',
  SETTINGS: 'settings',
  DISPUTES: 'disputes',
  TEMPLATES: 'evalTemplates'
};

const USER_ROLES = {
  QA_ANALYST: 'QA Analyst',
  AGENT_MANAGER: 'Agent Manager',
  QA_MANAGER: 'QA Manager',
  ADMIN: 'Admin'
};

// Global variables
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var currentUser = Session.getActiveUser().getEmail();

/**
 * Runs when the spreadsheet is opened
 * Sets up the main menu
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('QA Platform')
    .addItem('Show Main Panel', 'showMainSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('Evaluation')
      .addItem('Start New Evaluation', 'showEvaluationForm')
      .addItem('View My Evaluations', 'showMyEvaluations'))
    .addSubMenu(ui.createMenu('Disputes')
      .addItem('File Dispute', 'showDisputeForm')
      .addItem('Review Disputes', 'showDisputeReview'))
    .addSubMenu(ui.createMenu('Admin')
      .addItem('Manage Questions', 'showQuestionManager')
      .addItem('Manage Users', 'showUserManager')
      .addItem('System Settings', 'showSettingsPanel'))
    .addItem('Dashboard', 'showDashboard')
    .addItem('Import Audit Queue', 'importFromGmail')
    .addToUi();
}

/**
 * Shows the main sidebar with appropriate options based on user role
 */
function showMainSidebar() {
  const userRole = getUserRole(currentUser);
  const html = HtmlService.createTemplateFromFile('UI/MainSidebar')
    .evaluate()
    .setTitle('QA Platform')
    .setWidth(300);
  
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Includes HTML files for modular code organization
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Initialization function to set up the spreadsheet structure
 * This can be run manually to create all necessary sheets
 */
function initializeSystem() {
  // Create all required sheets if they don't exist
  createSheetIfNotExists(SHEET_NAMES.AUDIT_QUEUE, [
    'ID', 'Date', 'Agent', 'Customer', 'Interaction Type', 'Duration', 
    'Queue Status', 'Assigned To', 'Priority', 'Import Date'
  ]);
  
  createSheetIfNotExists(SHEET_NAMES.EVALUATIONS, [
    'ID', 'Audit ID', 'Evaluation Date', 'Evaluator', 'Agent', 'Score', 
    'Max Possible', 'Status', 'Disputed', 'Last Updated'
  ]);
  
  createSheetIfNotExists(SHEET_NAMES.QUESTIONS, [
    'ID', 'Template ID', 'Question', 'Category', 'Weight', 'Critical', 'Active'
  ]);
  
  createSheetIfNotExists(SHEET_NAMES.USERS, [
    'Email', 'Name', 'Role', 'Active', 'Last Login', 'Department'
  ]);
  
  createSheetIfNotExists(SHEET_NAMES.SETTINGS, [
    'Setting', 'Value', 'Description'
  ]);
  
  createSheetIfNotExists(SHEET_NAMES.DISPUTES, [
    'ID', 'Evaluation ID', 'Date Filed', 'Filed By', 'Status', 'Reviewer', 
    'Resolution Date', 'Resolution Notes', 'Original Score', 'Adjusted Score'
  ]);
  
  createSheetIfNotExists(SHEET_NAMES.TEMPLATES, [
    'ID', 'Name', 'Description', 'Interaction Type', 'Active', 'Created By', 'Creation Date'
  ]);
  
  // Add default settings
  const settingsSheet = spreadsheet.getSheetByName(SHEET_NAMES.SETTINGS);
  if (settingsSheet.getLastRow() <= 1) {
    settingsSheet.appendRow(['notification_email_enabled', 'true', 'Enable email notifications']);
    settingsSheet.appendRow(['dispute_window_days', '7', 'Number of days allowed for filing disputes']);
    settingsSheet.appendRow(['min_evaluations_per_agent', '4', 'Minimum evaluations per agent per month']);
    settingsSheet.appendRow(['passing_score_percentage', '80', 'Minimum passing score percentage']);
  }
  
  // Add the current user as admin if users sheet is empty
  const usersSheet = spreadsheet.getSheetByName(SHEET_NAMES.USERS);
  if (usersSheet.getLastRow() <= 1) {
    usersSheet.appendRow([currentUser, 'System Admin', USER_ROLES.ADMIN, 'TRUE', new Date(), 'Administration']);
  }
}

/**
 * Helper function to create a sheet if it doesn't exist
 */
function createSheetIfNotExists(sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.appendRow(headers);
    
    // Format the header row
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4285F4')
      .setFontColor('white')
      .setFontWeight('bold');
  }
  
  return sheet;
}