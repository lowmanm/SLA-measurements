/**
 * Admin Functions
 * 
 * This file contains functions for system administration,
 * including question management and system settings.
 */

/**
 * Show the question manager UI
 */
function showQuestionManager() {
  if (!hasPermission(USER_ROLES.ADMIN)) {
    SpreadsheetApp.getUi().alert('You do not have permission to manage questions');
    return;
  }
  
  const html = HtmlService.createTemplateFromFile('UI/QuestionManager')
    .evaluate()
    .setTitle('Question Manager')
    .setWidth(800)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Question Manager');
}

/**
 * Show the settings panel UI
 */
function showSettingsPanel() {
  if (!hasPermission(USER_ROLES.ADMIN)) {
    SpreadsheetApp.getUi().alert('You do not have permission to manage system settings');
    return;
  }
  
  const html = HtmlService.createTemplateFromFile('UI/AdminPanel')
    .evaluate()
    .setTitle('System Settings')
    .setWidth(600)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'System Settings');
}

/**
 * Get all system settings
 * 
 * @return {Array} Array of setting objects
 */
function getAllSettings() {
  if (!hasPermission(USER_ROLES.ADMIN)) {
    throw new Error('You do not have permission to view system settings');
  }
  
  return getDataFromSheet(SHEET_NAMES.SETTINGS);
}

/**
 * Update multiple system settings
 * 
 * @param {Object} settings - Object with setting names and values
 * @return {boolean} Success status
 */
function updateSystemSettings(settings) {
  if (!hasPermission(USER_ROLES.ADMIN)) {
    throw new Error('You do not have permission to update system settings');
  }
  
  for (const settingName in settings) {
    updateSetting(settingName, settings[settingName]);
  }
  
  return true;
}

/**
 * Get all evaluation templates
 * 
 * @return {Array} Array of template objects
 */
function getAllTemplates() {
  if (!hasPermission(USER_ROLES.ADMIN) && !hasPermission(USER_ROLES.QA_MANAGER)) {
    throw new Error('You do not have permission to view evaluation templates');
  }
  
  return getDataFromSheet(SHEET_NAMES.TEMPLATES);
}

/**
 * Create a new evaluation template
 * 
 * @param {string} name - Template name
 * @param {string} description - Template description
 * @param {string} interactionType - Type of interaction this template is for
 * @return {Object} Newly created template
 */
function createTemplate(name, description, interactionType) {
  if (!hasPermission(USER_ROLES.ADMIN)) {
    throw new Error('You do not have permission to create templates');
  }
  
  const template = {
    ID: generateUniqueId(),
    Name: name,
    Description: description,
    'Interaction Type': interactionType,
    Active: true,
    'Created By': currentUser,
    'Creation Date': new Date()
  };
  
  addRowToSheet(SHEET_NAMES.TEMPLATES, template);
  return template;
}

/**
 * Update an evaluation template
 * 
 * @param {string} templateId - Template ID
 * @param {Object} templateData - Template data to update
 * @return {boolean} Success status
 */
function updateTemplate(templateId, templateData) {
  if (!hasPermission(USER_ROLES.ADMIN)) {
    throw new Error('You do not have permission to update templates');
  }
  
  const template = findRowById(SHEET_NAMES.TEMPLATES, templateId);
  if (!template) {
    throw new Error(`Template with ID ${templateId} not found`);
  }
  
  updateRowInSheet(SHEET_NAMES.TEMPLATES, template.rowIndex, templateData);
  return true;
}

/**
 * Get all questions for a template
 * 
 * @param {string} templateId - Template ID
 * @return {Array} Array of question objects
 */
function getQuestionsForTemplate(templateId) {
  if (!hasPermission(USER_ROLES.ADMIN) && !hasPermission(USER_ROLES.QA_MANAGER)) {
    throw new Error('You do not have permission to view questions');
  }
  
  const questions = getDataFromSheet(SHEET_NAMES.QUESTIONS);
  return questions.filter(q => q['Template ID'] === templateId);
}

/**
 * Add a question to a template
 * 
 * @param {string} templateId - Template ID
 * @param {string} question - Question text
 * @param {string} category - Question category
 * @param {number} weight - Question weight
 * @param {boolean} critical - Whether the question is critical
 * @return {Object} Newly created question
 */
function addQuestion(templateId, question, category, weight, critical) {
  if (!hasPermission(USER_ROLES.ADMIN)) {
    throw new Error('You do not have permission to add questions');
  }
  
  // Verify template exists
  const template = findRowById(SHEET_NAMES.TEMPLATES, templateId);
  if (!template) {
    throw new Error(`Template with ID ${templateId} not found`);
  }
  
  const questionData = {
    ID: generateUniqueId(),
    'Template ID': templateId,
    Question: question,
    Category: category,
    Weight: weight,
    Critical: critical,
    Active: true
  };
  
  addRowToSheet(SHEET_NAMES.QUESTIONS, questionData);
  return questionData;
}

/**
 * Update a question
 * 
 * @param {string} questionId - Question ID
 * @param {Object} questionData - Question data to update
 * @return {boolean} Success status
 */
function updateQuestion(questionId, questionData) {
  if (!hasPermission(USER_ROLES.ADMIN)) {
    throw new Error('You do not have permission to update questions');
  }
  
  const question = findRowById(SHEET_NAMES.QUESTIONS, questionId);
  if (!question) {
    throw new Error(`Question with ID ${questionId} not found`);
  }
  
  updateRowInSheet(SHEET_NAMES.QUESTIONS, question.rowIndex, questionData);
  return true;
}

/**
 * Delete a question
 * 
 * @param {string} questionId - Question ID
 * @return {boolean} Success status
 */
function deleteQuestion(questionId) {
  if (!hasPermission(USER_ROLES.ADMIN)) {
    throw new Error('You do not have permission to delete questions');
  }
  
  const question = findRowById(SHEET_NAMES.QUESTIONS, questionId);
  if (!question) {
    throw new Error(`Question with ID ${questionId} not found`);
  }
  
  // Soft delete - mark as inactive
  updateRowInSheet(SHEET_NAMES.QUESTIONS, question.rowIndex, { Active: false });
  return true;
}