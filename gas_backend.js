/**
 * MISSION BOARD — GAS Backend
 * Single Source of Truth for SAT × G shared task board
 * 
 * Spreadsheet Schema (MB_Tasks sheet):
 * A:id | B:title | C:category | D:priority | E:status | F:createdBy | G:createdAt | H:notes | I:links | J:isDaily | K:dailyChecked
 */

const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SHEET_NAME = 'MB_Tasks';

// ═══════════════ ENTRY POINTS ═══════════════

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'list';
  let result;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    switch (action) {
      case 'list':
        result = listTasks(ss);
        break;
      case 'ping':
        result = { success: true, message: 'MISSION BOARD ONLINE', timestamp: getJST() };
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    switch (action) {
      case 'save':
        result = saveTask(ss, body.task);
        break;
      case 'delete':
        result = deleteTask(ss, body.id);
        break;
      case 'bulk':
        result = bulkSave(ss, body.tasks);
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════ CORE OPERATIONS ═══════════════

function listTasks(ss) {
  const sheet = getOrCreateSheet(ss);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, tasks: [] };
  
  const headers = data[0];
  const hMap = {};
  headers.forEach((h, i) => hMap[h] = i);
  
  const tasks = data.slice(1).map(row => ({
    id:        String(row[hMap['id']] || ''),
    title:     String(row[hMap['title']] || ''),
    category:  String(row[hMap['category']] || 'SVD-OS'),
    priority:  String(row[hMap['priority']] || 'medium'),
    status:    String(row[hMap['status']] || 'todo'),
    createdBy: String(row[hMap['createdBy']] || 'SAT'),
    createdAt: String(row[hMap['createdAt']] || ''),
    notes:     String(row[hMap['notes']] || ''),
    links:     parseLinks(row[hMap['links']]),
    isDaily:   String(row[hMap['isDaily']] || '') === 'true',
    dailyChecked: String(row[hMap['dailyChecked']] || '') || null
  })).filter(t => t.id && t.title);
  
  return { success: true, tasks: tasks };
}

function saveTask(ss, task) {
  if (!task || !task.title) return { error: 'Title is required' };
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getOrCreateSheet(ss);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const hMap = {};
    headers.forEach((h, i) => hMap[h] = i);
    
    // Generate ID if new
    if (!task.id) {
      task.id = 'mb-' + new Date().getTime().toString(36) + '-' + Math.random().toString(36).substr(2, 5);
    }
    if (!task.createdAt) task.createdAt = getJST();
    
    // Find existing row
    const idCol = data.map(r => String(r[0]));
    const existingRow = idCol.indexOf(task.id);
    
    const rowData = [
      task.id,
      task.title,
      task.category || 'SVD-OS',
      task.priority || 'medium',
      task.status || 'todo',
      task.createdBy || 'SAT',
      task.createdAt,
      task.notes || '',
      JSON.stringify(task.links || []),
      String(!!task.isDaily),
      task.dailyChecked || ''
    ];
    
    if (existingRow > 0) {
      // Update
      sheet.getRange(existingRow + 1, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // Append
      sheet.appendRow(rowData);
    }
    
    SpreadsheetApp.flush();
    return { success: true, task: task };
  } catch (err) {
    return { error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteTask(ss, id) {
  if (!id) return { error: 'ID is required' };
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getOrCreateSheet(ss);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true, deleted: id };
      }
    }
    
    return { error: 'Task not found: ' + id };
  } catch (err) {
    return { error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function bulkSave(ss, tasks) {
  if (!tasks || !tasks.length) return { error: 'No tasks provided' };
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = getOrCreateSheet(ss);
    
    // Clear existing data (keep headers)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    
    // Write all tasks
    const rows = tasks.map(t => [
      t.id || 'mb-' + new Date().getTime().toString(36) + '-' + Math.random().toString(36).substr(2, 5),
      t.title || '',
      t.category || 'SVD-OS',
      t.priority || 'medium',
      t.status || 'todo',
      t.createdBy || 'SAT',
      t.createdAt || getJST(),
      t.notes || '',
      JSON.stringify(t.links || []),
      String(!!t.isDaily),
      t.dailyChecked || ''
    ]);
    
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    
    SpreadsheetApp.flush();
    return { success: true, count: rows.length };
  } catch (err) {
    return { error: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════ UTILITIES ═══════════════

function getOrCreateSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'title', 'category', 'priority', 'status', 'createdBy', 'createdAt', 'notes', 'links', 'isDaily', 'dailyChecked']);
    sheet.getRange('A:A').setNumberFormat('@'); // Force text format for IDs
    SpreadsheetApp.flush();
  }
  return sheet;
}

function getJST() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss'+09:00'");
}

function parseLinks(val) {
  if (!val) return [];
  try {
    const parsed = JSON.parse(val);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}
