/**
 * MISSION BOARD — GAS Backend
 * Single Source of Truth for SAT × G shared task board
 * 
 * Spreadsheet Schema (MB_Tasks sheet):
 * A:id | B:title | C:category | D:priority | E:status | F:createdBy | G:createdAt | H:completedAt | I:notes | J:links | K:isDaily | L:dailyChecked
 *
 * VALID VALUES:
 *   category: SVD-OS | Personal | Session
 *   priority: high | medium | low
 *   status:   todo | in-progress | done
 */

const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SHEET_NAME = 'MB_Tasks';
const BM_SHEET_NAME = 'MB_Bookmarks';

// Validation constants
const VALID_CATS = ['SVD-OS', 'Personal', 'Session'];
const VALID_PRI = ['high', 'medium', 'low'];
const VALID_STATUS = ['todo', 'in-progress', 'done'];

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
      case 'listBm':
        result = listBookmarks(ss);
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
      case 'saveBm':
        result = saveBookmark(ss, body.bookmark);
        break;
      case 'deleteBm':
        result = deleteBookmark(ss, body.id);
        break;
      case 'bulkBm':
        result = bulkSaveBookmarks(ss, body.bookmarks);
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
  
  const tasks = data.slice(1).map(row => {
    const cat = String(row[hMap['category']] || '');
    const pri = String(row[hMap['priority']] || '');
    const sts = String(row[hMap['status']] || '');
    return {
      id:          String(row[hMap['id']] || ''),
      title:       String(row[hMap['title']] || ''),
      category:    VALID_CATS.includes(cat) ? cat : 'SVD-OS',
      priority:    VALID_PRI.includes(pri) ? pri : 'medium',
      status:      VALID_STATUS.includes(sts) ? sts : 'todo',
      createdBy:   String(row[hMap['createdBy']] || 'SAT'),
      createdAt:   String(row[hMap['createdAt']] || ''),
      completedAt: String(row[hMap['completedAt']] || '') || null,
      notes:       String(row[hMap['notes']] || ''),
      links:       parseLinks(row[hMap['links']]),
      isDaily:     String(row[hMap['isDaily']] || '') === 'true',
      dailyChecked: String(row[hMap['dailyChecked']] || '') || null
    };
  }).filter(t => t.id && t.title);
  
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
    
    // Validate category/priority/status
    const cat = VALID_CATS.includes(task.category) ? task.category : 'SVD-OS';
    const pri = VALID_PRI.includes(task.priority) ? task.priority : 'medium';
    const sts = VALID_STATUS.includes(task.status) ? task.status : 'todo';
    
    const rowData = [
      task.id,
      task.title,
      cat,
      pri,
      sts,
      task.createdBy || 'SAT',
      task.createdAt,
      task.completedAt || '',
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
      VALID_CATS.includes(t.category) ? t.category : 'SVD-OS',
      VALID_PRI.includes(t.priority) ? t.priority : 'medium',
      VALID_STATUS.includes(t.status) ? t.status : 'todo',
      t.createdBy || 'SAT',
      t.createdAt || getJST(),
      t.completedAt || '',
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
    sheet.appendRow(['id', 'title', 'category', 'priority', 'status', 'createdBy', 'createdAt', 'completedAt', 'notes', 'links', 'isDaily', 'dailyChecked']);
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

// ═══════════════ BOOKMARK OPERATIONS ═══════════════

const BM_HEADERS = ['id', 'name', 'links', 'notes', 'category', 'createdAt'];

function getOrCreateBmSheet(ss) {
  let sheet = ss.getSheetByName(BM_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(BM_SHEET_NAME);
    sheet.appendRow(BM_HEADERS);
    sheet.getRange('A:A').setNumberFormat('@');
    SpreadsheetApp.flush();
  } else {
    // Migrate old schema: rename 'url' → 'links', add 'notes' if missing
    const h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const urlIdx = h.indexOf('url');
    const linksIdx = h.indexOf('links');
    const notesIdx = h.indexOf('notes');
    if (urlIdx >= 0 && linksIdx < 0) {
      sheet.getRange(1, urlIdx + 1).setValue('links');
    }
    if (notesIdx < 0) {
      const catIdx = h.indexOf('category');
      if (catIdx >= 0) {
        sheet.insertColumnAfter(catIdx < 0 ? h.length : (urlIdx >= 0 ? urlIdx + 1 : h.length));
        sheet.getRange(1, (urlIdx >= 0 ? urlIdx + 2 : h.length + 1)).setValue('notes');
      }
    }
  }
  return sheet;
}

function listBookmarks(ss) {
  const sheet = getOrCreateBmSheet(ss);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, bookmarks: [] };
  
  const headers = data[0];
  const hMap = {};
  headers.forEach((h, i) => hMap[h] = i);
  
  const bookmarks = data.slice(1).map(row => ({
    id:        String(row[hMap['id']] || ''),
    name:      String(row[hMap['name']] || ''),
    links:     String(row[hMap['links']] || '[]'),
    notes:     String(row[hMap['notes'] !== undefined ? hMap['notes'] : -1] || ''),
    category:  VALID_CATS.includes(String(row[hMap['category']] || '')) ? String(row[hMap['category']]) : 'SVD-OS',
    createdAt: String(row[hMap['createdAt']] || '')
  })).filter(b => b.id);
  
  return { success: true, bookmarks: bookmarks };
}

function saveBookmark(ss, bm) {
  if (!bm) return { error: 'Bookmark data is required' };
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getOrCreateBmSheet(ss);
    if (!bm.id) bm.id = 'bm-' + new Date().getTime().toString(36) + '-' + Math.random().toString(36).substr(2, 4);
    if (!bm.createdAt) bm.createdAt = getJST();
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const hMap = {};
    headers.forEach((h, i) => hMap[h] = i);
    
    const data = sheet.getDataRange().getValues();
    const idCol = data.map(r => String(r[0]));
    const existingRow = idCol.indexOf(bm.id);
    
    // Build row matching header order
    const rowData = headers.map(h => {
      if (h === 'id') return bm.id;
      if (h === 'name') return bm.name || '';
      if (h === 'links') return typeof bm.links === 'string' ? bm.links : JSON.stringify(bm.links || []);
      if (h === 'notes') return bm.notes || '';
      if (h === 'category') return VALID_CATS.includes(bm.category) ? bm.category : 'SVD-OS';
      if (h === 'createdAt') return bm.createdAt;
      return '';
    });
    
    if (existingRow > 0) {
      sheet.getRange(existingRow + 1, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    
    SpreadsheetApp.flush();
    return { success: true, bookmark: bm };
  } catch (err) {
    return { error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteBookmark(ss, id) {
  if (!id) return { error: 'ID is required' };
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getOrCreateBmSheet(ss);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true, deleted: id };
      }
    }
    
    return { error: 'Bookmark not found: ' + id };
  } catch (err) {
    return { error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function bulkSaveBookmarks(ss, bookmarks) {
  if (!bookmarks || !bookmarks.length) return { error: 'No bookmarks provided' };
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = getOrCreateBmSheet(ss);
    
    // Ensure headers match new schema
    sheet.getRange(1, 1, 1, BM_HEADERS.length).setValues([BM_HEADERS]);
    
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
    
    const rows = bookmarks.map(b => [
      b.id || 'bm-' + new Date().getTime().toString(36) + '-' + Math.random().toString(36).substr(2, 4),
      b.name || '',
      typeof b.links === 'string' ? b.links : JSON.stringify(b.links || []),
      b.notes || '',
      VALID_CATS.includes(b.category) ? b.category : 'SVD-OS',
      b.createdAt || getJST()
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
