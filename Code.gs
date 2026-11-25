/**
 * Habit Garden: saves each month's ticks, intention and notes,
 * and brings them back when you change the Month or Year dropdown.
 * Paste this whole file into Extensions > Apps Script, then press Save.
 */
var TRACKER = 'Tracker';
var DATA = 'Data';
var MONTH_RANGE = 'W3:AA3';
var YEAR_RANGE = 'AE3:AG3';
var GRID = 'E15:AI24';        // 10 habits x 31 days
var INTENTION = 'W38';
var NOTES = 'AG38';
var MONTHS = ['January','February','March','April','May','June',
              'July','August','September','October','November','December'];

function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== TRACKER) return;
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return;
  try {
    var r = e.range;
    var monthHit = overlaps_(r, sh.getRange(MONTH_RANGE));
    var yearHit = overlaps_(r, sh.getRange(YEAR_RANGE));
    if (monthHit || yearHit) {
      var newM = monthNum_(sh.getRange(MONTH_RANGE).getValue());
      var newY = Number(sh.getRange(YEAR_RANGE).getValue());
      // Save what was on screen under the month/year that was showing before.
      if (e.oldValue !== undefined) {
        var oldM = monthHit ? monthNum_(e.oldValue) : newM;
        var oldY = yearHit ? Number(e.oldValue) : newY;
        if (oldM && oldY) save_(sh, oldY, oldM);
      }
      if (newM && newY) load_(sh, newY, newM);
      return;
    }
    if (overlaps_(r, sh.getRange(GRID)) || overlaps_(r, sh.getRange(INTENTION)) ||
        overlaps_(r, sh.getRange(NOTES))) {
      var m = monthNum_(sh.getRange(MONTH_RANGE).getValue());
      var y = Number(sh.getRange(YEAR_RANGE).getValue());
      if (m && y) save_(sh, y, m);
    }
  } finally {
    lock.releaseLock();
  }
}

/** Run once from the editor if you want to save the month on screen right now. */
function saveCurrentMonth() {
  var sh = SpreadsheetApp.getActive().getSheetByName(TRACKER);
  save_(sh, Number(sh.getRange(YEAR_RANGE).getValue()), monthNum_(sh.getRange(MONTH_RANGE).getValue()));
}

function monthNum_(v) {
  var i = MONTHS.indexOf(String(v).trim());
  return i < 0 ? 0 : i + 1;
}

function overlaps_(a, b) {
  return a.getRow() <= b.getLastRow() && b.getRow() <= a.getLastRow() &&
         a.getColumn() <= b.getLastColumn() && b.getColumn() <= a.getLastColumn();
}

function dataSheet_() {
  var ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(DATA);
}

/** Returns {start: row, count: n} of this month's block in Data (12 rows), or null. */
function findBlock_(d, y, m) {
  var last = lastRow_(d);
  if (last < 2) return null;
  var keys = d.getRange(2, 1, last - 1, 2).getValues();
  var start = -1, count = 0;
  for (var i = 0; i < keys.length; i++) {
    if (Number(keys[i][0]) === y && Number(keys[i][1]) === m) {
      if (start < 0) start = i + 2;
      count++;
    }
  }
  return start < 0 ? null : {start: start, count: count};
}

function lastRow_(d) {
  var col = d.getRange('A1:A' + d.getMaxRows()).getValues();
  for (var i = col.length - 1; i >= 0; i--) if (col[i][0] !== '') return i + 1;
  return 1;
}

function save_(sh, y, m) {
  var d = dataSheet_();
  var grid = sh.getRange(GRID).getValues();
  var intention = sh.getRange(INTENTION).getValue();
  var notes = sh.getRange(NOTES).getValue();
  var rows = [];
  for (var h = 0; h < 10; h++) rows.push([y, m, h + 1].concat(grid[h].map(function (v) { return v === true; })));
  var blank = []; for (var k = 0; k < 30; k++) blank.push('');
  rows.push([y, m, 'intention', intention].concat(blank));
  rows.push([y, m, 'notes', notes].concat(blank));
  var b = findBlock_(d, y, m);
  if (b && b.count === rows.length && contiguous_(d, b, y, m)) {
    d.getRange(b.start, 1, rows.length, 34).setValues(rows);
  } else {
    if (b) deleteMonth_(d, y, m);
    var at = lastRow_(d) + 1;
    if (at + rows.length - 1 > d.getMaxRows()) d.insertRowsAfter(d.getMaxRows(), 200);
    d.getRange(at, 1, rows.length, 34).setValues(rows);
  }
}

function contiguous_(d, b, y, m) {
  var keys = d.getRange(b.start, 1, b.count, 2).getValues();
  for (var i = 0; i < keys.length; i++)
    if (Number(keys[i][0]) !== y || Number(keys[i][1]) !== m) return false;
  return true;
}

function deleteMonth_(d, y, m) {
  var last = lastRow_(d);
  var keys = d.getRange(2, 1, last - 1, 2).getValues();
  for (var i = keys.length - 1; i >= 0; i--)
    if (Number(keys[i][0]) === y && Number(keys[i][1]) === m) d.deleteRow(i + 2);
}

function load_(sh, y, m) {
  var d = dataSheet_();
  var b = findBlock_(d, y, m);
  var grid = [], intention = '', notes = '';
  for (var h = 0; h < 10; h++) { var row = []; for (var c = 0; c < 31; c++) row.push(false); grid.push(row); }
  if (b) {
    var vals = d.getRange(b.start, 1, b.count, 34).getValues();
    vals.forEach(function (r) {
      if (Number(r[0]) !== y || Number(r[1]) !== m) return;
      if (r[2] === 'intention') intention = r[3];
      else if (r[2] === 'notes') notes = r[3];
      else if (r[2] >= 1 && r[2] <= 10) grid[r[2] - 1] = r.slice(3, 34).map(function (v) { return v === true; });
    });
  }
  sh.getRange(GRID).setValues(grid);
  sh.getRange(INTENTION).setValue(intention);
  sh.getRange(NOTES).setValue(notes);
}
