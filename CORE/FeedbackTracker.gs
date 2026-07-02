/**
 * FEEDBACK TRACKER
 * Lets managers send suggestions, ideas and bug reports from inside the portal.
 *
 * Identity is automatic: because the web app is deployed with access "DOMAIN"
 * and the userinfo.email scope is granted, Session.getActiveUser().getEmail()
 * returns the signed-in manager's company email — no login screen needed, as
 * long as everyone is on the same Google Workspace domain.
 *
 * Submissions are appended to the WF_FEEDBACK sheet and (best-effort) emailed
 * to the notify address (Script Property FEEDBACK_NOTIFY_EMAIL, falling back to
 * the deploying owner).
 */
var FeedbackTracker = {

  SHEET: 'WF_FEEDBACK',
  HEADERS: ['Timestamp', 'User', 'Type', 'Message', 'Page', 'Status', 'Updated', 'Completed At', 'Updated By'],
  TYPES: { suggestion: 'Suggestion', idea: 'Idea', bug: 'Bug' },

  // Lifecycle a manager's item moves through. "Understood" = acknowledged/read.
  ALL_STATUSES: ['New', 'Understood', 'In Progress', 'Completed', 'Declined', 'Archived'],
  DEFAULT_ARCHIVE_DAYS: 3, // Completed items auto-archive after this many days.

  // Always-admins (can change status), in ADDITION to the deploying owner and any
  // emails in the FEEDBACK_ADMINS Script Property. Kept here so it works without
  // any setup; edit this list or run setFeedbackAdmins('a@x.com, b@y.com') to
  // change who can mark items complete. Everyone else is read-only.
  OWNER_ADMINS: ['dylan.medina-marchand@telus.com'],

  _sheet: function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(this.SHEET);
    if (!sh) {
      sh = ss.insertSheet(this.SHEET);
      sh.getRange(1, 1, 1, this.HEADERS.length).setValues([this.HEADERS]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  },

  // Bring older sheets (6-column) up to the current header set without touching
  // existing data rows — the new columns simply read blank until first written.
  _ensureColumns: function() {
    var sh = this._sheet();
    var lastCol = sh.getLastColumn();
    var have = (lastCol > 0) ? sh.getRange(1, 1, 1, Math.max(lastCol, this.HEADERS.length)).getValues()[0] : [];
    var need = false;
    for (var i = 0; i < this.HEADERS.length; i++) { if (String(have[i] || '') !== this.HEADERS[i]) { need = true; break; } }
    if (need) sh.getRange(1, 1, 1, this.HEADERS.length).setValues([this.HEADERS]).setFontWeight('bold');
    return sh;
  },

  _archiveDays: function() {
    try {
      var v = parseInt(PropertiesService.getScriptProperties().getProperty('FEEDBACK_ARCHIVE_DAYS'), 10);
      if (v > 0) return v;
    } catch (e) {}
    return this.DEFAULT_ARCHIVE_DAYS;
  },

  // Who may change status. The deploying owner is always an admin; add more via
  // the FEEDBACK_ADMINS Script Property (comma/space separated emails).
  _admins: function() {
    var list = [];
    (this.OWNER_ADMINS || []).forEach(function(s) { s = String(s).trim().toLowerCase(); if (s) list.push(s); });
    try {
      var p = PropertiesService.getScriptProperties().getProperty('FEEDBACK_ADMINS') || '';
      p.split(/[,;\s]+/).forEach(function(s) { s = s.trim().toLowerCase(); if (s) list.push(s); });
    } catch (e) {}
    try { var eff = Session.getEffectiveUser().getEmail(); if (eff) list.push(String(eff).toLowerCase()); } catch (e) {}
    return list;
  },

  isAdmin: function(email) {
    if (!email) { try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {} }
    email = String(email || '').toLowerCase();
    if (!email) return false;
    return this._admins().indexOf(email) !== -1;
  },

  _fmt: function(v) {
    if (v instanceof Date) return Utilities.formatDate(v, 'America/Toronto', 'yyyy-MM-dd HH:mm');
    return String(v == null ? '' : v);
  },

  // Best-effort identity. Active user = the person viewing the app (same-domain
  // Workspace). Effective user is the deployer; used only as a last resort.
  whoAmI: function() {
    var email = '';
    try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
    if (!email) { try { email = Session.getEffectiveUser().getEmail() || ''; } catch (e) {} }
    return JSON.stringify({ email: email, name: this._displayName(email) });
  },

  // "dylan.medina-marchand@telus.com" -> "Dylan Medina-Marchand"
  _displayName: function(email) {
    if (!email) return '';
    var local = String(email).split('@')[0];
    return local.split(/[._]/).filter(function(p) { return p; }).map(function(p) {
      return p.charAt(0).toUpperCase() + p.slice(1);
    }).join(' ');
  },

  submit: function(type, message, page) {
    var msg = String(message == null ? '' : message).trim();
    if (!msg) return JSON.stringify({ ok: false, error: 'Message is empty.' });
    if (msg.length > 5000) msg = msg.substring(0, 5000);
    var label = this.TYPES[String(type).toLowerCase()] || 'Suggestion';

    var email = '';
    try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
    if (!email) { try { email = Session.getEffectiveUser().getEmail() || ''; } catch (e) {} }
    var who = email || 'unknown';

    var sh = this._ensureColumns();
    var ts = new Date();
    sh.appendRow([ts, who, label, msg, String(page || '').substring(0, 120), 'New', '', '', '']);

    this._notify(label, who, msg, page);
    return JSON.stringify({ ok: true, who: who, name: this._displayName(email) });
  },

  _notify: function(label, who, msg, page) {
    try {
      var to = '';
      try { to = PropertiesService.getScriptProperties().getProperty('FEEDBACK_NOTIFY_EMAIL') || ''; } catch (e) {}
      if (!to) { try { to = Session.getEffectiveUser().getEmail() || ''; } catch (e) {} }
      if (!to) return;
      var subject = '[MRC Portal] ' + label + ' from ' + (this._displayName(who) || who);
      var body = label + ' submitted via the MRC Operations Portal.\n\n' +
                 'From: ' + who + '\n' +
                 'Page: ' + (page || '—') + '\n' +
                 'When: ' + Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm') + '\n\n' +
                 '----------------------------------------\n' + msg + '\n';
      MailApp.sendEmail(to, subject, body);
    } catch (e) {
      // Notification is best-effort — never let a mail failure fail the submit.
      Logger.log('[FeedbackTracker] notify failed: ' + e);
    }
  },

  // Recent submissions, newest first. Returns active + archived buckets plus
  // whether the viewer may edit status, so managers see progress at a glance
  // and admins get inline controls. Completed items auto-move to Archived after
  // FEEDBACK_ARCHIVE_DAYS days (flipped lazily here so the sheet stays tidy).
  getList: function(limit) {
    var sh = this._ensureColumns();
    var last = sh.getLastRow();
    var admin = this.isAdmin();
    var base = { items: [], archived: [], isAdmin: admin, statuses: this.ALL_STATUSES, activeCount: 0, archivedCount: 0 };
    if (last < 2) return JSON.stringify(base);

    var readN = Math.min(last - 1, 500);
    var startRow = last - readN + 1;
    var rows = sh.getRange(startRow, 1, readN, this.HEADERS.length).getValues();
    var now = Date.now();
    var archiveMs = this._archiveDays() * 86400000;
    var self = this;
    var flips = [];
    var active = [], archived = [];

    rows.forEach(function(r, i) {
      var rowNum = startRow + i;
      var status = String(r[5] || 'New');
      var completedRaw = r[7];
      var completedMs = (completedRaw instanceof Date) ? completedRaw.getTime() : (completedRaw ? new Date(completedRaw).getTime() : 0);
      var isArch = /archived/i.test(status);
      if (!isArch && /completed/i.test(status) && completedMs && (now - completedMs) > archiveMs) {
        isArch = true; status = 'Archived'; flips.push(rowNum);
      }
      var item = {
        row: rowNum,
        ts: self._fmt(r[0]),
        user: String(r[1] || ''),
        type: String(r[2] || ''),
        message: String(r[3] || ''),
        page: String(r[4] || ''),
        status: status,
        updated: self._fmt(r[6]),
        completedAt: completedMs ? self._fmt(completedRaw) : '',
        updatedBy: String(r[8] || '')
      };
      (isArch ? archived : active).push(item);
    });

    // Persist auto-archive flips (best-effort; never block the read).
    flips.forEach(function(rn) { try { sh.getRange(rn, 6).setValue('Archived'); } catch (e) {} });

    active.reverse(); archived.reverse();
    base.activeCount = active.length;
    base.archivedCount = archived.length;
    base.items = active.slice(0, limit || 50);
    base.archived = archived.slice(0, 100);
    return JSON.stringify(base);
  },

  // Admin-only status change. `row` is the sheet row supplied by getList.
  setStatus: function(row, status) {
    if (!this.isAdmin()) return JSON.stringify({ ok: false, error: 'Not authorized to change status.' });
    row = parseInt(row, 10);
    if (!row || row < 2) return JSON.stringify({ ok: false, error: 'Invalid item.' });

    var norm = null;
    for (var i = 0; i < this.ALL_STATUSES.length; i++) {
      if (this.ALL_STATUSES[i].toLowerCase() === String(status || '').toLowerCase()) norm = this.ALL_STATUSES[i];
    }
    if (!norm) return JSON.stringify({ ok: false, error: 'Unknown status.' });

    var sh = this._ensureColumns();
    if (row > sh.getLastRow()) return JSON.stringify({ ok: false, error: 'Item not found.' });

    var email = '';
    try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
    var now = new Date();
    sh.getRange(row, 6).setValue(norm);   // Status
    sh.getRange(row, 7).setValue(now);    // Updated
    sh.getRange(row, 9).setValue(email);  // Updated By
    if (/completed/i.test(norm)) sh.getRange(row, 8).setValue(now);          // stamp Completed At
    else if (!/archived/i.test(norm)) sh.getRange(row, 8).setValue('');      // reopened -> clear it
    return JSON.stringify({ ok: true, status: norm });
  }
};

// Add/replace the feedback status admins (comma/space separated emails). The
// deploying owner is always an admin even if this is empty.
function setFeedbackAdmins(csv) {
  PropertiesService.getScriptProperties().setProperty('FEEDBACK_ADMINS', String(csv || ''));
  return 'FEEDBACK_ADMINS = ' + (csv || '(empty — owner only)');
}
