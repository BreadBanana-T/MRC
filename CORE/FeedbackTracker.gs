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
  HEADERS: ['Timestamp', 'User', 'Type', 'Message', 'Page', 'Status'],
  TYPES: { suggestion: 'Suggestion', idea: 'Idea', bug: 'Bug' },

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

    var sh = this._sheet();
    var ts = new Date();
    sh.appendRow([ts, who, label, msg, String(page || '').substring(0, 120), 'New']);

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

  // Recent submissions, newest first, so managers can see status and avoid
  // filing the same bug twice.
  getList: function(limit) {
    var sh = this._sheet();
    var last = sh.getLastRow();
    if (last < 2) return JSON.stringify([]);
    var max = limit || 50;
    var n = Math.min(max, last - 1);
    var rows = sh.getRange(last - n + 1, 1, n, this.HEADERS.length).getValues();
    var out = rows.map(function(r) {
      var ts = r[0];
      return {
        ts: (ts instanceof Date) ? Utilities.formatDate(ts, 'America/Toronto', 'yyyy-MM-dd HH:mm') : String(ts || ''),
        user: String(r[1] || ''),
        type: String(r[2] || ''),
        message: String(r[3] || ''),
        page: String(r[4] || ''),
        status: String(r[5] || 'New')
      };
    }).reverse();
    return JSON.stringify(out);
  }
};
