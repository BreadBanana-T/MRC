/**
 * PRESENCE & ATTRIBUTION
 *
 * "Who's using the tool right now" (active-now avatars) and "who did the last
 * update" (attribution), with real Google profile photos.
 *
 * Photos: the app executes as the deployer (USER_DEPLOYING), so the OAuth token
 * belongs to the owner — but the owner is a domain member, so we can resolve any
 * colleague's directory photo by email via the People API
 * (people:searchDirectoryPeople, scope directory.readonly). The signed-in
 * viewer is still identified by Session.getActiveUser().getEmail().
 *
 * Requires: People API enabled in the Cloud project + directory.readonly scope
 * (declared in appsscript.json). Falls back gracefully to no photo (the UI then
 * draws an initials avatar).
 */
var Presence = {

  PRESENCE: 'WF_PRESENCE',
  ACTIONS: 'WF_ACTIONS',
  PRESENCE_HEADERS: ['Email', 'Name', 'Photo', 'LastSeen'],
  ACTIONS_HEADERS: ['Action', 'Label', 'Email', 'Name', 'Timestamp'],
  ACTIVE_WINDOW_MIN: 3,

  _sheet: function(name, headers) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  },

  _activeEmail: function() {
    var email = '';
    try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
    if (!email) { try { email = Session.getEffectiveUser().getEmail() || ''; } catch (e) {} }
    return email;
  },

  _displayName: function(email) {
    if (!email) return '';
    var local = String(email).split('@')[0];
    return local.split(/[._]/).filter(function(p) { return p; }).map(function(p) {
      return p.charAt(0).toUpperCase() + p.slice(1);
    }).join(' ');
  },

  // Directory photo URL for an email, cached 6h (success and "none" both cached
  // so we don't re-hit the API for users with no photo).
  _photoForEmail: function(email) {
    if (!email) return '';
    var cache = CacheService.getScriptCache();
    var ckey = 'pfp_' + email.toLowerCase();
    var hit = cache.get(ckey);
    if (hit != null) return hit === '_none_' ? '' : hit;
    var url = this._lookupPhoto(email);
    cache.put(ckey, url || '_none_', 21600);
    return url;
  },

  // Preferred path is the People *advanced service* — it works with the default
  // Apps-Script-managed GCP project once "People API" is added under Services,
  // so there's no need to switch to a standard Cloud project. Falls back to a
  // raw REST call, then to no photo (the UI then draws an initials avatar).
  _lookupPhoto: function(email) {
    var people = null;
    try {
      if (typeof People !== 'undefined' && People.People && People.People.searchDirectoryPeople) {
        var r = People.People.searchDirectoryPeople({
          query: email,
          readMask: 'photos,emailAddresses,names',
          sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'],
          pageSize: 10
        });
        people = r && r.people;
      }
    } catch (e) { Logger.log('[Presence] advanced People failed: ' + e); }

    if (!people) {
      try {
        var token = ScriptApp.getOAuthToken();
        var endpoint = 'https://people.googleapis.com/v1/people:searchDirectoryPeople'
          + '?query=' + encodeURIComponent(email)
          + '&readMask=' + encodeURIComponent('photos,emailAddresses,names')
          + '&sources=DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'
          + '&pageSize=10';
        var resp = UrlFetchApp.fetch(endpoint, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
        if (resp.getResponseCode() === 200) people = (JSON.parse(resp.getContentText()) || {}).people;
        else Logger.log('[Presence] REST lookup ' + resp.getResponseCode() + ': ' + resp.getContentText().substring(0, 200));
      } catch (e2) { Logger.log('[Presence] REST People failed: ' + e2); }
    }

    if (!people || !people.length) return '';
    var match = null;
    for (var i = 0; i < people.length; i++) {
      var emails = (people[i].emailAddresses || []).map(function(e) { return String(e.value || '').toLowerCase(); });
      if (emails.indexOf(email.toLowerCase()) !== -1) { match = people[i]; break; }
    }
    if (!match && people.length === 1) match = people[0];
    if (match && match.photos && match.photos.length) {
      var url = match.photos[0].url || '';
      if (url) url = url.replace(/=s\d+(-c)?$/, '') + '=s96-c';  // crisp small avatar
      return url;
    }
    return '';
  },

  // My identity + photo (for the feedback modal etc.).
  me: function() {
    var email = this._activeEmail();
    return JSON.stringify({ email: email, name: this._displayName(email), photo: email ? this._photoForEmail(email) : '' });
  },

  // Called by the client on a timer. Stamps the viewer as present and returns
  // everyone seen within the active window.
  heartbeat: function() {
    var email = this._activeEmail();
    if (email) {
      var sh = this._sheet(this.PRESENCE, this.PRESENCE_HEADERS);
      var data = sh.getDataRange().getValues();
      var rowIdx = -1;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).toLowerCase() === email.toLowerCase()) { rowIdx = i; break; }
      }
      var photo = (rowIdx > 0) ? String(data[rowIdx][2] || '') : '';
      if (!photo) photo = this._photoForEmail(email);
      var row = [email, this._displayName(email), photo, new Date()];
      if (rowIdx > 0) sh.getRange(rowIdx + 1, 1, 1, this.PRESENCE_HEADERS.length).setValues([row]);
      else sh.appendRow(row);
    }
    return this.getActive();
  },

  getActive: function() {
    var sh = this._sheet(this.PRESENCE, this.PRESENCE_HEADERS);
    var last = sh.getLastRow();
    if (last < 2) return JSON.stringify([]);
    var data = sh.getRange(2, 1, last - 1, this.PRESENCE_HEADERS.length).getValues();
    var cutoff = new Date().getTime() - this.ACTIVE_WINDOW_MIN * 60000;
    var me = this._activeEmail().toLowerCase();
    var out = [];
    data.forEach(function(r) {
      var ls = (r[3] instanceof Date) ? r[3].getTime() : new Date(r[3]).getTime();
      if (isNaN(ls) || ls < cutoff) return;
      out.push({ email: String(r[0]), name: String(r[1] || ''), photo: String(r[2] || ''), lastSeen: ls, isMe: String(r[0]).toLowerCase() === me });
    });
    out.sort(function(a, b) { return b.lastSeen - a.lastSeen; });
    return JSON.stringify(out);
  },

  // Stamp who performed a mutating action. key is stable (one row per action
  // type, upserted); label is human-readable. Never throws.
  recordAction: function(key, label) {
    try {
      var email = this._activeEmail();
      var sh = this._sheet(this.ACTIONS, this.ACTIONS_HEADERS);
      var data = sh.getDataRange().getValues();
      var rowIdx = -1;
      for (var i = 1; i < data.length; i++) { if (String(data[i][0]) === key) { rowIdx = i; break; } }
      var row = [key, label || '', email, this._displayName(email), new Date()];
      if (rowIdx > 0) sh.getRange(rowIdx + 1, 1, 1, this.ACTIONS_HEADERS.length).setValues([row]);
      else sh.appendRow(row);
    } catch (e) { Logger.log('[Presence] recordAction failed: ' + e); }
  },

  getLastActions: function() {
    var sh = this._sheet(this.ACTIONS, this.ACTIONS_HEADERS);
    var last = sh.getLastRow();
    var map = {};
    if (last >= 2) {
      var data = sh.getRange(2, 1, last - 1, this.ACTIONS_HEADERS.length).getValues();
      var self = this;
      data.forEach(function(r) {
        var ts = (r[4] instanceof Date) ? r[4] : new Date(r[4]);
        var email = String(r[2] || '');
        map[String(r[0])] = {
          label: String(r[1] || ''), email: email, name: String(r[3] || ''),
          photo: self._photoForEmail(email),
          ts: isNaN(ts) ? '' : Utilities.formatDate(ts, 'America/Toronto', 'yyyy-MM-dd HH:mm'),
          epoch: isNaN(ts) ? 0 : ts.getTime()
        };
      });
    }
    return JSON.stringify(map);
  }
};
