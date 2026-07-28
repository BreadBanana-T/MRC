/**
 * MODULE: GENESYS ROSTER  (shared skill/level/language decoder)
 *
 * Genesys Cloud is now the source of truth for "who is an MRC monitoring
 * specialist, at what level, with which skills, in which language". The
 * export ("Agent Status Summary") gives one row per agent:
 *
 *   Agent Id | Agent Name | Skills | Logged In | On Queue | Idle | ...
 *   5bcb…f55 | Adam Paquette | English; French; SAM_Monit_Lv_Senior;
 *                              SAM_Monit_Sk_BckMRC; SAM_Monit_Sk_Escalation; …
 *
 * The Skills cell is a single semicolon-separated bag holding three different
 * kinds of token, which this module pulls apart:
 *
 *   LANGUAGE   "English" / "French"  → EN, FR, or BL when both are present
 *   LEVEL      SAM_Monit_Lv_<name>   → Junior = 1, Regular = 2, Senior = 4
 *   SKILL      SAM_Monit_Sk_<name>   → Questions, Video, SmartWear, Escalation,
 *                                      BckMRC (Back MRC), BTS, SME
 *
 * COUNTABILITY (why an agent is or isn't a monitoring specialist)
 *   Genesys carries the whole contact centre, not just MRC, so the roster has
 *   to be filtered before anything is measured against it:
 *     - no tokens at all            → not MRC     (reason 'no-profile')
 *     - language only, no Lv tag    → not MRC     (reason 'no-level')
 *     - SAM_Monit_Lv_Junior         → excluded    (reason 'junior')
 *     - Regular / Senior            → COUNTABLE
 *   `countable` is what Red Flags gates on; every other consumer is free to
 *   show the profile regardless (a Junior still appears on the floor).
 */

var GenesysRoster = {

  // Level tag → numeric ERC level. Junior is deliberately 1 so any existing
  // "level >= 2" rule in the app treats juniors as out of scope for free.
  LEVELS: {
    'junior':  { n: 1, label: 'Junior' },
    'regular': { n: 2, label: 'Regular' },
    'senior':  { n: 4, label: 'Senior' }
  },

  // Skill tag → human label. Unknown tags fall through to the raw suffix
  // with camelCase split, so a brand-new SAM_Monit_Sk_* still renders sanely.
  SKILL_LABELS: {
    'bckmrc':    'Back MRC',
    'bts':       'BTS',
    'escalation':'Escalation',
    'questions': 'Questions',
    'sme':       'SME',
    'smartwear': 'SmartWear',
    'video':     'Video'
  },

  LANG_LABELS: { 'EN': 'English', 'FR': 'French', 'BL': 'Bilingual' },

  /** True when a Skills cell looks like it came from Genesys (has SAM_Monit tags). */
  looksLikeGenesys: function (raw) {
    return /SAM[_\s-]*Monit[_\s-]*(Lv|Sk)[_\s-]/i.test(String(raw || ''));
  },

  /**
   * True when a parsed header row is the Genesys agent export rather than the
   * legacy SharePoint MasterList. Requires an "Agent Id" column plus "Skills";
   * the SharePoint export has employee_name / job_title instead.
   */
  looksLikeGenesysHeader: function (headers) {
    var low = (headers || []).map(function (h) { return String(h || '').toLowerCase().trim(); });
    var hasId = low.some(function (h) { return h === 'agent id' || h === 'agentid' || h === 'user id'; });
    var hasSkills = low.some(function (h) { return h === 'skills'; });
    var hasName = low.some(function (h) { return h.indexOf('agent name') !== -1 || h === 'name' || h === 'user name'; });
    return hasId && hasSkills && hasName;
  },

  _label: function (suffix) {
    var key = String(suffix || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    if (this.SKILL_LABELS[key]) return this.SKILL_LABELS[key];
    // Unknown tag: "SomeNewThing" → "Some New Thing"
    return String(suffix || '').replace(/([a-z0-9])([A-Z])/g, '$1 $2').trim();
  },

  /**
   * Decode a Genesys Skills cell.
   * @return {{
   *   isGenesys:boolean, level:(number|null), levelName:string,
   *   lang:string, langLabel:string, skills:string[], skillText:string,
   *   countable:boolean, reason:string, raw:string
   * }}
   *   reason is '' when countable, otherwise 'no-profile' | 'no-level' | 'junior'.
   */
  parseSkills: function (raw) {
    var self = this;
    var s = String(raw == null ? '' : raw).trim();
    var out = {
      isGenesys: false, level: null, levelName: '', lang: '', langLabel: '',
      skills: [], skillText: '', countable: false, reason: 'no-profile', raw: s
    };
    if (!s) return out;

    var tokens = s.split(/[;\n]+/).map(function (t) { return String(t).trim(); }).filter(Boolean);
    var hasEn = false, hasFr = false;

    tokens.forEach(function (t) {
      var m = t.match(/SAM[_\s-]*Monit[_\s-]*(Lv|Sk)[_\s-]+(.+)$/i);
      if (m) {
        out.isGenesys = true;
        var kind = m[1].toLowerCase();
        var suffix = String(m[2]).trim();
        if (kind === 'lv') {
          var lv = self.LEVELS[suffix.toLowerCase().replace(/[^a-z]/g, '')];
          if (lv) { out.level = lv.n; out.levelName = lv.label; }
          else { out.level = 2; out.levelName = self._label(suffix); }  // unknown tier: treat as Regular
        } else {
          var lbl = self._label(suffix);
          if (lbl && out.skills.indexOf(lbl) === -1) out.skills.push(lbl);
        }
        return;
      }
      if (/^(english|anglais|en)$/i.test(t)) { hasEn = true; return; }
      if (/^(french|fran[cç]ais|fr)$/i.test(t)) { hasFr = true; return; }
    });

    if (hasEn && hasFr) out.lang = 'BL';
    else if (hasFr) out.lang = 'FR';
    else if (hasEn) out.lang = 'EN';
    out.langLabel = this.LANG_LABELS[out.lang] || '';

    out.skills.sort();
    out.skillText = out.skills.join(', ');

    // Countability — see the module header.
    if (!out.isGenesys && !out.lang) out.reason = 'no-profile';
    else if (out.level === null) out.reason = 'no-level';
    else if (out.level < 2) out.reason = 'junior';
    else { out.reason = ''; out.countable = true; }

    return out;
  },

  /**
   * Parse a whole pasted Genesys export into roster rows.
   * Tolerates tab- or comma-separated paste, and the trailing status columns
   * (Logged In / On Queue / …) which we ignore — presence is read live
   * elsewhere, this is a roster sync only.
   *
   * @param {string[][]} rows  already-split cells, header row included
   * @return {{agents:Array, counts:Object}}
   */
  buildRoster: function (rows) {
    var self = this;
    var headers = rows[0] || [];
    var low = headers.map(function (h) { return String(h || '').toLowerCase().trim(); });
    var idxId = low.findIndex(function (h) { return h === 'agent id' || h === 'agentid' || h === 'user id'; });
    var idxName = low.findIndex(function (h) { return h.indexOf('agent name') !== -1 || h === 'name' || h === 'user name'; });
    var idxSkills = low.findIndex(function (h) { return h === 'skills'; });

    var agents = [];
    var counts = { total: 0, countable: 0, junior: 0, noProfile: 0, noLevel: 0, senior: 0, regular: 0, bilingual: 0, english: 0, french: 0 };
    var seen = {};

    for (var i = 1; i < rows.length; i++) {
      var cols = rows[i] || [];
      var name = String(cols[idxName] || '').replace(/(^"|"$)/g, '').trim();
      if (!name) continue;
      if (/^(total|grand total)/i.test(name)) continue;

      var key = (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey(name) : name.toLowerCase();
      if (seen[key]) continue;
      seen[key] = true;

      var p = self.parseSkills(cols[idxSkills]);
      counts.total++;
      if (p.countable) {
        counts.countable++;
        if (p.level >= 4) counts.senior++; else counts.regular++;
        if (p.lang === 'BL') counts.bilingual++;
        else if (p.lang === 'FR') counts.french++;
        else if (p.lang === 'EN') counts.english++;
      } else if (p.reason === 'junior') counts.junior++;
      else if (p.reason === 'no-level') counts.noLevel++;
      else counts.noProfile++;

      agents.push({
        name: name,
        id: String(cols[idxId] || '').trim(),
        profile: p
      });
    }

    return { agents: agents, counts: counts };
  }
};
