param(
    [string]$SourceHtml = (Join-Path $PSScriptRoot '..\MacroGuide.html'),
    [string]$OutDir = (Join-Path $PSScriptRoot '..\docs')
)

$ErrorActionPreference = 'Stop'

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$sourcePath = [System.IO.Path]::GetFullPath($SourceHtml)
$outPath = [System.IO.Path]::GetFullPath($OutDir)
$outFile = Join-Path $outPath 'index.html'

if (-not (Test-Path $sourcePath)) {
    throw "Source HTML not found: $sourcePath"
}

New-Item -ItemType Directory -Path $outPath -Force | Out-Null

$html = [System.IO.File]::ReadAllText($sourcePath, [System.Text.Encoding]::UTF8)

$previewStyle = @'
<style>
  .github-preview-banner {
    position: sticky;
    top: 0;
    z-index: 2200;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 10px;
    padding: 10px 18px;
    background: #0f766e;
    color: #ecfeff;
    border-bottom: 1px solid rgba(255,255,255,0.18);
    font-size: 0.9rem;
    font-weight: 650;
    text-align: center;
  }
  .github-preview-banner span:last-child {
    color: rgba(236,254,255,0.86);
    font-weight: 500;
  }
  .github-preview #adminRestart {
    display: none !important;
  }
  .github-preview .admin-usage {
    right: 80px;
  }
</style>
'@

$previewBanner = @'
<div class="github-preview-banner">
  <strong>GitHub Preview</strong>
  <span>Demo data only. Changes stay in this browser and do not touch the live Y: Macro Guide.</span>
</div>
'@

$previewShim = @'
<script>
(() => {
  const STORAGE_KEY = 'macroGuideGithubPreviewData.v2';
  const CRASH_RULES_VERSION = '2026-05-donut-duty-v4';
  const originalFetch = window.fetch.bind(window);
  const today = new Date().toISOString().slice(0, 10);

  document.documentElement.classList.add('github-preview');
  if (!sessionStorage.getItem('winUser')) sessionStorage.setItem('winUser', 'github-preview');
  if (!localStorage.getItem('mgUserName')) localStorage.setItem('mgUserName', 'Preview User');
  sessionStorage.removeItem('mgAdminUnlocked');
  sessionStorage.removeItem('nmtAdmin');

  function clone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function defaultPokeTargets() {
    return [
      {
        label: 'Dylan Shank',
        shortName: 'Shank',
        windowsUsers: ['dshank'],
        names: ['Dylan Shank'],
        aliases: ['shank'],
        enabled: true
      },
      {
        label: 'Jason Gagnon',
        shortName: 'Jason',
        windowsUsers: ['jgagnon'],
        names: ['Jason Gagnon'],
        aliases: ['jason', 'gagnon'],
        enabled: true
      },
      {
        label: 'Krupal Patel',
        shortName: 'Krupal',
        windowsUsers: ['kpatel'],
        names: ['Krupal Patel'],
        aliases: ['krupal', 'patel'],
        enabled: true
      },
      {
        label: 'Paul Lemelin',
        shortName: 'Paul',
        windowsUsers: ['plemelin'],
        names: ['Paul Lemelin'],
        aliases: ['paul', 'lemelin'],
        enabled: true
      }
    ];
  }

  function seedData() {
    return {
      changelog: [
        {
          id: 4,
          macroId: 'all',
          macroName: 'All Macros',
          macroIcon: '⚙️',
          macroColor: '#4f8ef7',
          date: '2026-04-16T09:30:00.000Z',
          type: 'feature',
          description: 'Added the browser-based GitHub preview with demo request data.',
          author: 'Dylan'
        },
        {
          id: 3,
          macroId: 'cadlink',
          macroName: 'Prepare for CADLink',
          macroIcon: '⚡',
          macroColor: '#a855f7',
          date: '2026-04-15T14:15:00.000Z',
          type: 'improve',
          description: 'Improved request workflow instructions for voting, poking, and comments.',
          author: 'Dylan'
        },
        {
          id: 2,
          macroId: 'pdmpdf',
          macroName: 'PDM PDF Generator',
          macroIcon: '📑',
          macroColor: '#06b6d4',
          date: '2026-04-12T18:10:00.000Z',
          type: 'fix',
          description: 'Polished launcher readiness checks so the browser waits until the local service is ready.',
          author: 'Dylan'
        },
        {
          id: 1,
          macroId: 'pdmdxf',
          macroName: 'PDM DXF Exporter',
          macroIcon: '✂️',
          macroColor: '#e879a0',
          date: '2026-04-08T11:45:00.000Z',
          type: 'doc',
          description: 'Updated guide copy for common export and check-in edge cases.',
          author: 'Dylan'
        }
      ],
      tickets: [
        {
          id: 103,
          macroId: 'cadlink',
          macroName: 'Prepare for CADLink',
          macroIcon: '⚡',
          macroColor: '#a855f7',
          title: 'Revision Handling',
          description: 'Check whether there is a drawing revision that should update the model revision before preparing CADLink data.',
          type: 'bug',
          priority: 'low',
          status: 'in-progress',
          date: '2026-04-08T10:00:00.000Z',
          createdBy: 'Paul',
          windowsUser: 'plemelin',
          votes: 0,
          voters: [],
          pokes: 0,
          pokers: [],
          comments: [],
          attachments: []
        },
        {
          id: 102,
          macroId: 'all',
          macroName: 'General',
          macroIcon: '⚙️',
          macroColor: '#4f8ef7',
          title: 'comment here',
          description: 'comment something',
          type: 'bug',
          priority: 'critical',
          status: 'in-progress',
          date: '2026-03-19T16:20:00.000Z',
          createdBy: 'Dylan',
          windowsUser: 'dlebel',
          votes: 1,
          voters: ['github-preview'],
          pokes: 0,
          pokers: [],
          comments: [
            {
              windowsUser: 'github-preview',
              displayName: 'Preview User',
              text: 'This is a sample comment in the GitHub preview.',
              gifUrl: '',
              date: '2026-03-19'
            }
          ],
          attachments: []
        },
        {
          id: 101,
          macroId: 'all',
          macroName: 'General',
          macroIcon: '⚙️',
          macroColor: '#4f8ef7',
          title: 'make better look',
          description: 'looks like this in classic',
          type: 'change',
          priority: 'high',
          status: 'done',
          date: '2026-03-19T12:10:00.000Z',
          createdBy: 'Paul',
          windowsUser: 'plemelin',
          votes: 3,
          voters: ['dlebel', 'jason', 'github-preview'],
          pokes: 1,
          pokers: ['dlebel'],
          comments: [],
          attachments: []
        }
      ],
      crashes: [
        {
          id: 1,
          user: 'Paul Lemelin',
          severity: 'major',
          timestamp: '2026-04-03T14:20:00.000Z',
          createdBy: 'Paul Lemelin',
          windowsUser: 'plemelin'
        },
        {
          id: 2,
          user: 'Paul Lemelin',
          severity: 'minor',
          timestamp: '2026-04-09T17:35:00.000Z',
          createdBy: 'Paul Lemelin',
          windowsUser: 'plemelin'
        },
        {
          id: 3,
          user: 'Krupal Patel',
          severity: 'minor',
          timestamp: '2026-04-12T13:10:00.000Z',
          createdBy: 'Krupal Patel',
          windowsUser: 'kpatel'
        },
        {
          id: 4,
          user: 'Paul Lemelin',
          severity: 'catastrophic',
          timestamp: '2026-04-24T19:45:00.000Z',
          createdBy: 'Paul Lemelin',
          windowsUser: 'plemelin'
        },
        {
          id: 5,
          user: 'Ayugma Acharya',
          severity: 'major',
          timestamp: '2026-05-01T14:10:00.000Z',
          createdBy: 'Ayugma Acharya',
          windowsUser: 'aacharya'
        },
        {
          id: 6,
          user: 'Ayugma Acharya',
          severity: 'minor',
          timestamp: '2026-05-02T16:05:00.000Z',
          createdBy: 'Ayugma Acharya',
          windowsUser: 'aacharya'
        }
      ],
      donutStatuses: [
        {
          monthKey: '2026-04',
          paid: false,
          updatedAt: '2026-05-01T13:00:00.000Z',
          updatedBy: 'dlebel',
          updatedByName: 'Dylan',
          paidAt: ''
        }
      ],
      usage: [
        { timestamp: '2026-04-16 09:12:40', windowsUser: 'dlebel', machine: 'D-1077', page: 'tickets', action: 'OPEN' },
        { timestamp: '2026-04-16 09:18:21', windowsUser: 'plemelin', machine: 'P-2044', page: 'macros', action: 'OPEN' },
        { timestamp: '2026-04-16 09:35:04', windowsUser: 'github-preview', machine: 'GitHub Pages', page: 'tickets', action: 'OPEN' }
      ],
      captchaQueue: [],
      pokeResets: [],
      pokeTargets: defaultPokeTargets(),
      crashConsents: [],
      tiePollVotes: [
        {
          pollId: 'crash-tie-break-2026-04',
          windowsUser: 'aacharya',
          displayName: 'Ayugma Acharya',
          optionId: 'split',
          votedAt: '2026-04-30T18:09:23.3912450Z'
        },
        {
          pollId: 'crash-tie-break-2026-04',
          windowsUser: 'dlebel',
          displayName: 'Dylan',
          optionId: 'weeks',
          votedAt: '2026-04-30T18:13:18.9283205Z'
        },
        {
          pollId: 'crash-tie-break-2026-04',
          windowsUser: 'dshank',
          displayName: 'Other Dylan',
          optionId: 'weeks',
          votedAt: '2026-04-30T18:22:39.6579588Z'
        },
        {
          pollId: 'crash-tie-break-2026-04',
          windowsUser: 'barrowsmith',
          displayName: 'Brandon Arrowsmith',
          optionId: 'weeks',
          votedAt: '2026-04-30T18:29:17.5788815Z'
        }
      ],
      version: 'github-preview-2026-04-16'
    };
  }

  function loadData() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const data = JSON.parse(raw);
        data.pokeResets = Array.isArray(data.pokeResets) ? data.pokeResets : [];
        data.pokeTargets = Array.isArray(data.pokeTargets) ? data.pokeTargets : defaultPokeTargets();
        data.tiePollVotes = Array.isArray(data.tiePollVotes) ? data.tiePollVotes : [];
        data.captchaQueue = Array.isArray(data.captchaQueue) ? data.captchaQueue : [];
        data.crashes = Array.isArray(data.crashes) ? data.crashes : [];
        data.donutStatuses = Array.isArray(data.donutStatuses) ? data.donutStatuses : [];
        data.crashConsents = Array.isArray(data.crashConsents) ? data.crashConsents : [];
        return data;
      }
    } catch {}
    const seeded = seedData();
    saveData(seeded);
    return seeded;
  }

  function saveData(data) {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(data)); } catch {}
  }

  function parseBody(init) {
    if (!init || init.body == null) return {};
    if (typeof init.body === 'string') {
      try { return JSON.parse(init.body || '{}'); } catch { return {}; }
    }
    return {};
  }

  function jsonResponse(body, status = 200) {
    return Promise.resolve(new Response(JSON.stringify(body), {
      status,
      headers: { 'Content-Type': 'application/json' }
    }));
  }

  function textResponse(body, status = 200) {
    return Promise.resolve(new Response(body, {
      status,
      headers: { 'Content-Type': 'text/html; charset=utf-8' }
    }));
  }

  function byId(entries, id) {
    return entries.find(t => String(t.id) === String(id));
  }

  function maxId(entries) {
    return entries.reduce((max, entry) => Math.max(max, Number(entry.id) || 0), 0);
  }

  function previewMonthKey(value = new Date()) {
    const d = value instanceof Date ? value : new Date(value);
    if (Number.isNaN(d.getTime())) return '';
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  }

  function previewDisplayKey(value) {
    return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
  }

  function previewCurrentUser() {
    return String(sessionStorage.getItem('winUser') || 'github-preview').trim().toLowerCase();
  }

  function previewDisplayName(fallback = '') {
    return String(fallback || localStorage.getItem('mgUserName') || sessionStorage.getItem('winUserDisplay') || previewCurrentUser()).trim();
  }

  function previewCrashConsentState(data, displayName = '') {
    data.crashConsents = Array.isArray(data.crashConsents) ? data.crashConsents : [];
    const windowsUser = previewCurrentUser();
    const currentMonthKey = previewMonthKey();
    const row = data.crashConsents.find(r => String(r.windowsUser || '').trim().toLowerCase() === windowsUser);
    const accepted = !!(row && row.rulesVersion === CRASH_RULES_VERSION && row.accepted);
    const declinedThisMonth = !!(!accepted && row && row.declinedMonthKey === currentMonthKey);
    return {
      ok: true,
      windowsUser,
      displayName: previewDisplayName(displayName || row?.displayName || ''),
      rulesVersion: CRASH_RULES_VERSION,
      accepted,
      acceptedAt: row?.acceptedAt || '',
      declinedThisMonth,
      declinedMonthKey: row?.declinedMonthKey || '',
      declinedAt: row?.declinedAt || '',
      deletedCrashCount: 0,
      excludedCrashCount: Number(row?.excludedCrashCount || 0),
      backedUpCrashCount: Array.isArray(row?.backedUpCrashes) ? row.backedUpCrashes.length : 0,
      currentMonthKey,
      canLog: accepted && !declinedThisMonth,
      needsDecision: !accepted && !declinedThisMonth
    };
  }

  function previewAddConsentParticipant(map, windowsUser, displayName = '', source = '') {
    const win = String(windowsUser || '').trim().toLowerCase();
    const display = String(displayName || '').trim();
    if (!win && !display) return;
    const displayKey = previewDisplayKey(display);
    const nameOnlyKey = displayKey ? 'name:' + displayKey : '';
    let matchingKey = '';
    if (displayKey) {
      for (const [existingKey, existing] of map.entries()) {
        if (previewDisplayKey(existing.displayName) === displayKey) {
          matchingKey = existingKey;
          break;
        }
      }
    }

    let key = '';
    if (win && map.has(win)) {
      key = win;
    } else if (win && matchingKey) {
      const existing = map.get(matchingKey);
      if (matchingKey === nameOnlyKey || !existing.windowsUser) {
        map.delete(matchingKey);
        map.set(win, existing);
        key = win;
      } else {
        key = matchingKey;
      }
    } else if (win) {
      key = win;
    } else if (matchingKey) {
      key = matchingKey;
    } else {
      key = nameOnlyKey;
    }

    if (!map.has(key)) {
      map.set(key, {
        windowsUser: win,
        displayName: display || win,
        sources: []
      });
    }
    const row = map.get(key);
    if (win && !row.windowsUser) row.windowsUser = win;
    if (display && (!row.displayName || row.displayName === row.windowsUser)) row.displayName = display;
    if (source && !row.sources.includes(source)) row.sources.push(source);
  }

  function previewCrashConsentRoster(data) {
    data.crashConsents = Array.isArray(data.crashConsents) ? data.crashConsents : [];
    data.crashes = Array.isArray(data.crashes) ? data.crashes : [];
    const currentMonthKey = previewMonthKey();
    const map = new Map();
    const consentByUser = new Map();

    for (const row of data.crashConsents) {
      const win = String(row?.windowsUser || '').trim().toLowerCase();
      if (!win) continue;
      consentByUser.set(win, row);
      previewAddConsentParticipant(map, win, row.displayName || win, 'terms');
    }

    const targets = Array.isArray(data.pokeTargets) ? data.pokeTargets : defaultPokeTargets();
    for (const target of targets) {
      if (!target || target.enabled === false) continue;
      const display = String(target.label || target.shortName || (Array.isArray(target.names) ? target.names[0] : '') || '').trim();
      (Array.isArray(target.windowsUsers) ? target.windowsUsers : []).forEach(wu => {
        previewAddConsentParticipant(map, wu, display, 'participant');
      });
    }

    for (const crash of data.crashes) {
      previewAddConsentParticipant(
        map,
        crash?.participantWindowsUser || crash?.windowsUser || '',
        crash?.user || crash?.createdBy || '',
        'crash'
      );
    }

    const participants = Array.from(map.values()).map(base => {
      const win = String(base.windowsUser || '').trim().toLowerCase();
      const row = win ? consentByUser.get(win) : null;
      const rulesVersion = row?.rulesVersion || '';
      const accepted = !!(row && rulesVersion === CRASH_RULES_VERSION && row.accepted);
      const declinedThisMonth = !!(!accepted && row && row.declinedMonthKey === currentMonthKey);
      let status = 'not_started';
      if (accepted) status = 'accepted';
      else if (declinedThisMonth) status = 'spectator';
      else if (row) status = 'needs_acceptance';
      return {
        displayName: String(row?.displayName || base.displayName || win || 'Unknown participant').trim(),
        windowsUser: win,
        status,
        accepted,
        acceptedAt: row?.acceptedAt || '',
        declinedThisMonth,
        declinedAt: row?.declinedAt || '',
        rulesVersion,
        currentRulesVersion: CRASH_RULES_VERSION,
        source: base.sources.join(', ')
      };
    }).sort((a, b) => {
      const order = { accepted: 0, spectator: 1, needs_acceptance: 2, not_started: 3 };
      return (order[a.status] ?? 4) - (order[b.status] ?? 4) ||
        a.displayName.localeCompare(b.displayName);
    });

    return {
      ok: true,
      rulesVersion: CRASH_RULES_VERSION,
      currentMonthKey,
      generatedAt: new Date().toISOString(),
      participants,
      totalCount: participants.length,
      acceptedCount: participants.filter(p => p.status === 'accepted').length,
      spectatorCount: participants.filter(p => p.status === 'spectator').length,
      needsDecisionCount: participants.filter(p => p.status === 'needs_acceptance' || p.status === 'not_started').length,
      source: 'GitHub preview local browser storage'
    };
  }

  function previewSaveCrashConsent(data, accepted, displayName = '', excludedCrashCount = 0, backedUpCrashes = []) {
    data.crashConsents = Array.isArray(data.crashConsents) ? data.crashConsents : [];
    // Players are free to flip between Spectator and Active mid-month via "Join now".
    const windowsUser = previewCurrentUser();
    const stamp = new Date().toISOString();
    let row = data.crashConsents.find(r => String(r.windowsUser || '').trim().toLowerCase() === windowsUser);
    if (!row) {
      row = { windowsUser };
      data.crashConsents.push(row);
    }
    row.displayName = previewDisplayName(displayName);
    row.rulesVersion = CRASH_RULES_VERSION;
    row.accepted = !!accepted;
    row.acceptedAt = accepted ? stamp : '';
    row.declinedMonthKey = accepted ? '' : previewMonthKey();
    row.declinedAt = accepted ? '' : stamp;
    row.deletedCrashCount = 0;
    row.excludedCrashCount = accepted ? 0 : excludedCrashCount;
    row.backedUpCrashes = accepted ? [] : clone(backedUpCrashes);
    saveData(data);
    return previewCrashConsentState(data, row.displayName);
  }

  function previewDeclineCrashConsent(data, displayName = '') {
    const windowsUser = previewCurrentUser();
    const displayKey = previewDisplayKey(previewDisplayName(displayName));
    const monthKey = previewMonthKey();
    data.crashes = Array.isArray(data.crashes) ? data.crashes : [];
    const backedUpCrashes = data.crashes.filter(c => {
      const inMonth = previewMonthKey(c.timestamp) === monthKey;
      const sameUser = String(c.participantWindowsUser || c.windowsUser || '').trim().toLowerCase() === windowsUser ||
        previewDisplayKey(c.user) === displayKey;
      return inMonth && sameUser;
    });
    const state = previewSaveCrashConsent(data, false, displayName, backedUpCrashes.length, backedUpCrashes);
    return { ...state, entries: clone(data.crashes) };
  }

  function usageSummary(entries) {
    const seen = new Set();
    const unique = [];
    entries.forEach(entry => {
      const user = entry.windowsUser || '';
      if (user && !seen.has(user)) {
        seen.add(user);
        unique.push(user);
      }
    });
    return {
      source: 'GitHub preview local browser storage',
      unique,
      entries: entries.slice().reverse().slice(0, 300)
    };
  }

  function normalizePreviewUser(value) {
    return String(value || '').trim().toLowerCase().replace(/^.*\\/, '');
  }

  function addKnownCaptchaTarget(map, windowsUser, displayName = '', details = {}) {
    const wu = normalizePreviewUser(windowsUser);
    if (!wu || wu === 'dlebel') return;
    const label = String(displayName || '').trim() || wu;
    if (!map.has(wu)) {
      map.set(wu, {
        label,
        shortName: label.split(/\s+/)[0] || wu,
        windowsUsers: [wu],
        names: [],
        aliases: [],
        enabled: true,
        machine: '',
        lastSeen: '',
        source: [],
        configured: false
      });
    }
    const row = map.get(wu);
    if (details.configured || row.label === wu) {
      row.label = label;
      row.shortName = label.split(/\s+/)[0] || wu;
    }
    if (label && !row.names.includes(label)) row.names.push(label);
    if (details.machine) row.machine = String(details.machine);
    if (details.lastSeen && (!row.lastSeen || String(details.lastSeen) > row.lastSeen)) row.lastSeen = String(details.lastSeen);
    if (details.source && !row.source.includes(details.source)) row.source.push(details.source);
    if (details.configured) row.configured = true;
  }

  function previewKnownCaptchaTargets(data) {
    const map = new Map();
    const targets = Array.isArray(data.pokeTargets) ? data.pokeTargets : defaultPokeTargets();
    targets.forEach(t => {
      if (!t || t.enabled === false) return;
      const wins = Array.isArray(t.windowsUsers) ? t.windowsUsers : [];
      wins.forEach(wu => {
        addKnownCaptchaTarget(map, wu, t.label || wu, { source: 'target', configured: true });
        const row = map.get(normalizePreviewUser(wu));
        if (!row) return;
        (Array.isArray(t.aliases) ? t.aliases : []).forEach(a => {
          const alias = normalizePreviewUser(a);
          if (alias && !row.aliases.includes(alias)) row.aliases.push(alias);
        });
        (Array.isArray(t.names) ? t.names : []).forEach(n => {
          const name = String(n || '').trim();
          if (name && !row.names.includes(name)) row.names.push(name);
        });
      });
    });

    (Array.isArray(data.usage) ? data.usage : []).forEach(u => {
      addKnownCaptchaTarget(map, u.windowsUser, '', { machine: u.machine || '', lastSeen: u.timestamp || '', source: 'usage' });
    });
    (Array.isArray(data.tickets) ? data.tickets : []).forEach(t => {
      addKnownCaptchaTarget(map, t.windowsUser, t.createdBy || '', { lastSeen: t.date || '', source: 'ticket' });
      (Array.isArray(t.voters) ? t.voters : []).forEach(v => addKnownCaptchaTarget(map, v, '', { source: 'vote' }));
      (Array.isArray(t.pokers) ? t.pokers : []).forEach(p => addKnownCaptchaTarget(map, p, '', { source: 'poke' }));
      (Array.isArray(t.comments) ? t.comments : []).forEach(c => {
        addKnownCaptchaTarget(map, c.windowsUser, c.displayName || '', { lastSeen: c.date || '', source: 'comment' });
      });
    });
    (Array.isArray(data.crashes) ? data.crashes : []).forEach(c => {
      addKnownCaptchaTarget(map, c.windowsUser, c.createdBy || c.user || '', { lastSeen: c.timestamp || '', source: 'crash' });
    });
    addKnownCaptchaTarget(map, sessionStorage.getItem('winUser') || 'github-preview', localStorage.getItem('mgUserName') || 'Preview User', {
      machine: 'GitHub Pages',
      lastSeen: new Date().toISOString(),
      source: 'presence'
    });

    return Array.from(map.values())
      .map(row => ({ ...row, source: row.source.join(', ') }))
      .sort((a, b) => Number(b.configured) - Number(a.configured) || String(a.label).localeCompare(String(b.label)));
  }

  const TIE_POLL_ID = 'crash-tie-break-2026-04';
  const TIE_POLL_CLOSES_AT = new Date(2026, 3, 30, 16, 30, 0).toISOString();
  const TIE_POLL_CLOSED_LABEL = '4:30 PM';
  const TIE_POLL_OPTIONS = [
    {
      id: 'wheel',
      icon: '🎡',
      title: 'Spin a wheel',
      description: 'Put everyone from the tie on a wheel. Whoever it lands on brings donuts.'
    },
    {
      id: 'split',
      icon: '🍩',
      title: 'Split snack duty',
      description: 'One tied person brings donuts. The other tied person brings cookies or muffins.'
    },
    {
      id: 'weeks',
      icon: '\uD83D\uDCC5',
      title: 'One week each',
      description: 'One tied person brings donuts one week. The other tied person gets the next week.'
    }
  ];

  function tiePollState(data) {
    data.tiePollVotes = Array.isArray(data.tiePollVotes) ? data.tiePollVotes : [];
    const valid = new Set(TIE_POLL_OPTIONS.map(o => o.id));
    const byUser = new Map();
    for (const vote of data.tiePollVotes) {
      if (!vote || vote.pollId !== TIE_POLL_ID || !valid.has(vote.optionId)) continue;
      const user = String(vote.windowsUser || '').trim().toLowerCase();
      if (user) byUser.set(user, vote);
    }
    const totals = Object.fromEntries(TIE_POLL_OPTIONS.map(o => [o.id, 0]));
    for (const vote of byUser.values()) totals[vote.optionId] = (totals[vote.optionId] || 0) + 1;
    const currentUser = (sessionStorage.getItem('winUser') || 'github-preview').trim().toLowerCase();
    return {
      pollId: TIE_POLL_ID,
      question: 'What should happen when crash duty ends in a tie?',
      options: clone(TIE_POLL_OPTIONS),
      totals,
      totalVotes: byUser.size,
      userVote: byUser.get(currentUser)?.optionId || null,
      closesAt: TIE_POLL_CLOSES_AT,
      closedLabel: TIE_POLL_CLOSED_LABEL,
      isClosed: Date.now() >= Date.parse(TIE_POLL_CLOSES_AT),
      source: 'GitHub preview local browser storage'
    };
  }

  function updateTicket(path, method, init, data, searchParams) {
    const match = path.match(/^\/api\/tickets\/(\d+)\/([^/]+)$/);
    if (!match) return null;
    const id = match[1];
    const action = match[2];
    const body = parseBody(init);
    const ticket = byId(data.tickets, id);

    if (action === 'delete' && method === 'PATCH') {
      data.tickets = data.tickets.filter(t => String(t.id) !== String(id));
      saveData(data);
      return { ok: true, entries: clone(data.tickets) };
    }

    if (!ticket) return { error: 'Ticket not found.' };

    if (action === 'status' && method === 'PATCH') {
      ticket.status = body.status || ticket.status;
    } else if (action === 'vote' && method === 'PATCH') {
      ticket.votes = (Number(ticket.votes) || 0) + 1;
      ticket.voters = Array.isArray(ticket.voters) ? ticket.voters : [];
      const voter = body.windowsUser || sessionStorage.getItem('winUser') || 'github-preview';
      if (!ticket.voters.includes(voter)) ticket.voters.push(voter);
    } else if (action === 'poke' && method === 'PATCH') {
      ticket.pokes = (Number(ticket.pokes) || 0) + 1;
      ticket.pokers = Array.isArray(ticket.pokers) ? ticket.pokers : [];
      const poker = body.windowsUser || sessionStorage.getItem('winUser') || 'github-preview';
      if (!ticket.pokers.includes(poker)) ticket.pokers.push(poker);
      ticket.lastPoke = today;
    } else if (action === 'cancel' && method === 'PATCH') {
      ticket.status = 'canceled';
      ticket.cancelReason = body.reason || 'Canceled in preview';
      ticket.canceledBy = body.canceledBy || localStorage.getItem('mgUserName') || 'Preview User';
      ticket.canceledDate = today;
    } else if (action === 'edit' && method === 'PATCH') {
      ticket.title = body.title || ticket.title;
      ticket.description = body.description || ticket.description;
      ticket.priority = body.priority || ticket.priority;
      ticket.type = body.type || ticket.type;
      ticket.lastEdited = today;
    } else if (action === 'comment' && method === 'POST') {
      ticket.comments = Array.isArray(ticket.comments) ? ticket.comments : [];
      ticket.comments.push({
        windowsUser: body.windowsUser || sessionStorage.getItem('winUser') || 'github-preview',
        displayName: body.displayName || localStorage.getItem('mgUserName') || 'Preview User',
        text: body.text || '',
        gifUrl: body.gifUrl || '',
        date: today
      });
    } else if (action === 'comment' && method === 'PATCH') {
      ticket.comments = Array.isArray(ticket.comments) ? ticket.comments : [];
      const user = body.windowsUser || sessionStorage.getItem('winUser') || 'github-preview';
      const comment = ticket.comments.find(c => c.windowsUser === user);
      if (comment) {
        comment.text = body.text || '';
        comment.gifUrl = body.gifUrl || '';
        comment.editedDate = today;
      }
    } else if (action === 'comment' && method === 'DELETE') {
      ticket.comments = (ticket.comments || []).filter(c => c.windowsUser !== body.windowsUser);
    } else if (action === 'attach' && method === 'POST') {
      const name = (searchParams && searchParams.get('name')) || 'preview-file.txt';
      ticket.attachments = Array.isArray(ticket.attachments) ? ticket.attachments : [];
      if (!ticket.attachments.includes(name)) ticket.attachments.push(name);
      saveData(data);
      return { ok: true, name, size: 0 };
    } else {
      return null;
    }

    saveData(data);
    return { ok: true, entries: clone(data.tickets) };
  }

  window.fetch = function previewFetch(input, init = {}) {
    const requestUrl = typeof input === 'string' ? input : input.url;
    const url = new URL(requestUrl, window.location.href);
    const method = (init.method || (typeof input === 'object' && input.method) || 'GET').toUpperCase();

    if (url.origin === window.location.origin && url.pathname === '/') {
      return textResponse('<!doctype html><title>Macro Guide Preview</title>');
    }

    if (url.origin !== window.location.origin || !url.pathname.startsWith('/api/')) {
      return originalFetch(input, init);
    }

    const data = loadData();
    const path = url.pathname;

    if (method === 'GET' && path === '/api/changelog') return jsonResponse(clone(data.changelog));
    if (method === 'GET' && path === '/api/tickets') return jsonResponse(clone(data.tickets));
    if (method === 'GET' && path === '/api/crashes') {
      data.crashes = Array.isArray(data.crashes) ? data.crashes : [];
      return jsonResponse(clone(data.crashes));
    }
    if (method === 'GET' && path === '/api/crash-donut-status') {
      data.donutStatuses = Array.isArray(data.donutStatuses) ? data.donutStatuses : [];
      return jsonResponse(clone(data.donutStatuses));
    }
    if (method === 'GET' && path === '/api/crash-consent') {
      return jsonResponse(previewCrashConsentRoster(data));
    }
    if (method === 'GET' && path === '/api/crash-consent/me') {
      return jsonResponse(previewCrashConsentState(data, url.searchParams.get('displayName') || ''));
    }
    if (method === 'GET' && path === '/api/whoami') return jsonResponse({
      windowsUser: sessionStorage.getItem('winUser') || 'github-preview',
      isAdmin: true,
      machine: 'GitHub Pages'
    });
    if (method === 'GET' && path === '/api/version') return jsonResponse({ version: data.version });
    if (method === 'GET' && path === '/api/usage') return jsonResponse(usageSummary(data.usage || []));
    if (method === 'GET' && path === '/api/restart') return jsonResponse({ ok: true, preview: true });
    if (method === 'PATCH' && path.match(/^\/api\/crash-donut-status\/\d{4}-\d{2}$/)) {
      const monthKey = path.split('/').pop();
      const body = parseBody(init);
      data.donutStatuses = Array.isArray(data.donutStatuses) ? data.donutStatuses : [];
      const paid = !!body.paid;
      const stamp = new Date().toISOString();
      const existing = data.donutStatuses.find(s => s.monthKey === monthKey);
      if (existing) {
        existing.paid = paid;
        existing.updatedAt = stamp;
        existing.updatedBy = sessionStorage.getItem('winUser') || 'github-preview';
        existing.updatedByName = body.displayName || localStorage.getItem('mgUserName') || 'Preview User';
        existing.paidAt = paid ? stamp : '';
      } else {
        data.donutStatuses.push({
          monthKey,
          paid,
          updatedAt: stamp,
          updatedBy: sessionStorage.getItem('winUser') || 'github-preview',
          updatedByName: body.displayName || localStorage.getItem('mgUserName') || 'Preview User',
          paidAt: paid ? stamp : ''
        });
      }
      saveData(data);
      return jsonResponse({ ok: true, statuses: clone(data.donutStatuses) });
    }
    if (method === 'POST' && path === '/api/crash-consent/accept') {
      const body = parseBody(init);
      const state = previewSaveCrashConsent(data, true, body.displayName || '');
      return jsonResponse(state, state.error ? 409 : 200);
    }
    if (method === 'POST' && path === '/api/crash-consent/decline') {
      const body = parseBody(init);
      return jsonResponse(previewDeclineCrashConsent(data, body.displayName || ''));
    }
    if (method === 'GET' && path === '/api/crash-tie-poll') return jsonResponse(tiePollState(data));
    if (method === 'POST' && path === '/api/crash-tie-poll/vote') {
      const body = parseBody(init);
      const optionId = String(body.optionId || '').trim().toLowerCase();
      if (!TIE_POLL_OPTIONS.some(o => o.id === optionId)) return jsonResponse({ error: 'Choose one of the poll options.' }, 400);
      if (Date.now() >= Date.parse(TIE_POLL_CLOSES_AT)) {
        return jsonResponse({ ...tiePollState(data), error: 'The tie-break poll closed at 4:30 PM.' }, 409);
      }
      const windowsUser = (sessionStorage.getItem('winUser') || 'github-preview').trim().toLowerCase();
      data.tiePollVotes = (data.tiePollVotes || []).filter(v => !(v.pollId === TIE_POLL_ID && String(v.windowsUser || '').trim().toLowerCase() === windowsUser));
      data.tiePollVotes.push({
        pollId: TIE_POLL_ID,
        windowsUser,
        displayName: body.displayName || localStorage.getItem('mgUserName') || 'Preview User',
        optionId,
        votedAt: new Date().toISOString()
      });
      saveData(data);
      return jsonResponse(tiePollState(data));
    }
    if (method === 'POST' && path === '/api/presence/ping') return jsonResponse({
      ok: true,
      preview: true,
      lastSeen: new Date().toISOString(),
      source: 'GitHub preview local browser storage'
    });
    if (method === 'POST' && path === '/api/presence/leave') return jsonResponse({ ok: true, preview: true });
    if (method === 'GET' && path === '/api/admin/presence') return jsonResponse({
      active: [{
        windowsUser: sessionStorage.getItem('winUser') || 'github-preview',
        displayName: localStorage.getItem('mgUserName') || 'Preview User',
        machine: 'GitHub Pages',
        lastSeen: new Date().toISOString()
      }],
      queue: clone(data.captchaQueue || []),
      targets: previewKnownCaptchaTargets(data)
    });
    if (method === 'GET' && path === '/api/captcha-check') return jsonResponse({ id: null });
    if (method === 'POST' && path === '/api/captcha-ack') return jsonResponse({ ok: true, preview: true });
    if (method === 'GET' && path === '/api/captcha-lite/state') {
      const caller = (sessionStorage.getItem('winUser') || 'github-preview').trim().toLowerCase();
      if (!['dshank', 'dlebel', 'github-preview'].includes(caller)) return jsonResponse({ error: 'Not authorized.' }, 403);
      const protectedTargets = new Set(['dlebel']);
      const allowedTarget = value => value && !protectedTargets.has(String(value).trim().toLowerCase());
      const filteredTargets = previewKnownCaptchaTargets(data).map(t => ({
        ...t,
        windowsUsers: (Array.isArray(t.windowsUsers) ? t.windowsUsers : []).filter(allowedTarget)
      })).filter(t => t.windowsUsers.length > 0);
      const queue = clone(data.captchaQueue || []).filter(q => allowedTarget(q.target));
      return jsonResponse({
        ok: true,
        preview: true,
        caller,
        active: [{
          windowsUser: caller,
          displayName: localStorage.getItem('mgUserName') || 'Preview User',
          machine: 'GitHub Pages',
          lastSeen: new Date().toISOString()
        }].filter(u => allowedTarget(u.windowsUser)),
        queue,
        targets: filteredTargets,
        protected: Array.from(protectedTargets)
      });
    }
    if (method === 'POST' && path === '/api/admin/captcha-send') {
      const body = parseBody(init);
      data.captchaQueue = Array.isArray(data.captchaQueue) ? data.captchaQueue : [];
      const queuedAt = new Date().toISOString();
      data.captchaQueue.push({
        id: 'preview-captcha-' + Date.now(),
        target: String(body.target || 'preview').trim().toLowerCase(),
        count: Math.max(1, Math.min(7, parseInt(body.count || '3', 10) || 3)),
        challenge: String(body.challenge || '').trim(),
        queuedBy: sessionStorage.getItem('winUser') || 'github-preview',
        queuedAt,
        status: 'pending'
      });
      saveData(data);
      return jsonResponse({ ok: true, preview: true, id: 'preview-captcha', queuedAt });
    }
    if (method === 'POST' && path === '/api/captcha-lite/send') {
      const caller = (sessionStorage.getItem('winUser') || 'github-preview').trim().toLowerCase();
      if (!['dshank', 'dlebel', 'github-preview'].includes(caller)) return jsonResponse({ error: 'Not authorized.' }, 403);
      const body = parseBody(init);
      const target = String(body.target || '').trim().toLowerCase();
      if (!target) return jsonResponse({ error: 'target is required.' }, 400);
      if (target === 'dlebel') return jsonResponse({ error: 'That target is protected.' }, 403);
      data.captchaQueue = Array.isArray(data.captchaQueue) ? data.captchaQueue : [];
      const queuedAt = new Date().toISOString();
      const id = 'preview-lite-captcha-' + Date.now();
      data.captchaQueue.push({
        id,
        target,
        count: Math.max(1, Math.min(7, parseInt(body.count || '3', 10) || 3)),
        challenge: String(body.challenge || '').trim(),
        queuedBy: caller,
        queuedAt,
        status: 'pending'
      });
      saveData(data);
      return jsonResponse({ ok: true, preview: true, id, queuedAt });
    }
    if (method === 'POST' && path === '/api/admin/log-attempt') return jsonResponse({ ok: true, preview: true });
    if (method === 'GET' && (path === '/api/poke-targets' || path === '/api/admin/poke-targets')) {
      data.pokeTargets = Array.isArray(data.pokeTargets) ? data.pokeTargets : defaultPokeTargets();
      saveData(data);
      return jsonResponse({
        targets: clone(data.pokeTargets),
        count: data.pokeTargets.length,
        source: 'GitHub preview local browser storage'
      });
    }

    if (method === 'PATCH' && path === '/api/admin/poke-targets') {
      const body = parseBody(init);
      const targets = Array.isArray(body.targets) ? body.targets : [];
      if (targets.length === 0) return jsonResponse({ error: 'At least one target is required.' }, 400);
      data.pokeTargets = clone(targets);
      saveData(data);
      return jsonResponse({
        ok: true,
        preview: true,
        targets: clone(data.pokeTargets),
        count: data.pokeTargets.length,
        source: 'GitHub preview local browser storage'
      });
    }

    if (method === 'POST' && path === '/api/admin/poke-reset') {
      const body = parseBody(init);
      const target = String(body.target || '').trim().toLowerCase();
      if (!target) return jsonResponse({ error: 'target is required.' }, 400);
      data.pokeResets = Array.isArray(data.pokeResets) ? data.pokeResets : [];
      const resetAt = new Date().toISOString();
      const existing = data.pokeResets.find(e => String(e.user || '').trim().toLowerCase() === target);
      if (existing) existing.resetAt = resetAt;
      else data.pokeResets.push({ user: target, resetAt });
      saveData(data);
      return jsonResponse({ ok: true, preview: true, target, resetAt });
    }

    if (method === 'GET' && path === '/api/poke-reset-check') {
      const user = String(url.searchParams.get('user') || '').trim().toLowerCase();
      const entry = (data.pokeResets || []).find(e => String(e.user || '').trim().toLowerCase() === user);
      return jsonResponse({ resetAt: entry ? entry.resetAt : null });
    }

    if (method === 'POST' && path === '/api/usage/open') {
      const body = parseBody(init);
      data.usage = Array.isArray(data.usage) ? data.usage : [];
      data.usage.push({
        timestamp: new Date().toLocaleString(),
        windowsUser: body.windowsUser || sessionStorage.getItem('winUser') || 'github-preview',
        machine: 'GitHub Pages',
        page: body.page || '',
        action: 'OPEN'
      });
      saveData(data);
      return jsonResponse({ ok: true });
    }

    if (method === 'POST' && path === '/api/changelog/add') {
      const entry = parseBody(init);
      entry.id = maxId(data.changelog) + 1;
      data.changelog.push(entry);
      saveData(data);
      return jsonResponse({ ok: true, count: data.changelog.length, entries: clone(data.changelog) });
    }

    if (method === 'POST' && path === '/api/tickets/add') {
      const ticket = parseBody(init);
      ticket.id = ticket.id || Date.now();
      ticket.votes = ticket.votes || 0;
      ticket.voters = ticket.voters || [];
      ticket.pokes = ticket.pokes || 0;
      ticket.pokers = ticket.pokers || [];
      ticket.comments = ticket.comments || [];
      ticket.attachments = ticket.attachments || [];
      data.tickets.push(ticket);
      saveData(data);
      return jsonResponse({ ok: true, count: data.tickets.length, entries: clone(data.tickets) });
    }

    if (method === 'POST' && path === '/api/crashes/add') {
      const crash = parseBody(init);
      const consent = previewCrashConsentState(data, crash.createdBy || crash.user || '');
      if (!consent.canLog) {
        return jsonResponse({ error: 'You need to accept the crash donut duty rules before logging crashes.', consent }, 403);
      }
      data.crashes = Array.isArray(data.crashes) ? data.crashes : [];
      crash.id = maxId(data.crashes) + 1;
      crash.user = crash.user || localStorage.getItem('mgUserName') || 'Preview User';
      crash.severity = crash.severity || 'minor';
      crash.timestamp = crash.timestamp || new Date().toISOString();
      crash.createdBy = crash.createdBy || crash.user;
      crash.windowsUser = sessionStorage.getItem('winUser') || 'github-preview';
      crash.participantWindowsUser = crash.participantWindowsUser || crash.windowsUser;
      data.crashes.push(crash);
      saveData(data);
      return jsonResponse({ ok: true, count: data.crashes.length, entries: clone(data.crashes) });
    }

    if (path.match(/^\/api\/crashes\/\d+$/)) {
      const id = path.split('/').pop();
      data.crashes = Array.isArray(data.crashes) ? data.crashes : [];
      const crash = byId(data.crashes, id);
      if (method === 'DELETE') {
        data.crashes = data.crashes.filter(c => String(c.id) !== String(id));
        saveData(data);
        return jsonResponse({ ok: true, count: data.crashes.length, entries: clone(data.crashes) });
      }
      if (method === 'PATCH') {
        if (!crash) return jsonResponse({ error: 'Crash not found.' }, 404);
        const patch = parseBody(init);
        if (patch.user) crash.user = patch.user;
        if (patch.severity) crash.severity = patch.severity;
        if (patch.timestamp) crash.timestamp = patch.timestamp;
        if (patch.windowsUser != null) crash.windowsUser = patch.windowsUser;
        if (patch.createdBy) crash.createdBy = patch.createdBy;
        saveData(data);
        return jsonResponse({ ok: true, count: data.crashes.length, entries: clone(data.crashes) });
      }
    }

    if (path.startsWith('/api/tickets/')) {
      const result = updateTicket(path, method, init, data, url.searchParams);
      if (result) {
        if (result.error) return jsonResponse({ error: result.error }, 404);
        return jsonResponse(result);
      }
    }

    if (path.startsWith('/api/attachments/')) {
      return jsonResponse({ error: 'Attachments are not available in the GitHub preview.' }, 404);
    }

    return jsonResponse({ error: 'Preview API route not implemented: ' + method + ' ' + path }, 404);
  };
})();
</script>
'@

if ($html -notmatch '</head>') { throw 'Could not find </head> in source HTML.' }
if ($html -notmatch '<body>') { throw 'Could not find <body> in source HTML.' }
if ($html -notmatch '<script>') { throw 'Could not find application <script> in source HTML.' }

$html = $html.Replace('</head>', "$previewStyle`r`n</head>")
$html = $html.Replace('<body>', "<body>`r`n$previewBanner")

$scriptIndex = $html.IndexOf('<script>')
if ($scriptIndex -lt 0) { throw 'Could not locate first script tag.' }
$html = $html.Insert($scriptIndex, "$previewShim`r`n")

$html = "<!-- Generated by deploy/Build-GitHubPreview.ps1 from MacroGuide.html. Do not edit directly. -->`r`n$html"

[System.IO.File]::WriteAllText($outFile, $html, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $outPath '.nojekyll'), '', $utf8NoBom)

Write-Host "GitHub preview written to $outFile"
