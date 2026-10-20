/**
 * Inventory Adjust: Slack /adjust form -> Google Sheet.
 * See README.md for setup, limitations, and maintenance notes.
 *
 * Bound to the Sheet (Extensions -> Apps Script), deployed as a Web app.
 *
 * Tabs:
 *   Products  Paste your inventory system's product export. Overwrite freely on re-export.
 *             Headers matched by name, any order:
 *             required: Name, Internal Reference
 *             optional: Material Type, Unit of Measure
 *   Aliases   Hand-kept nicknames, keyed by SKU so re-exports never shift them.
 *             Headers matched by name, any order: SKU, Aliases (comma-separated);
 *             other columns (e.g. Name) are ignored. Created by setup().
 *   Learned   Created by setup(). "search -> part" pairs picked from closest-match
 *             suggestions. Tick Approved and the search works as a nickname.
 *   Log       Created by setup(). "Exported" column is stamped by /adjust export.
 *
 * Script Properties (Project Settings):
 *   SLACK_BOT_TOKEN, SLACK_VERIFICATION_TOKEN, ADJUST_CHANNEL_ID
 *   EXPORT_USERS  comma-separated Slack user IDs allowed to run /adjust export
 *
 * Slack scopes: commands, chat:write, files:write (export upload)
 */

const PRODUCTS_TAB = 'Products';
const ALIASES_TAB = 'Aliases';
const LEARNED_TAB = 'Learned';
const LEARNED_HEADERS = ['Search', 'SKU', 'Product', 'Times', 'Last used', 'Approved'];
const LOG_TAB = 'Log';
const START_SLOTS = 3;      // part/qty pairs when the form opens
const ADD_SLOTS = 3;        // pairs added per "+ Add more parts" click
const MAX_SLOTS = 45;       // Slack modals hold 100 blocks: 2 per pair + reason, notes, button
const MAX_RESULTS = 50;     // options per search (Slack max: 100)
const CLOSEST = 3;          // suggestions shown when nothing matches exactly
const FUZZY_MIN = 0.5;      // 0-1, how close a suggestion must be
const TYPEAHEAD = true;     // false = search from command text instead: /adjust m6 bolt
const REASONS = ['Scrap / damaged', 'Maintenance / repair', 'R&D / testing',
                 'Sample / customer', 'Rework', 'Count correction', 'Other'];
const LOG_HEADERS = ['Timestamp', 'Submitted by', 'Batch', 'SKU', 'Product', 'Qty', 'UoM',
                     'Reason', 'Notes', 'Status', 'Adjusted by', 'Adjusted on', 'Exported'];
const STATUSES = ['Pending', 'Done', 'Rejected'];

const PROPS = PropertiesService.getScriptProperties();


// ---------- entry point: every Slack request lands here ----------
function doPost(e) {
  const token = PROPS.getProperty('SLACK_VERIFICATION_TOKEN');

  // Interactivity: typeahead search + form submission
  if (e.parameter.payload) {
    const payload = JSON.parse(e.parameter.payload);
    if (payload.token !== token) return empty_();
    if (payload.type === 'block_actions' && payload.actions[0].action_id === 'add_rows') {
      return addRows_(payload);
    }
    if (payload.type === 'block_suggestion') {
      return json_(suggestions_(payload));
    }
    if (payload.type === 'view_submission' && payload.view.callback_id === 'adjust_submit') {
      return handleSubmit_(payload);
    }
    return empty_();
  }

  // Slash command
  if (e.parameter.token !== token) return empty_();
  const command = e.parameter.command;
  const text = (e.parameter.text || '').trim();
  if (command === '/adjustexport' || (command === '/adjust' && text.toLowerCase() === 'export')) {
    return startExport_(e.parameter.user_id, e.parameter.channel_id);
  }
  if (command === '/adjust') return openForm_(e.parameter.trigger_id, text);
  return empty_();
}


// ---------- products ----------
// Cached as arrays [sku, name, type, uom, aliases] to stay under CacheService's 100 KB limit.
// Long-lived cache: a cold sheet read is the slowest part of a search, and Slack allows 3 s.
function products_() {
  const cache = CacheService.getScriptCache();
  let rows = JSON.parse(cache.get('products') || 'null');
  if (!rows) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(PRODUCTS_TAB);
    if (!sheet) throw new Error(`No "${PRODUCTS_TAB}" tab`);
    const values = sheet.getDataRange().getDisplayValues();
    const head = values[0].map(h => String(h).trim().toLowerCase());
    const col = name => head.indexOf(name);
    const c = [col('internal reference'), col('name'), col('material type'), col('unit of measure')];
    if (c[0] < 0 || c[1] < 0) {
      throw new Error(`"${PRODUCTS_TAB}" needs "Name" and "Internal Reference" headers`);
    }
    const aliases = aliasMap_();
    rows = values.slice(1)
      .map(r => c.map(i => (i < 0 ? '' : String(r[i]).trim())))
      .filter(r => r[0] && r[1])
      .map(r => r.concat(aliases[r[0]] || ''));
    const s = JSON.stringify(rows);
    if (s.length < 100000) cache.put('products', s, 21600);  // 6 h max; cleared on edit/refresh
  }
  return rows.map(([sku, name, type, uom, aliases]) => ({
    sku, name, type, uom, aliases,
    search: [name, sku, type, aliases].join(' ').toLowerCase(),
    aliasList: aliases.toLowerCase().split(',').map(a => a.trim()).filter(Boolean),
  }));
}

// SKU -> "alias, alias". Rows with the same SKU are merged.
// 1-based SKU and Aliases columns of the Aliases tab, found by header.
// Accepts "SKU" or "Internal Reference", and "Aliases"/"Alias"/"Nicknames".
function aliasCols_(sheet) {
  const head = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0]
    .map(h => h.trim().toLowerCase());
  const find = names => head.findIndex(h => names.includes(h)) + 1;
  return { sku: find(['sku', 'internal reference']), aliases: find(['aliases', 'alias', 'nicknames']) };
}

function aliasMap_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ALIASES_TAB);
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  const c = aliasCols_(sheet);
  if (!c.sku || !c.aliases) {
    throw new Error(`"${ALIASES_TAB}" needs "SKU" and "Aliases" headers in row 1`);
  }
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  values.forEach(r => {
    const sku = String(r[c.sku - 1]).trim();
    const names = String(r[c.aliases - 1]).trim();
    if (sku && names) map[sku] = map[sku] ? `${map[sku]}, ${names}` : names;
  });
  learnedRows_().filter(r => r.approved).forEach(({ search, sku }) => {
    map[sku] = map[sku] ? `${map[sku]}, ${search}` : search;
  });
  return map;
}

function productBySku_(sku) {
  return products_().find(p => p.sku === sku);
}

// Every word must appear in name, SKU, type, or aliases, in any order.
// Ranking: exact SKU > exact nickname > name starts with query > everything else.
function score_(p, query, words) {
  if (!words.every(w => p.search.includes(w))) return 0;
  if (p.sku.toLowerCase() === query) return 4;
  if (p.aliasList.includes(query)) return 3;
  if (p.name.toLowerCase().startsWith(query)) return 2;
  return 1;
}

// exact: every word found (ranked by score_). close: only when exact is empty,
// the CLOSEST most similar parts by typo-tolerant word matching.
function search_(rawQuery, limit) {
  const query = rawQuery.toLowerCase().trim().replace(/\s+/g, ' ');
  const words = query.split(' ').filter(Boolean);
  if (!words.length) return { exact: [], close: [] };
  const products = products_();
  const exact = products
    .map(p => ({ p, score: score_(p, query, words) }))
    .filter(x => x.score > 0)
    .sort((a, b) => b.score - a.score || a.p.name.localeCompare(b.p.name))
    .slice(0, limit)
    .map(x => x.p);
  if (exact.length) return { exact, close: [] };

  // Typed words as-is; only part names get the extra split/joined forms. Expanding the
  // typed words too let "m7" partly match "m6" through its "m" piece.
  const qWords = query.replace(/[^a-z0-9]+/g, ' ').split(' ').filter(Boolean);
  const close = products
    .map(p => ({ p, score: fuzzy_(qWords, tokens_(p.search)) }))
    .filter(x => x.score >= FUZZY_MIN)
    .sort((a, b) => b.score - a.score || a.p.name.localeCompare(b.p.name))
    .slice(0, CLOSEST)
    .map(x => x.p);
  return { exact, close };
}

function option_(p) {
  return { text: plain_(label_(p)), value: p.sku };
}

// Flat list: used by the static-dropdown fallback and testSearch().
function partOptions_(rawQuery, limit) {
  const r = search_(rawQuery, limit);
  return (r.exact.length ? r.exact : r.close).map(option_);
}

// Typeahead reply. Remembers searches that found nothing exact, so a pick from the
// closest matches can be recorded in Learned when the form is submitted.
function suggestions_(payload) {
  const query = (payload.value || '').trim();
  const r = search_(query, MAX_RESULTS);
  const viewId = payload.view && payload.view.id;
  if (viewId) {
    const key = `q:${viewId}:${payload.block_id}`;
    const cache = CacheService.getScriptCache();
    if (r.exact.length) cache.remove(key);
    else cache.put(key, query.toLowerCase().replace(/\s+/g, ' '), 3600);
  }
  if (r.exact.length) return { options: r.exact.map(option_) };
  if (r.close.length) {
    return { option_groups: [{ label: plain_('No exact match. Closest parts:'), options: r.close.map(option_) }] };
  }
  return { options: [] };
}

// Words, plus letter/number splits ("15mm" also as "15", "mm") and joined neighbours
// ("heat shrink" also as "heatshrink").
function tokens_(text) {
  const words = text.toLowerCase().replace(/[^a-z0-9]+/g, ' ').split(' ').filter(Boolean);
  const out = [];
  words.forEach((t, i) => {
    out.push(t);
    const parts = t.match(/[a-z]+|[0-9]+/g);
    if (parts && parts.length > 1) out.push(...parts);
    if (i > 0) out.push(words[i - 1] + t);
  });
  return out;
}

// Average over query words of each word's best match among the part's words (0-1).
function fuzzy_(qWords, pWords) {
  if (!qWords.length) return 0;
  let total = 0;
  for (const w of qWords) {
    let best = 0;
    for (const t of pWords) {
      best = Math.max(best, wordSim_(w, t));
      if (best === 1) break;
    }
    total += best;
  }
  return total / qWords.length;
}

function wordSim_(w, t) {
  if (w === t) return 1;
  if (w.length >= 2 && t.startsWith(w)) return 0.9;            // "spac" -> "spacer"
  if (t.length >= 3 && w.startsWith(t)) return 0.8;            // "spacers" -> "spacer"
  if (/\d/.test(w) || /\d/.test(t)) return 0;                  // never fuzz numbers: M6 != M8
  if (w.length >= 3 && t.includes(w)) return 0.7;
  if (Math.abs(w.length - t.length) > 2 || w.length < 3) return 0;
  const d = editDistance_(w, t);
  const allowed = w.length <= 4 ? 1 : 2;
  return d <= allowed ? 0.85 - 0.15 * d : 0;                   // 1 typo: 0.7, 2 typos: 0.55
}

// Edit distance counting a swap of neighbours as one typo ("gladaitor").
function editDistance_(a, b) {
  const d = [];
  for (let i = 0; i <= a.length; i++) d[i] = [i];
  for (let j = 0; j <= b.length; j++) d[0][j] = j;
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + cost);
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) {
        d[i][j] = Math.min(d[i][j], d[i - 2][j - 2] + 1);
      }
    }
  }
  return d[a.length][b.length];
}


// ---------- learned nicknames ----------
function learnedRows_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(LEARNED_TAB);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, LEARNED_HEADERS.length).getValues()
    .map((r, i) => ({ row: i + 2, search: String(r[0]).trim().toLowerCase(), sku: String(r[1]).trim(),
                      times: Number(r[3]) || 0, approved: r[5] === true }))
    .filter(r => r.search && r.sku);
}

// pairs: [{ search, p }] from this submission. Same search+SKU again bumps Times.
function learn_(pairs) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(LEARNED_TAB);
  if (!sheet || !pairs.length) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;  // learning is a bonus; never hold up a submission
  try {
    const existing = learnedRows_();
    const now = new Date();
    pairs.forEach(({ search, p }) => {
      const hit = existing.find(r => r.search === search && r.sku === p.sku);
      if (hit) {
        hit.times += 1;
        sheet.getRange(hit.row, 4, 1, 2).setValues([[hit.times, now]]);
      } else {
        sheet.appendRow([search, p.sku, p.name, 1, now, false]);
        sheet.getRange(sheet.getLastRow(), 6).insertCheckboxes();
      }
    });
  } finally {
    lock.releaseLock();
  }
}

function label_(p) {
  const uom = p.uom ? ` (${p.uom})` : '';
  const full = `[${p.sku}] ${p.name}`;
  return full.length + uom.length <= 75 ? full + uom : full.slice(0, 74 - uom.length) + '…' + uom;
}


// ---------- form ----------
function openForm_(triggerId, query) {
  if (!TYPEAHEAD) {
    if (!query) return text_('Add what you are looking for, e.g. `/adjust m6 bolt`');
    if (!partOptions_(query, 1).length) {
      return text_(`No parts match "${query}". Try other words, or tag the adjuster.`);
    }
  }
  const res = slack_('views.open', { trigger_id: triggerId, view: modal_(START_SLOTS, query) });
  if (!res.ok) return text_(`Could not open the form (${res.error}). Try again.`);
  return empty_();
}

// "+ Add more parts": rebuild the form with more pairs. Slack keeps what was already
// entered because every existing block keeps its block_id and action_id.
function addRows_(payload) {
  const meta = JSON.parse(payload.view.private_metadata || '{}');
  const slots = Math.min((meta.slots || START_SLOTS) + ADD_SLOTS, MAX_SLOTS);
  slack_('views.update', {
    view_id: payload.view.id,
    hash: payload.view.hash,  // Slack rejects the update if the form changed meanwhile
    view: modal_(slots, meta.query || ''),
  });
  return empty_();
}

function partElement_(query) {
  if (TYPEAHEAD) {
    return { type: 'external_select', action_id: 'part_pick', min_query_length: 2,
             placeholder: plain_('Type name, SKU, or nickname...') };
  }
  return { type: 'static_select', action_id: 'part_pick', options: partOptions_(query, 100),
           placeholder: plain_('Pick a part...') };
}

function modal_(slots, query) {
  const partElement = partElement_(query);
  const blocks = [];
  for (let i = 1; i <= slots; i++) {
    const optional = i > 1;
    blocks.push(
      { type: 'input', block_id: `part${i}`, optional,
        label: plain_(`Part ${i}`), element: partElement },
      { type: 'input', block_id: `qty${i}`, optional,
        label: plain_(`Qty ${i}`),
        element: { type: 'number_input', action_id: 'v', is_decimal_allowed: true, min_value: '0' } });
  }
  blocks.push(
    { type: 'input', block_id: 'reason', label: plain_('Reason'),
      element: { type: 'static_select', action_id: 'v',
                 options: REASONS.map(r => ({ text: plain_(r), value: r })) } },
    { type: 'input', block_id: 'notes', optional: true,
      label: plain_('Notes (what was it used for?)'),
      element: { type: 'plain_text_input', action_id: 'v', multiline: true } });
  if (slots < MAX_SLOTS) {
    // Placed right after the last pair so it's where people look when they run out.
    blocks.splice(slots * 2, 0, {
      type: 'actions', block_id: 'more',
      elements: [{ type: 'button', action_id: 'add_rows',
                   text: plain_(`+ Add ${Math.min(ADD_SLOTS, MAX_SLOTS - slots)} more parts`) }],
    });
  }

  return {
    type: 'modal',
    callback_id: 'adjust_submit',
    private_metadata: JSON.stringify({ slots, query }),
    title: plain_('Inventory Adjustment'),
    submit: plain_('Log it'),
    close: plain_('Cancel'),
    blocks,
  };
}

function handleSubmit_(payload) {
  // A slow reply makes Slack show an error and keep the form open, though the first
  // submit still finishes. Clicking "Log it" again resends the same view.id: drop it.
  const cache = CacheService.getScriptCache();
  const seenKey = 'view:' + payload.view.id;
  if (cache.get(seenKey)) return empty_();

  const v = payload.view.state.values;
  const items = [];
  const errors = {};
  const slots = Object.keys(v).filter(k => /^part\d+$/.test(k)).length;
  for (let i = 1; i <= slots; i++) {
    const opt = v[`part${i}`].part_pick.selected_option;
    const qtyRaw = v[`qty${i}`].v.value;
    if (!opt && !qtyRaw) continue;                                   // empty slot
    if (!opt) { errors[`part${i}`] = 'Pick a part for this quantity'; continue; }
    const qty = Number(qtyRaw);
    if (!(qty > 0)) { errors[`qty${i}`] = 'Quantity must be more than 0'; continue; }
    const p = productBySku_(opt.value);
    if (!p) { errors[`part${i}`] = 'Part no longer in the list, pick again'; continue; }
    items.push({ p, qty, slot: i });
  }
  if (Object.keys(errors).length) return json_({ response_action: 'errors', errors });

  cache.put(seenKey, '1', 3600);
  const batch = Utilities.getUuid().slice(0, 8);
  const reason = v.reason.v.selected_option.value;
  const notes = v.notes.v.value || '';
  const now = new Date();
  const rows = items.map(({ p, qty }) => [
    now, payload.user.name, batch, p.sku, p.name, qty, p.uom, reason, notes, 'Pending', '', '', '',
  ]);
  writeRows_(rows);

  // Searches that found nothing exact, followed by a pick: candidates for Learned.
  const keys = items.map(it => `q:${payload.view.id}:part${it.slot}`);
  const searches = cache.getAll(keys);
  learn_(items
    .map((it, n) => ({ search: searches[keys[n]], p: it.p }))
    .filter(x => x.search && !x.p.aliasList.includes(x.search)));

  const lines = items.map(({ p, qty }) => `• *${qty}${p.uom ? ' ' + p.uom : ''}* × [${p.sku}] ${p.name}`);
  slack_('chat.postMessage', {
    channel: PROPS.getProperty('ADJUST_CHANNEL_ID'),
    text: `<@${payload.user.id}> logged ${items.length} item(s) — ${reason}` +
          `${notes ? ': ' + notes : ''}\n${lines.join('\n')}`,
  });
  return empty_();
}

// One write for the whole batch; lock keeps simultaneous submissions from overlapping.
function writeRows_(rows) {
  const log = SpreadsheetApp.getActive().getSheetByName(LOG_TAB);
  const lock = LockService.getScriptLock();
  if (lock.tryLock(1500)) {
    try {
      log.getRange(log.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    } finally {
      lock.releaseLock();
    }
  } else {
    rows.forEach(r => log.appendRow(r));
  }
}


// ---------- export: /adjust export -> .xlsx of every Log row not yet exported ----------
// Slack wants a reply within 3 seconds, so the command only queues the job and a
// one-off trigger does the work. Rows are stamped "Exported" only after the upload succeeds.
function startExport_(userId, channelId) {
  const allowed = (PROPS.getProperty('EXPORT_USERS') || '').split(',').map(s => s.trim());
  if (!allowed.includes(userId)) return text_('Only the adjuster can run exports.');
  const pending = JSON.parse(PROPS.getProperty('EXPORT_REQUEST') || 'null');
  if (pending && Date.now() - pending.at < 10 * 60 * 1000) {
    return text_('An export is already running, it will show up here shortly.');
  }
  PROPS.setProperty('EXPORT_REQUEST', JSON.stringify({ userId, channelId, at: Date.now() }));
  ScriptApp.newTrigger('runExport').timeBased().after(1000).create();
  return text_('Building the export, it will show up in this channel in about a minute.');
}

// Trigger handler (public name so the trigger can call it).
function runExport() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runExport')
    .forEach(t => ScriptApp.deleteTrigger(t));
  const req = JSON.parse(PROPS.getProperty('EXPORT_REQUEST') || 'null');
  PROPS.deleteProperty('EXPORT_REQUEST');
  if (!req) return;

  const log = SpreadsheetApp.getActive().getSheetByName(LOG_TAB);
  const exportedCol = LOG_HEADERS.indexOf('Exported');
  const last = log.getLastRow();
  const values = last < 2 ? [] : log.getRange(2, 1, last - 1, LOG_HEADERS.length).getValues();
  const picked = values
    .map((r, i) => ({ r, rowNumber: i + 2 }))
    .filter(({ r }) => r[0] !== '' && !r[exportedCol]);
  if (!picked.length) {
    slack_('chat.postMessage', { channel: req.channelId, text: 'Nothing new since the last export.' });
    return;
  }

  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const cells = picked.map(({ r }) => r.slice(0, exportedCol)
    .map(v => (Object.prototype.toString.call(v) === '[object Date]' ? Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm') : v)));
  const filename = `inventory-adjustments_${Utilities.formatDate(now, tz, 'yyyy-MM-dd_HHmm')}.xlsx`;
  const blob = xlsxBlob_(LOG_HEADERS.slice(0, exportedCol), cells, filename);

  const err = uploadToSlack_(blob, req.channelId,
    `<@${req.userId}> ${picked.length} row(s) since the last export.`);
  if (err) {
    slack_('chat.postMessage', { channel: req.channelId, text: `Export failed (${err}). Nothing was marked as exported.` });
    return;
  }

  // Stamp only the exported rows; rows added meanwhile stay unexported.
  const stamp = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm');
  picked.forEach(({ rowNumber }) => log.getRange(rowNumber, exportedCol + 1).setValue(stamp));
}

// Slack's upload flow: get an upload URL, send the bytes, then attach the file to the channel.
function uploadToSlack_(blob, channelId, comment) {
  const bytes = blob.getBytes();
  const step1 = slackForm_('files.getUploadURLExternal', { filename: blob.getName(), length: String(bytes.length) });
  if (!step1.ok) return step1.error;
  const put = UrlFetchApp.fetch(step1.upload_url, {
    method: 'post', contentType: 'application/octet-stream', payload: bytes, muteHttpExceptions: true,
  });
  if (put.getResponseCode() !== 200) return `upload HTTP ${put.getResponseCode()}`;
  const step3 = slackForm_('files.completeUploadExternal', {
    files: JSON.stringify([{ id: step1.file_id, title: blob.getName() }]),
    channel_id: channelId,
    initial_comment: comment,
  });
  return step3.ok ? null : step3.error;
}

// Minimal .xlsx (one sheet, inline strings) built by hand, so no Drive access is needed.
function xlsxParts_(header, rows) {
  const esc = v => String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const cell = v => (typeof v === 'number' && isFinite(v)
    ? `<c><v>${v}</v></c>`
    : `<c t="inlineStr"><is><t xml:space="preserve">${esc(v == null ? '' : v)}</t></is></c>`);
  const row = r => `<row>${r.map(cell).join('')}</row>`;
  const ns = 'http://schemas.openxmlformats.org';
  const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  return {
    '[Content_Types].xml': `${head}<Types xmlns="${ns}/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `</Types>`,
    '_rels/.rels': `${head}<Relationships xmlns="${ns}/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${ns}/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
      `</Relationships>`,
    'xl/workbook.xml': `${head}<workbook xmlns="${ns}/spreadsheetml/2006/main" xmlns:r="${ns}/officeDocument/2006/relationships">` +
      `<sheets><sheet name="Adjustments" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `${head}<Relationships xmlns="${ns}/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${ns}/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
      `</Relationships>`,
    'xl/worksheets/sheet1.xml': `${head}<worksheet xmlns="${ns}/spreadsheetml/2006/main">` +
      `<sheetData>${[header, ...rows].map(row).join('')}</sheetData></worksheet>`,
  };
}

function xlsxBlob_(header, rows, filename) {
  const parts = xlsxParts_(header, rows);
  const blobs = Object.keys(parts).map(name => Utilities.newBlob(parts[name], 'application/xml', name));
  return Utilities.zip(blobs, filename)
    .setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}


// ---------- run these from the editor ----------
// One-time: creates the Log tab, Status dropdown, and checks config. Read the output in the execution log.
function setup() {
  const ss = SpreadsheetApp.getActive();
  const log = ss.getSheetByName(LOG_TAB) || ss.insertSheet(LOG_TAB);
  if (log.getLastRow() === 0) {
    log.appendRow(LOG_HEADERS);
    log.setFrozenRows(1);
    log.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold');
  }
  const aliases = ss.getSheetByName(ALIASES_TAB) || ss.insertSheet(ALIASES_TAB);
  if (aliases.getLastRow() === 0) {
    aliases.appendRow(['Name', 'SKU', 'Aliases']);
    aliases.setFrozenRows(1);
    aliases.getRange(1, 1, 1, 3).setFontWeight('bold');
  }

  const learned = ss.getSheetByName(LEARNED_TAB) || ss.insertSheet(LEARNED_TAB);
  if (learned.getLastRow() === 0) {
    learned.appendRow(LEARNED_HEADERS);
    learned.setFrozenRows(1);
    learned.getRange(1, 1, 1, LEARNED_HEADERS.length).setFontWeight('bold');
  }

  // Older Log tabs: add any header columns added since (e.g. Exported).
  const head = log.getRange(1, 1, 1, LOG_HEADERS.length).getValues()[0];
  LOG_HEADERS.forEach((h, i) => { if (!head[i]) log.getRange(1, i + 1).setValue(h).setFontWeight('bold'); });

  // The Aliases SKU column only accepts SKUs that exist in Products, so typos can't happen.
  // Clear the rule from every column first: older versions put it on column A.
  const products = ss.getSheetByName(PRODUCTS_TAB);
  const skuCol = products.getRange(1, 1, 1, products.getLastColumn()).getDisplayValues()[0]
    .map(h => h.trim().toLowerCase()).indexOf('internal reference') + 1;
  const aliasSkuCol = aliasCols_(aliases).sku;
  aliases.getRange(2, 1, aliases.getMaxRows() - 1, aliases.getMaxColumns()).clearDataValidations();
  if (skuCol > 0 && aliasSkuCol > 0) {
    aliases.getRange(2, aliasSkuCol, aliases.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(products.getRange(2, skuCol, products.getMaxRows() - 1, 1), false)
        .setAllowInvalid(false)
        .setHelpText('Must be an Internal Reference from the Products tab')
        .build());
  }

  const status = LOG_HEADERS.indexOf('Status') + 1;
  log.getRange(2, status, log.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());

  CacheService.getScriptCache().remove('products');
  const list = products_();
  const known = new Set(list.map(p => p.sku));
  const orphans = Object.keys(aliasMap_()).filter(sku => !known.has(sku));
  const missing = ['SLACK_BOT_TOKEN', 'SLACK_VERIFICATION_TOKEN', 'ADJUST_CHANNEL_ID', 'EXPORT_USERS']
    .filter(k => !PROPS.getProperty(k));
  Logger.log(`Products loaded: ${list.length}`);
  Logger.log(`Alias SKUs not in Products: ${orphans.join(', ') || 'none'}`);
  Logger.log(`Missing Script Properties: ${missing.join(', ') || 'none'}`);

  // Every 5 min: keep the product cache filled and the script warm, to cut slow first replies.
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'warmUp')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('warmUp').timeBased().everyMinutes(5).create();
  Logger.log('Warm-up trigger installed (every 5 min)');
}

function warmUp() {
  products_();
}

// Sheet menu: File -> Import doesn't fire onEdit, so refresh by hand after re-importing Products.
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Inventory Adjust')
    .addItem('Refresh parts list', 'refreshParts')
    .addToUi();
}

function refreshParts() {
  CacheService.getScriptCache().remove('products');
  const n = products_().length;
  SpreadsheetApp.getActive().toast(`Parts list refreshed: ${n} products`);
}

// Try searches without Slack.
function testSearch() {
  ['controller board', 'relay module', 'EL-2001', 'bolt m6', 'controler bord', 'spacers 15 mm'].forEach(q =>
    Logger.log(`${q} -> ${partOptions_(q, 5).map(o => o.text.text).join(' | ') || '(no match)'}`));
}

// Simple trigger: editing Products or Aliases drops the cache so changes show immediately.
function onEdit(e) {
  if (e && [PRODUCTS_TAB, ALIASES_TAB, LEARNED_TAB].includes(e.range.getSheet().getName())) {
    CacheService.getScriptCache().remove('products');
  }
}


// ---------- helpers ----------
function slack_(method, body) {
  const res = UrlFetchApp.fetch('https://slack.com/api/' + method, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    headers: { Authorization: 'Bearer ' + PROPS.getProperty('SLACK_BOT_TOKEN') },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  return JSON.parse(res.getContentText());
}

// Form-encoded variant: Slack's file upload methods don't take JSON bodies.
function slackForm_(method, params) {
  const res = UrlFetchApp.fetch('https://slack.com/api/' + method, {
    method: 'post',
    headers: { Authorization: 'Bearer ' + PROPS.getProperty('SLACK_BOT_TOKEN') },
    payload: params,
    muteHttpExceptions: true,
  });
  return JSON.parse(res.getContentText());
}

function plain_(text) {
  return { type: 'plain_text', text };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function text_(message) {
  return ContentService.createTextOutput(message);
}

function empty_() {
  return ContentService.createTextOutput('');
}
