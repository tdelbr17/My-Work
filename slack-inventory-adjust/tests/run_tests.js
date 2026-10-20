// Local tests for src/Code.js. Runs the Apps Script code in Node with small
// stand-ins for SpreadsheetApp, CacheService, UrlFetchApp, etc. Nothing touches
// real Slack or Google.  Usage: node tests/run_tests.js
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

// ---------- sample data ----------
function parseCsv(text) {
  const rows = [];
  let row = [], field = '', quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') { field += '"'; i++; }
      else if (ch === '"') quoted = false;
      else field += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === ',') { row.push(field); field = ''; }
    else if (ch === '\n') { row.push(field.replace(/\r$/, '')); rows.push(row); row = []; field = ''; }
    else field += ch;
  }
  if (field || row.length) { row.push(field); rows.push(row); }
  return rows;
}
const products = parseCsv(fs.readFileSync(path.join(__dirname, 'sample-products.csv'), 'utf8'));

// ---------- stand-ins for Apps Script services ----------
function makeEnv() {
  const env = {
    props: { SLACK_VERIFICATION_TOKEN: 'tok', ADJUST_CHANNEL_ID: 'C1', SLACK_BOT_TOKEN: 'x', EXPORT_USERS: 'UADJ' },
    cache: {}, log: [], learned: [], triggers: [], calls: [], zip: null,
    aliases: [['Name', 'SKU', 'Aliases'], ['15mm Block Spacer', 'HW-1015', 'spacer block'], ['', 'HW-1015', '15 spacer']],
    responses: { 'files.getUploadURLExternal': { ok: true, upload_url: 'https://files.example/up', file_id: 'F1' } },
  };
  const grid = rows => ({
    getLastRow: () => rows.length,
    getLastColumn: () => rows[0].length,
    getRange: (r, c, n, m) => ({ getDisplayValues: () => rows.slice(r - 1, r - 1 + n).map(x => x.slice(c - 1, c - 1 + m)) }),
  });
  const sheets = {
    Products: { getDataRange: () => ({ getDisplayValues: () => products }) },
    get Aliases() { return grid(env.aliases); },
    Learned: {
      getLastRow: () => env.learned.length + 1,
      getRange: (r, c, n, m) => ({
        getValues: () => env.learned.map(x => x.slice()),
        setValues: v => env.learned[r - 2].splice(c - 1, v[0].length, ...v[0]),
        insertCheckboxes() {},
      }),
      appendRow: r => env.learned.push(r),
    },
    Log: {
      getLastRow: () => env.log.length + 1,
      getRange: (r, c, n, m) => (n
        ? { setValues: v => env.log.push(...v), getValues: () => env.log.map(x => x.slice(0, m)) }
        : { setValue: v => { env.log[r - 2][c - 1] = v; } }),
      appendRow: r => env.log.push(r),
    },
  };
  const ctx = {
    PropertiesService: { getScriptProperties: () => ({
      getProperty: k => env.props[k] ?? null, setProperty: (k, v) => { env.props[k] = v; }, deleteProperty: k => delete env.props[k] }) },
    CacheService: { getScriptCache: () => ({
      get: k => env.cache[k] ?? null,
      getAll: ks => Object.fromEntries(ks.filter(k => k in env.cache).map(k => [k, env.cache[k]])),
      put: (k, v) => { env.cache[k] = v; }, remove: k => delete env.cache[k] }) },
    SpreadsheetApp: { getActive: () => ({ getSheetByName: n => sheets[n] }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    ScriptApp: {
      newTrigger: n => ({ timeBased: () => ({ after: () => ({ create: () => env.triggers.push(n) }) }) }),
      getProjectTriggers: () => env.triggers.map(n => ({ getHandlerFunction: () => n })),
      deleteTrigger: () => env.triggers.splice(0),
    },
    Session: { getScriptTimeZone: () => 'Etc/UTC' },
    Utilities: {
      getUuid: () => 'abcdef12-3456',
      formatDate: (d, tz, f) => (f.includes('_') ? '2026-01-01_0900' : '2026-01-01 09:00'),
      newBlob: (data, type, name) => ({ data, name }),
      zip: (blobs, name) => { env.zip = { blobs, name }; return { setContentType() { return this; }, getName: () => name, getBytes: () => [1, 2, 3] }; },
    },
    UrlFetchApp: { fetch: (url, o) => {
      const method = url.split('/').pop();
      env.calls.push([method, typeof o.payload === 'string' ? JSON.parse(o.payload) : o.payload]);
      return { getResponseCode: () => 200, getContentText: () => JSON.stringify(env.responses[method] || { ok: true }) };
    } },
    ContentService: { createTextOutput: s => ({ body: s, setMimeType() { return this; } }), MimeType: { JSON: 'json' } },
    Logger: { log() {} },
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'src', 'Code.js'), 'utf8'), ctx);
  env.run = code => vm.runInContext(code, ctx);
  env.post = params => env.run(`doPost(${JSON.stringify({ parameter: params })})`).body;
  env.interact = payload => env.post({ payload: JSON.stringify({ token: 'tok', ...payload }) });
  return env;
}

function form(items, { reason = 'Rework', notes = '', slots = 3 } = {}) {
  const v = {};
  for (let i = 1; i <= slots; i++) {
    v[`part${i}`] = { part_pick: { selected_option: null } };
    v[`qty${i}`] = { v: { value: null } };
  }
  items.forEach(([slot, sku, qty]) => {
    if (sku) v[`part${slot}`].part_pick.selected_option = { value: sku };
    if (qty != null) v[`qty${slot}`].v.value = String(qty);
  });
  v.reason = { v: { selected_option: { value: reason } } };
  v.notes = { v: { value: notes } };
  return v;
}
const submit = (env, viewId, values) => env.interact({
  type: 'view_submission', user: { id: 'U1', name: 'tester' },
  view: { id: viewId, callback_id: 'adjust_submit', state: { values } },
});
const names = (env, q, limit = 5) => env.run(`JSON.stringify(partOptions_(${JSON.stringify(q)}, ${limit}))`);
const firstLabel = (env, q) => JSON.parse(names(env, q))[0]?.text.text;
const search = (env, q) => JSON.parse(env.run(`JSON.stringify(search_(${JSON.stringify(q)}, 5))`));

// ---------- tests ----------
const tests = [];
const test = (name, fn) => tests.push([name, fn]);

test('loads every product and fits the cache', () => {
  const env = makeEnv();
  assert.strictEqual(env.run('products_().length'), products.length - 1);
  assert.ok(env.cache.products.length < 100000);
});

test('every word must match, in any order', () => {
  const env = makeEnv();
  assert.deepStrictEqual(search(env, 'board controller').exact.map(p => p.sku), ['EL-2001', 'EL-2002']);
});

test('exact SKU ranks first', () => {
  assert.strictEqual(firstLabel(makeEnv(), 'el-2002'), '[EL-2002] Controller Board Rev B');
});

test('nicknames from the Aliases tab (Name, SKU, Aliases in any order)', () => {
  const env = makeEnv();
  assert.strictEqual(firstLabel(env, 'spacer block'), '[HW-1015] 15mm Block Spacer');
  env.aliases = [['Aliases', 'Internal Reference'], ['blue tape', 'CS-2001']];
  delete env.cache.products;
  assert.strictEqual(firstLabel(env, 'blue tape'), '[CS-2001] Masking Tape 2in Blue');
});

test('missing Aliases headers give a clear error', () => {
  const env = makeEnv();
  env.aliases = [['Name', 'Nick'], ['x', 'y']];
  assert.throws(() => env.run('aliasMap_()'), /needs "SKU" and "Aliases" headers/);
});

test('labels fit Slack\'s 75-character limit', () => {
  const env = makeEnv();
  const long = JSON.stringify({ sku: 'X-1', name: 'N'.repeat(120), uom: 'Units' });
  assert.ok(env.run(`label_(${long})`).length <= 75);
});

test('closest matches when nothing matches exactly', () => {
  const env = makeEnv();
  const cases = {
    'controler bord': 'EL-2001',   // typos
    'smart relya': 'EL-3001',      // swapped letters
    'spacers': 'HW-1015',          // plural
    'heatshrink': 'CS-2101',       // joined words
  };
  for (const [q, sku] of Object.entries(cases)) {
    const r = search(env, q);
    assert.strictEqual(r.exact.length, 0, q);
    assert.ok(r.close.some(p => p.sku === sku), `${q} -> ${JSON.stringify(r.close.map(p => p.sku))}`);
    assert.ok(r.close.length <= 3, q);
  }
  assert.deepStrictEqual(search(env, 'xyzzy').close, []);
});

test('numbers are never fuzzed (M6 must not count as close to M8)', () => {
  const env = makeEnv();
  for (const [a, b] of [['m7', 'm6'], ['m8', 'm6'], ['16s', '18s'], ['20mm', '25mm']]) {
    assert.strictEqual(env.run(`wordSim_(${JSON.stringify(a)}, ${JSON.stringify(b)})`), 0, `${a} vs ${b}`);
  }
  // Only the "bolt" word matches, so "bolt m7" scores at most half.
  assert.ok(env.run('fuzzy_(["bolt", "m7"], tokens_("bolt m6 x 20mm"))') <= 0.5);
});

test('typeahead reply groups closest matches under a heading', () => {
  const env = makeEnv();
  const body = JSON.parse(env.interact({ type: 'block_suggestion', value: 'controler bord', block_id: 'part1', view: { id: 'V1' } }));
  assert.strictEqual(body.option_groups[0].label.text, 'No exact match. Closest parts:');
});

test('wrong token is ignored', () => {
  const env = makeEnv();
  assert.strictEqual(env.post({ token: 'nope', command: '/adjust', trigger_id: 'T' }), '');
  assert.strictEqual(env.calls.length, 0);
});

test('/adjust opens the form with 3 rows and an add button', () => {
  const env = makeEnv();
  env.post({ token: 'tok', command: '/adjust', trigger_id: 'T', text: '' });
  const [method, body] = env.calls[0];
  assert.strictEqual(method, 'views.open');
  assert.strictEqual(body.view.blocks.filter(b => /^part\d+$/.test(b.block_id)).length, 3);
  assert.ok(body.view.blocks.some(b => b.block_id === 'more'));
});

test('"+ Add more parts" grows the form and stops at 45', () => {
  const env = makeEnv();
  const click = slots => {
    env.interact({ type: 'block_actions', actions: [{ action_id: 'add_rows' }],
                   view: { id: 'V1', hash: 'h', private_metadata: JSON.stringify({ slots, query: '' }) } });
    return JSON.parse(env.calls.at(-1)[1].view.private_metadata).slots;
  };
  assert.strictEqual(click(3), 6);
  assert.strictEqual(click(44), 45);
  const full = env.run('modal_(45, "")');
  assert.ok(full.blocks.length <= 100);
  assert.ok(!full.blocks.some(b => b.block_id === 'more'));
});

test('submit writes one row per item with a shared batch and one post', () => {
  const env = makeEnv();
  submit(env, 'V1', form([[1, 'EL-2001', 2], [3, 'HW-1015', 1.5]]));
  assert.deepStrictEqual(env.log.map(r => [r[3], r[5], r[2]]), [['EL-2001', 2, 'abcdef12'], ['HW-1015', 1.5, 'abcdef12']]);
  assert.strictEqual(env.calls.filter(c => c[0] === 'chat.postMessage').length, 1);
});

test('validation errors land on the right fields', () => {
  const env = makeEnv();
  const body = JSON.parse(submit(env, 'V1', form([[1, 'EL-2001', 0], [2, null, 4], [3, 'NOPE', 1]])));
  assert.deepStrictEqual(Object.keys(body.errors).sort(), ['part2', 'part3', 'qty1']);
  assert.strictEqual(env.log.length, 0);
});

test('a resubmitted form (same view id) adds nothing', () => {
  const env = makeEnv();
  submit(env, 'V1', form([[1, 'EL-2001', 1]]));
  submit(env, 'V1', form([[1, 'EL-2001', 1]]));
  assert.strictEqual(env.log.length, 1);
});

test('learned nicknames: record, count, approve', () => {
  const env = makeEnv();
  const pick = viewId => {
    env.interact({ type: 'block_suggestion', value: 'controler bord', block_id: 'part1', view: { id: viewId } });
    submit(env, viewId, form([[1, 'EL-2001', 1]]));
  };
  pick('V1');
  pick('V2');
  assert.strictEqual(JSON.stringify(env.learned.map(r => r.slice(0, 4))),
    JSON.stringify([['controler bord', 'EL-2001', 'Controller Board Rev A', 2]]));
  env.learned[0][5] = true;
  delete env.cache.products;
  assert.strictEqual(search(env, 'controler bord').exact[0].sku, 'EL-2001');
});

test('an exact search clears the remembered no-match search', () => {
  const env = makeEnv();
  env.interact({ type: 'block_suggestion', value: 'controler bord', block_id: 'part1', view: { id: 'V1' } });
  env.interact({ type: 'block_suggestion', value: 'controller', block_id: 'part1', view: { id: 'V1' } });
  submit(env, 'V1', form([[1, 'EL-2001', 1]]));
  assert.strictEqual(env.learned.length, 0);
});

test('export: only allowed users, one at a time', () => {
  const env = makeEnv();
  assert.match(env.post({ token: 'tok', command: '/adjust', text: 'export', user_id: 'UX', channel_id: 'C9' }), /Only the adjuster/);
  assert.match(env.post({ token: 'tok', command: '/adjustexport', user_id: 'UADJ', channel_id: 'C9' }), /Building the export/);
  assert.match(env.post({ token: 'tok', command: '/adjust', text: 'export', user_id: 'UADJ', channel_id: 'C9' }), /already running/);
  assert.deepStrictEqual(env.triggers, ['runExport']);
});

test('export: uploads unexported rows, stamps them, then reports nothing new', () => {
  const env = makeEnv();
  submit(env, 'V1', form([[1, 'EL-2001', 2], [2, 'HW-1015', 1]]));
  env.post({ token: 'tok', command: '/adjust', text: 'export', user_id: 'UADJ', channel_id: 'C9' });
  env.calls.length = 0;
  env.run('runExport()');
  assert.deepStrictEqual(env.calls.map(c => c[0]), ['files.getUploadURLExternal', 'up', 'files.completeUploadExternal']);
  assert.ok(env.log.every(r => r[12] === '2026-01-01 09:00'));
  assert.deepStrictEqual(env.triggers, []);

  env.post({ token: 'tok', command: '/adjust', text: 'export', user_id: 'UADJ', channel_id: 'C9' });
  env.calls.length = 0;
  env.run('runExport()');
  assert.match(env.calls[0][1].text, /Nothing new/);
});

test('export: a failed upload stamps nothing', () => {
  const env = makeEnv();
  submit(env, 'V1', form([[1, 'EL-2001', 2]]));
  env.responses['files.completeUploadExternal'] = { ok: false, error: 'not_in_channel' };
  env.post({ token: 'tok', command: '/adjust', text: 'export', user_id: 'UADJ', channel_id: 'C9' });
  env.run('runExport()');
  assert.match(env.calls.at(-1)[1].text, /Export failed \(not_in_channel\)/);
  assert.strictEqual(env.log[0][12], '');
});

test('export: the .xlsx parts are well formed and escaped', () => {
  const env = makeEnv();
  const parts = env.run(`xlsxParts_(['A', 'B'], [['x < y & "z"', 3]])`);
  assert.deepStrictEqual(Object.keys(parts).sort(),
    ['[Content_Types].xml', '_rels/.rels', 'xl/_rels/workbook.xml.rels', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
  const sheet = parts['xl/worksheets/sheet1.xml'];
  assert.ok(sheet.includes('x &lt; y &amp; &quot;z&quot;'));
  assert.ok(sheet.includes('<c><v>3</v></c>'));
});

// ---------- run ----------
let failed = 0;
for (const [name, fn] of tests) {
  try { fn(); console.log(`ok    ${name}`); }
  catch (e) { failed++; console.log(`FAIL  ${name}\n      ${e.message}`); }
}
console.log(`\n${tests.length - failed}/${tests.length} passed`);
process.exit(failed ? 1 : 0);
