# Slack inventory adjustment form (Google Apps Script)

A Slack `/adjust` command that opens a form where people log parts they took out
of inventory for something other than normal production (scrap, repairs,
testing, samples, and so on). Every entry lands in a Google Sheet with an exact
SKU, ready for whoever adjusts the inventory system.

**The problem it solves:** people in a shared Slack channel called the same part
by different names ("blue tape", "painters tape", "2in tape"), so whoever did
the stock adjustments had to guess which product each message meant. The
requirement was **100% correct entries**, which ruled out any approach that
guesses names: fuzzy matching, embeddings, or RAG. Instead, the form only lets
people pick real products from the catalog, so there's nothing to match
afterwards.

Runs free on Google Apps Script. No server or always-on computer needed.

## Who built it

- **Code: written by Claude** (Anthropic's AI, model Claude Opus 5.5, used
  through Claude Code). That covers all of `src/`, the tests, and this README.
- **The repo owner:**
  - identified the problem and set the requirements
  - made the design decisions
  - did the Slack, Google, and inventory-system setup
  - ran the real-world testing and reported what broke

This is AI-written code, published as a working example. Please read it
before trusting it.

## Features

- **Searchable part picker:** type a name, SKU, or nickname. Every typed word
  must appear, in any order ("board controller" finds "Controller Board Rev A").
  Ranking: exact SKU, then exact nickname, then name starts with the search,
  then the rest A–Z.
- **Closest matches when nothing matches:** the 3 most similar parts, shown
  under a *"No exact match. Closest parts:"* heading. Tolerates typos and
  swapped letters, plurals, and spacing ("15 mm" / "15mm", "heatshrink" /
  "heat shrink"). **Numbers are never fuzzed**, so M6 never counts as close to M8.
- **Nicknames (Aliases tab):** `SKU | Aliases`, comma-separated, kept by hand.
  Keyed by SKU, so re-importing the product list never shifts a nickname onto
  the wrong part.
- **Self-learning nicknames (Learned tab):** when a search found nothing exact
  and the person then picked a part, the script records *search → part* with a
  count. A reviewer ticks **Approved** and the search becomes a nickname. It's
  not automatic, so one wrong pick can't teach everyone the wrong part.
- **Many items per form:** starts with 3 rows, and **"+ Add 3 more parts"**
  adds more, up to 45 (Slack allows 100 blocks per form). Shared Reason and
  Notes. Each item becomes its own Log row with a shared Batch ID.
- **Validation:** a quantity with no part, a part with no quantity, or a
  quantity of 0 or less is flagged on that field.
- **Duplicate guard:** if Slack times out on submit, the entry has usually
  saved anyway. Clicking submit again is ignored instead of creating a second
  copy.
- **`/adjust export`** (or a separate `/adjustexport` command): posts an
  `.xlsx` of every Log row not yet exported, then stamps those rows.
  - Limited to listed users.
  - A failed upload stamps nothing, so no row is ever skipped.
  - The `.xlsx` is built by hand (zipped XML), so no Google Drive access is
    needed.
- **Speed:** the product list is cached for 6 h and cleared on edit. A
  warm-up job runs every 5 min, and a sheet menu offers **Refresh parts list**.

## How it works

```
Slack ──► Apps Script web app (doPost) ──► Google Sheet
            /adjust          → opens the form (views.open)
            typing in Part   → search reply (external_select options)
            "+ Add parts"    → views.update with more rows
            submit           → validate, write Log rows, post a summary
            /adjust export   → queue a one-off trigger → build .xlsx → Slack upload
```

Sheet tabs:

| Tab | Contents |
|---|---|
| `Products` | Your inventory system's product export, pasted as is. Required headers `Name`, `Internal Reference`; optional `Material Type`, `Unit of Measure`. Any order |
| `Aliases` | `SKU` (or `Internal Reference`) and `Aliases` headers, any order. Other columns (e.g. `Name`) are ignored |
| `Learned` | Created by `setup()`: `Search, SKU, Product, Times, Last used, Approved` |
| `Log` | Created by `setup()`: `Timestamp, Submitted by, Batch, SKU, Product, Qty, UoM, Reason, Notes, Status, Adjusted by, Adjusted on, Exported` |

## Setup

1. **Sheet:** create a Google Sheet. Rename the first tab `Products` and import
   your product export into it (File → Import → Replace current sheet).
2. **Script:** in the Sheet, open Extensions → Apps Script. Paste
   `src/Code.js` into `Code.gs`. In Project Settings, tick *Show
   appsscript.json* and replace its contents with `src/appsscript.json`. Set
   your own `timeZone`.
3. In the function dropdown, pick **`setup`** and click Run. Approve the
   permissions. If you're signed in to several Google accounts and get *"An
   unknown error has occurred"*, use an incognito window with only one account.
   Then run **`testSearch`** and check the log. (You'll need to edit its
   sample searches to suit your parts list.)
4. **Deploy:** Deploy → New deployment → Web app. Execute as: **Me**. Who has
   access: **Anyone**. Copy the `/exec` URL.
5. **Slack app:** at api.slack.com/apps → Create New App → From a manifest.
   Replace `WEB_APP_URL` everywhere:
   ```yaml
   display_information:
     name: Inventory Adjust
   features:
     bot_user:
       display_name: Inventory Adjust
       always_online: true
     slash_commands:
       - command: /adjust
         url: WEB_APP_URL
         description: Log an inventory adjustment
         should_escape: false
   oauth_config:
     scopes:
       bot:
         - commands
         - chat:write
         - files:write
   settings:
     interactivity:
       is_enabled: true
       request_url: WEB_APP_URL
       message_menu_options_url: WEB_APP_URL
   ```
   Install it, then invite the bot to your channel.
6. **Script Properties** (Apps Script → Project Settings):

   | Property | Value |
   |---|---|
   | `SLACK_BOT_TOKEN` | OAuth & Permissions → Bot User OAuth Token (`xoxb-…`) |
   | `SLACK_VERIFICATION_TOKEN` | Basic Information → App Credentials → Verification Token |
   | `ADJUST_CHANNEL_ID` | The channel's ID (`C…`), from its details panel |
   | `EXPORT_USERS` | Comma-separated Slack member IDs allowed to export (`U…`) |

7. Run `setup` again. The log should say `Missing Script Properties: none`.

Tunables are at the top of `Code.js`:
- `START_SLOTS`, `ADD_SLOTS`, `MAX_SLOTS`: form rows
- `MAX_RESULTS`: search results
- `CLOSEST`, `FUZZY_MIN`: closest-match count and strictness
- `REASONS`: the Reason dropdown
- `TYPEAHEAD`: set to `false` to search from the command text
  (`/adjust m6 bolt`) instead of the live search box

## Tests

```
node tests/run_tests.js
```

- 21 tests run `src/Code.js` in Node with small stand-ins for the Apps Script
  services, against a made-up product list (`tests/sample-products.csv`).
- They cover search and ranking, closest matches, the number rule,
  nicknames, form rows, validation, the duplicate guard, the learning loop,
  and the export flow, including the failure path and `.xlsx` escaping.
- **They can't prove real Slack or Google behaviour**, only the logic.

## Limitations: read before using

- **Slack's 3-second limit:** Apps Script can start slowly, so Slack
  occasionally shows a timeout, most often when opening the form. Running the
  command again works. A submit that times out has usually saved anyway (the
  duplicate guard covers a retry). The only full fix is a host that never
  sleeps, e.g. a Cloudflare Worker in front, or an always-on bot using Slack
  Socket Mode.
- **Deprecated verification token:** Apps Script can't read request headers,
  so it can't check Slack's modern request signatures. It checks Slack's
  older verification token instead. That's weaker, and if Slack ever stops
  sending the token, every request will be rejected.
- **"Anyone" access:** Slack's servers aren't signed into Google, so the web
  app must allow anonymous access. Requests without the token are ignored.
  The script can only search the product list, add Log rows, and post to the
  channel.
- **Runs as the deployer:** if that Google account is disabled, the bot stops.
- **Sheet editors can see the tokens:** anyone with edit access to the Sheet
  can open the bound script and its Script Properties. Keep edit access to
  the people who need it.
- **Cache size:** Apps Script caches up to 100 KB per item. A product list
  plus nicknames bigger than that still works, but searches read the sheet
  every time and get slower.

## Maintenance gotchas

- **A code push doesn't update the live bot.** The web app is pinned to a
  version. After changing code, use Deploy → Manage deployments → edit → New
  version. "New deployment" creates a new URL that Slack won't use.
- **A stale editor tab overwrites changes.** If you edit the code with
  [clasp](https://github.com/google/clasp) while an Apps Script editor tab is
  open, running or saving anything in that old tab writes the old code back.
  Reload the tab after every push.
- **Run `setup` after updates** that use a new Google service, so its
  permission gets approved. It's always safe to rerun.
- **File → Import doesn't trigger a cache refresh.** Use the sheet menu
  *Inventory Adjust → Refresh parts list* after importing products.
- **Don't rename tabs or headers.** The `Log` tab uses column positions, so
  only add columns to the right of `Exported`.

## License

MIT. See [LICENSE](LICENSE).
