# My Work

A library of my public projects: tools I've built to solve real problems,
mostly around hardware, production, and day-to-day operations. Each project
lives in its own folder with its own README, license, and (where it makes
sense) tests.

## Projects

| Project | What it does | Built with | Status |
|---|---|---|---|
| [slack-inventory-adjust](slack-inventory-adjust/) | A Slack `/adjust` form that logs parts taken out of inventory into a Google Sheet with exact SKUs. Searchable part picker, closest-match suggestions, self-learning nicknames, multi-item forms, and an `.xlsx` export | Google Apps Script, Slack API, Google Sheets | Pilot |

## How I work

- **Real problems first.** Most of these started as a repetitive or
  error-prone task at work or at home. The README for each project explains
  the problem, the requirements, and why I chose the approach I did, including
  what I ruled out.
- **Honest about AI.** I use AI coding assistants (mainly Claude) as part of
  how I build. When a project was built that way, its README says plainly who
  did what: what the AI wrote, and what I did, such as defining the problem,
  making the design calls, setting up the integrations, and testing in the
  real world.
- **Limitations are written down.** Every project lists what it doesn't do
  well, and what to watch out for when maintaining it.
- **Nothing private.** Projects that started at work are anonymised: no
  company names, real data, credentials, or internal links. Sample data is
  made up.

## Status labels

| Label | Meaning |
|---|---|
| Idea | Planned or designed, not built |
| Prototype | Built and works on my machine, not used for real |
| Pilot | Used for real by a small group, still being proven |
| In use | Used day to day |
| Archived | No longer maintained, kept for reference |

## Layout

```
My-Work/
├── README.md                  this index
└── <project-name>/            one folder per project
    ├── README.md              problem, how it works, setup, limitations, credits
    ├── LICENSE
    ├── src/                   the code
    └── tests/                 tests and sample data, where it makes sense
```

## License

Each project carries its own license in its folder. Unless a project says
otherwise, it's MIT.
