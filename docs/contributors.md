# Contributors

The people and tools that worked on ClerdoGrid are listed in
[`contributors/contributors.json`](../contributors/contributors.json). The app shows the list in
**File > About ClerdoGrid**, together with the version, the licence, links to the source code,
releases and issues, and the third-party notices from [`NOTICE`](../NOTICE).

## The file

```json
{
  "contributors": [
    { "name": "Clerdo", "role": "Maintainer", "url": "https://clerdo.app", "pinned": true, "commits": 2 }
  ]
}
```

| Field | |
| --- | --- |
| `name` | Shown in the app |
| `role` | Optional, shown next to the name |
| `url` | Optional `https:` link; other links are ignored |
| `github` | Optional GitHub login, filled in by the script |
| `pinned` | Kept at the top and never removed by the script |
| `commits` | Filled in by the script; not shown in the app |

Email addresses are never stored.

## Refreshing it

Maintainers refresh the list before a release:

```bash
npm run contributors                 # git history plus the GitHub contributor list
npm run contributors -- --offline    # git history only
npm run contributors -- --check      # exit 1 if the file is out of date, write nothing
```

The script (`scripts/contributors.mjs`) reads commit authors and `Co-authored-by` trailers with
`git shortlog`, merges the repository's contributor list from the GitHub API when it can reach it
(`GITHUB_TOKEN` raises the rate limit), skips bots, and keeps entries added by hand. Different
spellings of the same contributor are merged, for example every "Claude ..." co-author becomes
"Claude (AI assistant)". Without a network connection it falls back to git history alone.

The app reads the file at build time; it never fetches anything at run time.

## Adding yourself

Add an entry by hand in your pull request, or leave it to the maintainers' next refresh.
