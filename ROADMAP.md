# Roadmap — ATO Travel Allowance Diary

Future features and maintenance tasks, roughly in priority order.
GitHub renders the checkboxes below as a live task list — tick them off by editing this file as items ship.

## Recurring maintenance

- [ ] **FY2026-27 rates** — when the ATO publishes the new determination (expected early July 2026), append an entry to the `RATE_YEARS` array in both edition files. Additive change only; historical years stay untouched.
- [ ] **Review high-cost centre classifications annually** — a small number of towns move between the capital/high-cost and regional tiers each year. Cross-check the `LOCATIONS` list against the new determination when adding the year's rates.

## Planned features

- [ ] **Full PWA support** — add `vite-plugin-pwa` (manifest, service worker, icons) for offline use and a proper home-screen app on iPhone. Priority: enables on-the-road entry in regional blackspots.
- [ ] **Year-specific location tiers** — the location lookup currently uses the 2025-26 tier list for all years. Make the capital/regional classification date-aware like the rates (only matters for towns that changed tier between years).
- [ ] **Print-friendly trip report** — a clean single-trip summary view for attaching to workpapers, as an alternative to the full Excel export.
- [ ] **Salary band support** — rates currently assume Table 1 (salary ≤ threshold). Add Table 2/3 selection if remuneration ever crosses the band.

## Ideas (not committed)

- [ ] Optional venue note per meal — not required by the ATO, but useful if a meal exceeds its cap and needs receipt-based substantiation instead
- [ ] Cross-device sync beyond manual JSON exchange (would require a backend or private gist — evaluate whether the JSON workflow is actually painful enough to justify it)
- [ ] Merge the two editions into one app with a model toggle (percentage vs daily cap) instead of separate builds

## Done

- [x] Two-phase model: pre-trip allowance determination + post-trip expenditure diary
- [x] Percentage-of-ATO-rate allowances (removes arbitrary daily cap, scales by location)
- [x] Max claimable column with per-meal ATO capping
- [x] Year-aware rate engine (TD 2024/3 + TD 2025/4, auto-selected per day) — both editions
- [x] Financial year selector filtering trips, stats, and exports
- [x] JSON backup split per financial year, with merge-safe import
- [x] Excel export with full determination → expenditure audit trail
- [x] GitHub Pages deployment
