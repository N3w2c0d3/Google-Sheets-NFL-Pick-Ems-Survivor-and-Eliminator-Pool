# Multi-Pool Extension — custom features on top of Ben's template

This is a **fork** of [boilermaker-ben/Google-Sheets-NFL-Pick-Ems-Survivor-and-Eliminator-Pool](https://github.com/boilermaker-ben/Google-Sheets-NFL-Pick-Ems-Survivor-and-Eliminator-Pool).

**All credit for the original belongs to Ben Powers.** The form generation, weekly sheets, ESPN
schedule import, member management and everything else that makes this work is his. If you're
looking for the actual tool, [start there](https://github.com/boilermaker-ben/Google-Sheets-NFL-Pick-Ems-Survivor-and-Eliminator-Pool)
— and [buy him a coffee](https://www.buymeacoffee.com/benpowers) while you're at it.

What's here is one commissioner's extension of it: instead of a single weekly pick 'em, it runs
**seven simultaneous pools** off the same set of picks, tracks the money for all of them, and
adds some tooling for running a season without doing it by hand every Tuesday.

> **Based on** upstream commit [`d400f32`](https://github.com/boilermaker-ben/Google-Sheets-NFL-Pick-Ems-Survivor-and-Eliminator-Pool/commit/d400f3259639b55a4458b27d187ce6eb0ae9a1fb) (19 Aug 2026).
>
> **Status:** new for the 2026 season. The scoring and payout logic is unit-tested, but it has
> not yet been through a full live season. Treat it accordingly.

---

## What it adds

### Seven pools off one set of picks

| Pool | Cadence | How you win |
| --- | --- | --- |
| **Weekly Wins** | weekly | Most points that week. Multiple places paid, ordered by a two-step tiebreaker |
| **Weekly Consensus** | weekly | Beat the crowd — strictly more correct picks than the group's majority-vote entry |
| **Season % Correct** | end of regular season | Your best 16 of 18 weeks by percentage, pooled |
| **Season Total Points** | weekly → season | 5/4/3/2/1 per week by finish, summed over the season |
| **Season Consensus** | end of regular season | Beat the crowd's cumulative season total |
| **Post Season** | playoffs | 5/4/3/2/1 per playoff week, one cumulative total |
| **Perfect Week** | any week | Get every game right. Each perfect week is a share of the pot |

### A two-step tiebreaker on the right game

The weekly tiebreaker asks two questions — the **combined final score** of the tiebreaker game,
then the **winning team's score** — and it's pinned to the **late Monday night game** (the later
one when a week has two, falling back to the last kickoff of the week in the playoffs).

The game is chosen once and stored on the week's game plan, so the form, the import and the
outcome writer all act on the same game rather than each deciding independently.

### A consensus ("beat the crowd") entry

Each weekly sheet grows a **`Consensus pick`** row holding the group's majority pick per game,
and a **`🤝 Consensus`** column flagging `1`/`0` for whether each member beat it. Matching the
crowd exactly doesn't count as beating it.

Two edge cases are handled separately, which is easy to conflate:

- **A game the NFL ties** is a freebie — every member *and* the consensus bank a point, so it
  can't change who beats the crowd.
- **A 50/50 pick split** leaves the crowd with no pick for that game (shown as `SPLIT`), so it
  simply scores nothing there while members score normally.

### A money engine

Every pot is derived — `fee × entries × weeks` — so pots self-correct as members join instead of
being typed in and going stale. There are no hardcoded dollar amounts anywhere in the code.

**Ties consume the places below them.** A three-way tie for first takes the percentages for
places 1, 2 and 3, adds them together, and splits the total three ways; the next finisher is
then 4th. Payouts round to whole cents with the residual assigned to the largest payout, so the
dollars paid always reconcile exactly to the dollars collected.

The number of paid places scales with the size of the pool (3 places under 15 entries, up to 7
above 35), and the percentage ladder for each is a single editable constant.

Season-long pools show standings and points all season but **only assign dollars once their
period closes** — so any dollar figure in the workbook is money actually won, never a
projection.

### Tabs it builds

`Standings` · `Wkly Payout` · `Wkly Consensus` · `Season Points` · `Season Totals` ·
`Playoff Points` · `Summary Payout` · `Wkly Rank` · `Season Rank` · `Rank Jump Chart`

`Standings` is the member-facing one — sorted by money won, with each pool's position — and it
lands as the leftmost tab, since the whole workbook typically gets shared read-only with the group.

`Summary Payout` is the commissioner's view: winnings per pool per member, plus entry, a manual
fines column, and the fee schedule block that defines every pot.

`Season Totals` shows the best-16 maths with each member's **dropped weeks faded out**, so anyone
can check the arithmetic themselves.

`Rank Jump Chart` plots everyone's rank across the season with the axis inverted, so first place
sits at the top.

### A standalone sign-up form

Ben's template enrolls people through a weekly form's "New User" page, which means joining
requires submitting Week 1 picks — no use in August. This adds a separate sign-up form asking
for a team name, the member's real name and an email address, and **no picks at all**, so it can
go out weeks before kickoff.

`Import Sign-Ups` shows you exactly who's about to be added, skips duplicates and blanks, and
writes nothing until you confirm.

### Commissioner tooling

| Menu item | What it does |
| --- | --- |
| **✅ Close Out Week** | The whole end-of-week routine in one action: pull final scores, refresh formulas, recompute every payout tab. Refuses to pay out a week that isn't fully scored |
| **🩺 Health Check** | Duplicate members, members missing from a payout tab, missing named ranges, config drift, and money reconciliation — dollars paid vs dollars collected, per pool |
| **📸 Snapshot Spreadsheet** | Timestamped full copy, also offered automatically before anything that rebuilds tabs |
| **📧 Remind Non-Respondents** | Emails whoever hasn't submitted. Manual with a confirmation list, plus an automatic send a few hours before the week's first kickoff |
| **🔒 Kickoff Lock On/Off** | Closes each new form at the week's first kickoff |
| **💵 Update Payouts** | Recomputes every money tab from the weekly sheets |

### Robustness changes to the base

- ESPN calls go through one wrapper with browser-like headers and retry-with-backoff on
  transient 403/429/5xx, instead of throwing on the first failure.
- Three bugs in the upstream template were fixed along the way and reported back to Ben:
  a `Ui.alert` signature error that masked real exceptions in the form builder, an undeclared
  variable in `buildPickemQuestions`, and the tiebreaker being derived independently in two places.

---

## Configuring it for your own pool

1. **Set your fees.** The `POOLS` constant near the top of `picks.gs` ships with **$1
   placeholders**. Change the `fee` values to your own. Delete the line for any pool you don't
   run — its column disappears from `Summary Payout` and its pot stops being collected.
2. **Set your payout ladder.** `PAYOUT_LADDERS` and `PAYOUT_PLACE_BREAKPOINTS` control how many
   places get paid at what percentages, by entry count.
3. Other constants worth knowing: `SEASON_PCT_WEEKS` (best-16), `PLAYOFF_WEEKS`,
   `REMINDER_HOURS_BEFORE_KICKOFF`.

## Assumptions baked in

This was built for one specific pool, so some of its choices are opinions rather than options:

- **Straight-up picks**, no spreads (`pickemsAts: false`)
- **No Survivor or Eliminator** — the base supports both; this doesn't extend them
- **One bonus game per week**, chosen manually, worth its multiplier. Weekly placement and
  Season Points rank on bonus-inclusive points; the consensus pools and Season % use raw
  correct picks
- **One entry covers every pool** — no per-pool opt-in
- Regular season weeks 1–18, playoffs 19/20/21/23

## Install

Same as the original: copy the template spreadsheet, then **Extensions → Apps Script**, and
replace `picks.gs` with the version here. The HTML panel files are unchanged from upstream except
`formCreatorPanel.html`, which is aligned with the tiebreaker change.

## License

MIT, same as the upstream project.
