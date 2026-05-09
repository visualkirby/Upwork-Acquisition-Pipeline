# Upwork Acquisition Pipeline - Session Log

---

## Session: 2026-05-09

### What Was Done

**Full system audit and upgrade cycle -- 5 commits to visualkirby/Upwork-Acquisition-Pipeline**

#### Starting point (from prior session)
- LinkedIn Post 6 confirmed live; rollout-plan.md updated (Posts 4, 5, 6 all POSTED)
- Session log rows S022-S033 read; upgrade requests extracted from Notes column
- Full sheet structure (18 sheets), monolithic apps script (2340 lines, 16 sections), and
  formula export all analyzed
- 4 confirmed bugs, 5 structural issues, 14 formula issues identified
- 55 APPLY jobs in Job_Scoring with no proposals sent (104 APPLY total, 49 sent)

#### Commit 9049c0c: Script split into 14 modules
Monolithic `upwork_acquisition_system.gs` split into numbered files under `apps-script/`:
- 01_Menu.gs, 02_Setup.gs, 03_Helpers.gs, 04_Reset.gs, 05_AI_Context.gs
- 06_Quick_Notes.gs, 07_Workflow_Analyzer.gs, 08_Keyword_Mining.gs, 09_Bid_Engine.gs
- 10_Proposal_Generator.gs, 11_Job_Classifier.gs, 12_Session_Management.gs
- 13_Snapshot.gs, 14_Edit_Trigger.gs

One bug fixed in the split: `AI_Proposal` column reference in the APPLY auto-trigger
corrected to `AI_Generated_Proposal` (was silently failing every time).

#### Commit 1342ed5: Three bug fixes (12_Session_Management.gs, 14_Edit_Trigger.gs)
1. **Replenishment doubling (S029)**: `Connect_Replenishment` onEdit now guards on
   `e.oldValue` -- only accumulates into `Total_Connects_Purchased` when entering a
   fresh value into a blank cell. Editing an existing value no longer double-counts.
2. **Skip proposals not counted (S025/S017)**: Edit trigger now stamps `Proposal_Skip_Date`
   when `Proposal_Status` = "Skip". END_SESSION uses that date to count skips within the
   session time window (same logic as sent proposals).
3. **Connects returned (S024)**: New `Connect_Returned` metric handler accumulates into
   `Total_Connects_Returned` and stamps `Connect_Returned_Date` with the same
   idempotency guard as replenishment.

Requires 3 new rows in Connects_Helper sheet: `Connect_Returned`, `Total_Connects_Returned`,
`Connect_Returned_Date`. Also requires `Proposal_Skip_Date` column in Proposal_Generator.

#### Commit 54ff251: Formula fixes (15_Formula_Fixes.gs)
One-time `APPLY_FORMULA_FIXES()` function under System Tools menu. Fixes:
- **F4a**: Affordability_Check rows 3+ used `$B$14` (Total_Connects_Used=552) instead
  of `$B$16` (Current_Connect_Balance=88). All rows now use `$B$16`.
- **F4b**: Final_Decision gate `IF(AF2=0,"SKIP",...)` never fired because AF contains
  text. Fixed to `IF(AF2="Cannot Afford","SKIP",...)`.
- **F7**: Tool_Detected in Proposal_Generator missed Google Sheets, SQL, Python. Added
  before the fallback.
- **F9**: Portfolio_Project formula referenced two non-existent dashboards (Customer
  Service, Email Marketing). Replaced with the 3 real project full names and expanded
  from 5 to 16 keyword triggers.

Must run once after deploying the new script files.

#### Commit 439af40: Session pattern analysis (16_Session_Analysis.gs)
`ANALYZE_SESSION_PATTERNS()` under System Tools menu. Reads all Session_Log rows and
produces time-of-day and weekday aggregate reports (yield, proposals, connects, duration,
saturation count). Works retroactively from existing Date/Start_Time columns.

Verified against S022-S033:
- Morning (6-12): 4 sessions, avg yield 9.2, avg props 2.5
- Afternoon (12-17): 6 sessions, avg yield 8.3, avg props 1.7
- Evening (17-21): 2 sessions, avg yield 9.0, avg props 1.5, 1 saturation

#### Commit 84e5097: Bid pattern analysis (17_Bid_Analysis.gs)
`ANALYZE_BID_PATTERNS()` under System Tools menu. Reads Proposal_Generator rows with
bid data and reports competition levels and boost patterns.

Verified against 44 rows with bid data (S022-S033):
- Low (0-9): 8 jobs, 88% send rate
- Medium (10-29): 15 jobs, 87% send rate
- High (30-49): 6 jobs, 50% send rate
- Extreme (50+): 15 jobs, 20% send rate -- 3 sends flagged (Bid1=58, 70, 81)
- 13 boosts recorded; 1 EXCESSIVE (123%, correctly skipped), 1 HEAVY (86%, sent)
- Send accuracy: 77% of sends went into low/medium competition

Thresholds are named vars at top of file for easy tuning.

### Pending Manual Steps (to deploy)
1. Copy all 17 `.gs` files from `apps-script/` into Google Apps Script editor
2. Run `System Tools > Apply Formula Fixes` once
3. Add to Connects_Helper sheet: rows `Connect_Returned`, `Total_Connects_Returned`,
   `Connect_Returned_Date`
4. Add `Proposal_Skip_Date` column to Proposal_Generator sheet

### Key Data Points
- 104 APPLY jobs in Job_Scoring total; 49 proposals sent; 55 with no proposal
- With F4b fixed, some of those 55 may flip to SKIP (those requiring connects > balance of 88)
- Current Connect_Balance: 88 (row 16 of Connects_Helper)
- Total_Connects_Used: 552; Total_Proposals_Sent: 49
- 3 extreme-competition sends (Bid1=58/70/81) account for wasted connects
