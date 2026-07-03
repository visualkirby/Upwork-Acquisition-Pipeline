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

---

## Session: 2026-05-09 (continued)

### What Was Done

#### F4a header name confirmed and fixed (15_Formula_Fixes.gs)
Column is `Connects_Affordability` (not `Affordability_Check`). Added as first lookup
in `getCol_` call. Re-running `Apply Formula Fixes` will now apply all 4 fixes cleanly.

#### Connect_Returned entry instructions provided
Three rows to add to Connects_Helper: `Connect_Returned` (input field, leave blank),
`Total_Connects_Returned` (start at 0), `Connect_Returned_Date` (auto-stamped).
Enter the number of returned connects in `Connect_Returned`; trigger accumulates and dates.

#### Commit c317e7f: Pipeline funnel replaces per-job workflow analyzer (07_Workflow_Analyzer.gs)
`ANALYZE_JOB_WORKFLOW` rewritten as a 4-stage pipeline funnel:
- Stage 1: Job_Discovery -- Discovery_Action (Move to Scoring / Review Later / Other)
- Stage 2: Job_Scoring -- Final_Decision (APPLY / HOLD / SKIP)
- Stage 3: Proposal_Generator -- Proposal_Status (Sent / Skip / Ready)
- Stage 4: Proposal_Tracker -- Hired / Interview / Client_Replied (Y/N counts)
- End-to-end conversion rates at every stage and overall

`getWorkflowAnalysis_` (per-job AI breakdown) kept in file for future wiring.

### Pipeline State as of 2026-05-09
```
Discovery:   350 logged  ->  173 Move to Scoring (49%)  |  177 Review Later (untapped)
Scoring:     173 scored  ->  104 APPLY (60%)  |  46 SKIP (27%)  |  23 HOLD (13%)
Proposals:   104 queued  ->   41 Sent (39%)   |  62 Skip (60%)  |   1 Ready
Outcomes:     49 tracked ->    1 Hired (2%)   |   1 Interview    |   1 Reply
Overall:     0.3%  (1 hire / 350 logged)
```

Key gap: 60% of APPLY jobs are being skipped in Proposal_Generator (no proposal written/sent).
177 Review Later jobs are a re-evaluation pool that never reached scoring.

---

## Session: 2026-05-09 (continued 2)

### What Was Done

#### All 4 formula fixes confirmed applied
F4a (Connects_Affordability $B$16), F4b (Final_Decision gate), F7 (Tool_Detected),
F9 (Portfolio_Project) all reported Applied. Pipeline is now on correct formula logic.

#### API key property name fixed across all modules (commits 1a2fda2, 4bbf2de)
- SETUP_API_KEY rewritten to use ui.prompt() dialog -- no more code editing required
- CHECK_API_KEY added to menu for status verification
- Root cause: all modules were reading `OPENAI_API_KEY` but property was stored as
  `UPWORK_OPENAI_API_KEY`. Fixed in 6 files: 02_Setup, 06_Quick_Notes,
  07_Workflow_Analyzer, 09_Bid_Engine, 10_Proposal_Generator, 11_Job_Classifier

#### Skip rate root cause diagnosed (no code change yet)
Cross-referenced formula export + sheet structure against Proposal_Generator behavior.
Three compounding causes:
1. Final_Decision APPLY gate has no competition cap -- high Proposal_Count jobs
   (40-50+) pass scoring and land in Proposal_Generator, then get manually skipped
2. Current_Age_Days in Proposal_Generator samples show 50-65 days -- jobs are ancient
   by the time proposals are written; user correctly skips them
3. FILTER formula accumulates all-time APPLY jobs with no expiry mechanism

Proposed fix (not yet built): add F10 to APPLY_FORMULA_FIXES -- two new gates in
Final_Decision: `J2<=35` (Proposal_Count cap) and `AL2<=21` (age cap in days).
This would auto-remove stale rows from Proposal_Generator via the live FILTER.

#### SaaS architecture planned -- full doc saved locally
File: `C:\Users\kirby\OneDrive\Desktop\ClaudeCodeTest\upwork-saas-poc-plan.txt`

Key decisions made:
- Web app + smart paste as POC; mobile (Expo share extension) as Phase 3
- FastAPI + Supabase + React stack
- Scoring thresholds fully configurable from UI (preset profiles + individual overrides)
- Browser extension ruled out; OS share sheet is better UX for mobile
- All formula logic moves to Python backend services
- FILTER pipeline replaced by event-driven row writes

Open decisions (in plan doc): pricing, product name, smart paste mode (URL vs text),
portfolio matching approach, demand validation channel, build vs. hire.

This product is the third Benchline Analytics SaaS product (was TBD on checklist).
Full plan + screen list saved to `upwork-saas-poc-plan.txt` in ClaudeCodeTest.

### What Is Next
- Build F10 formula fix (competition + age gates) for Final_Decision
- Tackle SaaS product planning next session (add to May checklist as third product)
- Re-evaluate 177 Review Later jobs as pipeline refill pool

---

## Session: 2026-07-03

### What Was Done

**Discovered the git repo had drifted badly from the live product.** The real, actively-deployed source is `C:\Users\kirby\OneDrive\Desktop\ClaudeCodeTest\freelanceflow-template\` (clasp-connected to script ID `1bYZnJPfm5nqnfzFcZjhRym7HryuHdYP_uqjz75mko75EoOoUDx0sUAuK`) -- it has a full setup wizard (`00_Setup_Wizard.gs` + `SetupWizard.html`) that this repo's `apps-script/` never had, plus genericized/Settings-driven versions of 05, 09, 10, 12 that had diverged from what was committed here. Product has also been named **FreelanceFlow** (beat ProposalPilot/BidForge/AcquireIQ) and is launching at $47/$127 via Gumroad.

**Split the Apps Script into a Library + thin client** (the actual work requested: customers who buy the template get a full copy of the bound script today, including every AI prompt and the wizard's formula-building mechanics -- the goal was to stop giving that away).

1. Synced `apps-script/` to match the real template source; removed the stale pre-split monolith `upwork_acquisition_system.gs`; fixed README's dead link + wrong `apps_script/` directory name.
2. Refactored every AI/config function to accept `apiKey`/`settings`/`journeyContext` as parameters instead of fetching internally (`PropertiesService`/`getSettings_()` calls moved to call sites) -- prerequisite for the Library split since a Library has its own separate property store.
3. Extracted pure functions: `mineTriplets` (08, taxonomy + mining algorithm), `bucketBidPatterns`/`fmtBucketLine` (17, competition thresholds), and split `00_Setup_Wizard.gs`'s formula writers into pure `build*` functions (data in, formula-string out) vs. `apply*_` sheet-I/O wrappers.
4. Created a new standalone Apps Script Library project ("FreelanceFlow-Library", script ID `1tfy21pviJ9oMStv7Ur2nHCxKt6USYO1SU3NWWHZJapbT7Ra_7OTEGk55`), pushed via clasp, cut as **version 1** (pinned, not dev/Head mode). Holds: `Lib_AIContext`, `Lib_QuickNotes`, `Lib_WorkflowAnalyzer`, `Lib_JobClassifier`, `Lib_BidEngine`, `Lib_ProposalGenerator`, `Lib_Mining`, `Lib_BidAnalysis`, `Lib_WizardFormulas`, `Lib_WizardAI`. **Important Apps Script detail learned**: trailing-underscore function names are private to a Library and cannot be called cross-project -- every function the client needs to call had to be renamed without the underscore (e.g. `getJourneyStage_` -> `buildJourneyStage`); purely-internal Library helpers (`buildPortfolioFormula_`, `colLetter_`, `getQuickNotesRegex_`) kept the underscore on purpose.
5. Stripped the proprietary bodies out of the thin client, replaced every call site with `FFLib.*`. Verified via grep sweep -- zero leftover local references to any moved function, and every `FFLib.*` name matches an actual Library export.
6. Added `appsscript.json` Library dependency (pinned to v1) to the template's manifest.
7. Repurposed the previously-inert `15_Formula_Fixes.gs` into `REPAIR_FORMULAS()` (re-applies Library-built formulas + copies down to existing data rows -- for when a Library fix ships after a customer's sheet was already set up) and `SEND_DIAGNOSTIC_REPORT()` (customer-run, shows/emails current Settings + actual formula text to `support@benchlineanalytics.com`, confirmed correct). Both wired into the System Tools menu.
8. Corrected a wrong assumption made mid-session: the Job_Scoring/Proposal_Generator scoring formulas were NOT converted to hidden Library-computed values, because they're already Settings-sheet-driven (Conservative/Standard/Aggressive profiles) as an intentional customer-tunable feature -- hiding them would have broken that.
9. Repo now tracks both `apps-script/` (thin client) and a new `library/` folder (Library source) -- all committed and pushed (4 commits: sync baseline, the Library split, a placeholder-comment cleanup, and removing an unused empty `Lib_Helpers.gs`).

**Set up live browser testing.** Added `.mcp.json` (Playwright MCP server) and a temporary `mcp__playwright__*` allow rule in `.claude/settings.local.json` for this session only -- reverted (moved to `deny`) at Session End per standing instruction. User is restarting Claude Code now to pick up the new `.mcp.json` (a session that predates a `.mcp.json` file doesn't discover it via `/mcp` alone).

**Built a full test plan for the live "fresh copy" verification (Task 8, not yet executed):** a test persona ("Maria Chen", bookkeeper/QuickBooks niche -- deliberately different from Sawandi's own) with exact Setup Wizard inputs and pre-computed expected Settings values under the Aggressive profile; 4 test job rows (Job_Scoring) chosen for full Final_Decision branch coverage (APPLY, HOLD, SKIP-via-low-score, SKIP-via-Cannot-Afford) with expected Tool_Detected/Portfolio_Project outcomes; a planned bug-injection/fix/repair cycle (flip `<=` to `>=` in the Final_Decision connects comparison in a Library v2, observe Job A flip from APPLY to SKIP, fix in v3, confirm REPAIR_FORMULAS restores APPLY) to test the whole diagnostic/repair loop end-to-end.

### What Is Next
- User restarting Claude Code, will run `/mcp` again to connect the Playwright server
- Once connected: execute the fresh-copy test plan (Phase 1 Setup Wizard with Maria Chen persona -> Phase 2 four test job rows -> Phase 3 Send Diagnostic Report) and report back
- Once Phase 1-3 pass: Claude pushes the intentional Library v2 bug, user observes the break, Claude pushes the v3 fix, user runs Repair Formulas to confirm
- Final step of Task 8: open the fresh copy's Apps Script editor and manually confirm no proprietary logic is visible (only `FFLib.*` calls), and confirm the copy owner cannot open the Library project itself
- Loose end: leftover unused `library/Lib_Helpers.gs` already removed this session (done, no longer open)
