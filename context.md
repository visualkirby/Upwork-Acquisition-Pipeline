# Upwork Acquisition Pipeline - Session Log

---

## Session: 2026-07-15

### What Was Done

**Full Walkthrough demo script (`Benchline_Analytics_Infrastructure_Setup.docx`, Section 4) updated: new "Interview Chat Logging" segment added, script retimed.** Sawandi wanted a piece on the Client Chat Log sidebar inserted between the existing "Log Proposal Bid" (4:30-6:00) and "Hired and Ongoing Tracking" (6:00-7:30) segments. Reviewed `19_Chat_Import.gs` to confirm actual behavior before drafting copy -- the sidebar auto-opens when Proposal_Tracker's Interview flag flips to Y, and lets the user paste a raw copied Upwork chat transcript, which AI splits into one row per message in `Client_Chat_Log` (sender/direction/time tagged, keyed to Discovery_ID). Went through several wording rounds on the opening line; landed on "Once your proposal gets a response, the Client Chat Log sidebar opens automatically..." after Sawandi corrected an earlier draft that implied the client has to take an action to "move" the freelancer to interview status -- Upwork counts any reply as an interview automatically.

Final retimed script (handed to Sawandi as copy-paste text -- the doc itself was not edited directly, no write access to it this session):
- New: 6:00-7:00 Interview Chat Logging
- Hired and Ongoing Tracking shifts to 7:00-8:30
- Performance Analysis shifts to 8:30-9:30
- Close shifts to 9:30-10:30

**Launch date locked: Friday, July 17, 2026.** Beta tester copy being sent to the LinkedIn beta tester today (7/15) -- Sawandi confirmed all planned tests are complete, no further re-verification needed.

**Full social launch plan drafted and saved** to `G:\My Drive\Benchline Analytics\Important_Plans\FreelanceFlow_Launch_Posts_2026-07-17.docx` (new file, built via python-docx since Write can't produce binary docx directly): LinkedIn launch post (story format, ~400 words, several wording rounds to fit brand voice.md rules -- no em dash, "you" not "we"), Reddit launch post for r/freelance + r/upwork plus a separate r/digitalnomad variant (location-independence angle instead of Upwork-mechanics angle, per Sawandi's request), an 8-tweet Twitter/X launch thread, and two Indie Hackers posts (a 7/16 building-in-public intro, a 7/17 launch-day post). Researched best-day-to-launch data via WebSearch before finalizing the schedule (Product Hunt: Tue-Thu best for raw traffic, Fri/Sat/Sun better odds for the #1 badge due to less competition; LinkedIn: Wed best day, Thu and Fri close behind, Mon/Tue worst) -- Friday 7/17 confirmed as a solid choice, not the mathematically optimal LinkedIn day but close, with no conflicting data.

**Dated schedule established:**
- Wed 7/15 (today): send beta copy, LinkedIn teaser post (not yet drafted)
- Thu 7/16: Indie Hackers intro post (drafted), Twitter/X teaser (not yet drafted)
- Fri 7/17 (launch day): all 6 drafted platform posts go out
- Sat-Sun 7/18-19: engage-only, no new posts
- Mon-Thu 7/20-23: LinkedIn follow-up post, Twitter build-in-public updates, Indie Hackers milestone update (none drafted yet)
- Week 4: ProductHunt launch (per existing master plan)

**Answered two of Sawandi's status questions by checking the actual code/docs rather than assuming:**
- Repair Formulas question: confirmed it's only needed on copies that ran the *old* Setup Wizard before Dashboard/Hourly_Log Status existed (pre-2026-07-08). Any copy running Setup Wizard fresh now gets both automatically -- no Repair Formulas needed there.
- onEdit-dispatch-gap audit (open item from 2026-07-08): checked `14_Edit_Trigger.gs` directly. Contract_Tracker and Hourly_Log already use the safe `e.value`/`e.oldValue` delta pattern (confirmed at the milestone-status and hourly-log blocks). Job_Discovery's keyword-count increment (lines ~63-76) still uses a live-read "check if Date_Found is blank, then set" guard -- same race shape as the original bug, never rewritten. **Not fully resolved** -- Job_Discovery's guard is the one remaining unaudited spot.

**Later the same day: Job_Discovery's keyword-count guard fixed, all remaining FreelanceFlow launch content drafted, and the two Infrastructure/Plan docs edited directly.** Continuing from the audit finding above:
- `14_Edit_Trigger.gs`'s Job_Discovery block rewritten so `isFirstLog` (the keyword-count guard) uses `e.oldValue` on the edited Description cell instead of a live read of `Date_Found` -- same fix shape as Contract_Tracker/Hourly_Log, immune to the same double-count race. Committed `73feb3d` and pushed. **Not yet deployed** to the 4 live Apps Script copies (Production Master, Gumroad, Personal, Demo) -- needs Sheet ID verification against Drive first per this project's documented ID-drift history, flagged as a separate step rather than done silently.
- `Benchline_Analytics_Infrastructure_Setup.docx` Section 4 actually edited this time (earlier note above about "no write access this session" is now stale) -- pasted the retimed Full Walkthrough script directly via python-docx, including the new 6:00-7:00 Interview Chat Logging segment and the shifted Hired and Ongoing Tracking/Performance Analysis/Close timings, plus updated the two "~9-9.5 min" duration labels to "~10-10.5 min". Backed up the original first; verified the full paragraph list afterward (not just the touched region) to rule out the silent-corruption risk python-docx edits carry.
- `Benchline_Product_And_Job_Search_Plan_2026-06-30.docx` updated: the vague "LinkedIn launch post + Reddit posts -- launch week" bullet under Week 3 replaced with the confirmed date and full channel list, and a new "Week 3 Check-In: July 15-17, 2026" section appended at the end (same format as the existing Week 2 Check-In block) with a day-by-day Wed/Thu/Fri breakdown.
- Today's (7/15) LinkedIn teaser and tomorrow's (7/16) Twitter/X teaser drafted and saved into `FreelanceFlow_Launch_Posts_2026-07-17.docx` as a new "Pre-Launch Teasers" section ahead of the existing launch-day content.
- Post-launch follow-up content (LinkedIn results post, Twitter build-in-public updates, Indie Hackers milestone update) drafted as templates with `[bracketed placeholders]` for real launch numbers -- deliberately not fabricated, since no launch-day data exists yet. Saved as a new section in the same Launch Posts doc, explicitly labeled as templates not to be posted as-is.
- The uncommitted 2026-07-15 log entry itself (this section) got committed (`3ac8644`) with the reset-formatting coloring answer folded in.

### What Is Next

**FreelanceFlow launch -- immediate:**
1. Post the 6 saved launch-day pieces (LinkedIn, Reddit x2, Twitter thread, Indie Hackers x2) on Friday 7/17 -- manual action for Sawandi, launch day itself, not something that can be executed ahead of time
2. Before posting the 7/20-23 follow-up content: fill in the `[bracketed placeholders]` in the LinkedIn/Twitter/Indie Hackers templates with real launch numbers once they exist -- do not post the templates as-is
3. Deploy the Job_Discovery keyword-count guard fix (`73feb3d`) to the 4 live Apps Script copies (Production Master, Gumroad, Personal, Demo) -- verify each Sheet ID against Drive first, per this project's ID-drift incident history, before pushing

**FreelanceFlow -- other status changes from today:**
- Connect Balance manual entry: Sawandi said to take this off the list (deprioritized, not done)
- `Keyword_Strategy!Actual_Count` backfill: confirmed off the list, starts from zero going forward (no action needed, informational only)
- Dashboard tab drag to position 1 on other copies: done
- Reset-formatting partial-column-coloring recheck: clarified -- the partial (A, H:I) coloring did not reappear. Resolved, no bug.
- 3 FreelanceFlow demo videos: recorded

**Still open, untouched today** -- Repair Formulas run on any pre-2026-07-08 already-set-up copies (see clarified scope above), Gmail send-mail-as confirmation, Pro tier Cal.com event (created, waiting on Gumroad's 30-day account-age restriction to lift), WooCommerce/Stripe (deferred to real revenue), Drive cleanup (needs Sawandi), PipelineIQ crossref

---

## Session: 2026-07-14

### What Was Done

**Found and fixed two dead-metric bugs while Sawandi was working on the demo script, plus a reset-formatting bug found while testing the fixes -- all tested on the Demo Copy and shipped to all 4 deployment targets + git:**

1. **Connects Balance never moved.** `Connects_Helper!Current_Connect_Balance` was only ever written once, at Setup Wizard time (`initConnectsHelper_`, `00_Setup_Wizard.gs`) -- nothing decremented it when a proposal was sent or incremented it on replenish/return, even though it feeds `Connects_Affordability` (a live APPLY/SKIP decision formula in `Lib_JobScoringFormulas.gs`) and the Dashboard's "Connects Balance" KPI tile. Fixed by wiring 3 new `incrementConnectsHelperMetric_` calls into `14_Edit_Trigger.gs`: the Sent-status path in `handleProposalStatusChange_` (decrements by the total connect cost), and the `Connect_Returned`/`Connect_Replenishment` handlers (increment).

2. **Keyword Strategy sheet was disconnected.** `Keyword_Strategy!Actual_Count` was never written by anything, so its `Recommended_Action` formula (`Complete`/`Keep Testing`/`Avoid`, built in `Lib_KeywordStrategy.gs`) was permanently dead. Added a new shared helper `incrementKeywordStrategyActualCount_(ss, keyword)` (`18_Keyword_Strategy.gs`) and wired it into both Job_Discovery entry paths: the direct-cell-paste trigger in `14_Edit_Trigger.gs` (guarded so it only fires on the first log of a row, not re-edits) and the `job_saveEntry` sidebar function (`21_Job_Discovery_Sidebar.gs`).

3. **Reset left stale duplicate-flag colors behind.** Sawandi ran a demo test session on "FreelanceFlow - Demo Copy," got a duplicate-Job_Link alert he couldn't visually confirm (checked the URL list, saw no obvious dupe), ignored it, then ran `RESET_TO_AFTER_SETUP` -- and the reset sheet came back with red/teal coloring on rows that were otherwise blank. Root cause: `clearByHeaders_` (`04_Reset.gs`) only calls `.clearContent()`, which wipes values but not formatting. `colorDuplicateJobLinks()` (`09_Bid_Engine.gs`) paints literal cell backgrounds (not conditional formatting) when it finds duplicate `Job_Link` values -- its palette starts `#E63946`/`#2A9D8F`, the exact red/teal seen. That paint survives `.clearContent()` and was left over from before the reset. Fixed by adding `clearJobDiscoveryBackgrounds_(ss)` to `RESET_TO_AFTER_SETUP`, clearing backgrounds across the same `FORMULA_PREFILL_ROWS` (500-row) range the sheet's other formatting uses. **Unconfirmed loose end:** some of the colored rows in the screenshot showed color in only columns A and H:I (not full row width), which doesn't match how `colorDuplicateJobLinks` paints (always full row or fully cleared) -- suspected pre-existing manual formatting on those two column ranges, not a script bug. The blanket clear wipes it regardless, but if it reappears after a future reset, that would confirm something else is repainting it.

Sawandi tested all three fixes on the Demo Copy and confirmed working before anything was committed.

### Backfill discussion (open, unresolved)

Sawandi wants the live sheets' historical `Current_Connect_Balance` and `Keyword_Strategy!Actual_Count` backfilled (not reset to zero/today) so the demo video doesn't look off, since both were dead before this fix. He rejected a COUNTIF-formula-then-paste-as-values approach for `Actual_Count` as inconsistent with the system's plain-value-via-script data model. He then asked Claude to just compute and write the correct values directly -- blocked by (a) no live Sheets-write tool access this session (Google Drive/Gmail/Calendar connectors are available account-wide but none can write Sheet cells) and (b) the workspace's "Calculations" standing rule (Claude verifies, doesn't assert computed results as fact). Proposed alternative: a one-time Apps Script backfill function (same precedent as `15_Formula_Fixes.gs`'s `REPAIR_FORMULAS`) that computes `Actual_Count` from real `Job_Discovery` data and writes it via script -- **Sawandi has not yet confirmed whether to build this.** `Current_Connect_Balance` has no code-path backfill option either way; it needs Sawandi to manually type in the real current Upwork balance.

Also open from that thread: whether `Job_Scoring`'s `Keyword_Search` column can ever diverge from `Job_Discovery`'s (hand-edits after the fact) -- relevant to whether `Job_Discovery` alone is a safe source of truth if the backfill function gets built.

**Resolved same day:** Sawandi decided against building the one-time backfill function -- `Keyword_Strategy!Actual_Count` starts from zero going forward rather than backfilling history. `Current_Connect_Balance` still needs the real current balance typed in manually (unaffected by this decision, no code path either way).

### Deployment

Verified all 4 deployment target Sheet IDs against Drive metadata before touching anything (title match, not trashed) -- this project has a documented history of `.clasp.json` scriptId drift causing mispushes, so stored memory IDs are never trusted on faith. All 4 checked out clean:
- Production Master Template -- `17x3oS3OLoEhuaWzN5UbgXHNUOYeDnfG0aJjOiAR7OEM`
- Gumroad copy -- `1u0uh5NhgrjXmEmFnsNOwGEo04tl1UbVMg2JywEp6iZs`
- Personal copy ("Sawandi's Upwork Pipeline") -- `1oxQ5nmykAtvlbsGkkmeTfxVIwcbVK3NgbpbpRYLxAM4`
- Loom Demo Copy ("FreelanceFlow - Demo Copy") -- `1lu7bQn2lE2_UeWxEj4_NegvCvIjlosJFRsWrV6pn-a4` (already had this session's code from Sawandi's manual test)

Pushed code to Production Master, Gumroad copy, and Personal copy via the usual swap-`.clasp.json`-scriptId-and-push pattern; `.clasp.json` restored to the Production Master default afterward. Committed to git (`9fa5ffb`, 4 files: `04_Reset.gs`, `14_Edit_Trigger.gs`, `18_Keyword_Strategy.gs`, `21_Job_Discovery_Sidebar.gs`) and pushed to `visualkirby/Upwork-Acquisition-Pipeline`.

### Key Notes

- A separate, unrelated session-log entry for the earlier "3 things" batch (Manage Projects sidebar, Ended Early + Notes, Dashboard sheet -- code already committed as `f73d1a4`) was found sitting uncommitted in this file's working tree. Left untouched/uncommitted this session since it wasn't part of today's work -- still needs a commit.
- Clarified for Sawandi what "Google MCP" actually is: the `mcp__claude_ai_Google_Drive__*` tools (plus Gmail and Calendar equivalents) are Anthropic's account-level connectors, managed in Settings > Connectors, not a project-local `.mcp.json` server -- that's why they don't show next to `playwright` in `/mcp`. None of them can write Sheet cells, which is why the backfill still needs a script-based approach.

### What Is Next

**FreelanceFlow -- immediate, from today's session:**
1. Manually enter the real current Upwork Connect Balance into `Connects_Helper!Current_Connect_Balance` on the personal copy (and any other live copy that needs it) -- no code path, pure manual entry
2. Commit the still-uncommitted `f73d1a4`-batch session log entry sitting in this file's working tree (separate from today's work)
3. If the reset-formatting fix gets exercised again, confirm whether the partial column-only (A, H:I) coloring reappears -- would confirm a separate manual-formatting source rather than a script bug

**FreelanceFlow -- carried forward, still open from 2026-07-08 batch (`f73d1a4`, fully deployed):**
1. Run `System Tools > Repair Formulas` on any other already-set-up FreelanceFlow copy (not just personal) to retrofit the Dashboard sheet and Hourly_Log Status column
2. Manually drag the Dashboard tab to position 1 on any other already-set-up copy -- new tab-order code only applies to fresh Setup Wizard runs
3. Record the 3 FreelanceFlow demo videos in OBS Studio (Promo, Setup Walkthrough, Full Walkthrough -- scripts finalized 2026-07-09/10), upload per `Benchline_Analytics_Infrastructure_Setup.docx` Section 4's checklists
4. Gmail "Send mail as" confirmation (Infrastructure doc Section 1.1)
5. Before sending access to the LinkedIn beta tester: re-verify the full fresh-account signup flow end-to-end, prepare 2-3 specific feedback questions
6. Pro tier ($127): built on Gumroad, withheld pending the 30-day account-age restriction; "FreelanceFlow Setup Call" Cal.com event type still needs creating once it lifts
7. WooCommerce + Stripe (self-hosted Gumroad alternative) -- deferred until FreelanceFlow has real revenue
8. Drive cleanup (no delete tool available, needs Sawandi): 10 throwaway test/verification copies flagged 2026-07-05, 4 superseded 2026-07-05-batch spreadsheets, old FreelanceFlow_Setup_Guide doc, 5 stale duplicate `Benchline_Analytics_Infrastructure_Setup` copies
9. Cross-reference contract/chat/revenue logic against the PipelineIQ SaaS app before that build starts (see memory `project_freelanceflow_pipelineiq_crossref.md`)
10. Not audited: whether the "Sheets sometimes never dispatches onEdit during rapid entry" root cause (found in Hourly_Log) also affects other `14_Edit_Trigger.gs` blocks (Job_Discovery, Contract_Tracker, Proposal_Generator)

---

## Session: 2026-07-08 (cont.)

### What Was Done

**Three new features built, tested on the personal copy, and shipped to all 4 deployment targets + git**, following up on "ok test done, 3 things" from Sawandi after the earlier 2026-07-08 session's reset-and-test pass succeeded:

1. **Manage Projects sidebar** (`24_Projects_Sidebar.gs` + `ProjectsSidebar.html`, new "Manage Projects" menu item) -- add/edit up to 10 portfolio projects (Upwork's real cap) post-setup, no longer stuck with whatever the wizard's Step 6 captured at signup. `projects_saveProject` enforces the 10-project cap for new projects and, critically, re-calls the existing `applyProposalGeneratorFormulas_` after every save so Proposal_Generator's `Portfolio_Project` IFS() formula picks up the change immediately -- without this, new/edited projects wouldn't affect AI job-matching until a manual Repair Formulas run.

2. **Ended Early + Notes added to Log Contract Progress** -- the Fixed/Milestone panel previously had no way to trigger Contract_Tracker's "Ended Early" status (only possible via direct cell edit) and no Notes field despite Milestone_Tracker having a Notes column in its schema all along. Extracted the existing two-prompt Ended Early logic out of `handleEdit` (`14_Edit_Trigger.gs`) into a shared `endContractEarly_(ss, discoveryId, contractType, amountReceived, wasReleased)` (`03_Helpers.gs`), reused by both the direct-cell-edit path and a new `progress_endContractEarly` sidebar function. Both Fixed and Hourly Contract Status dropdowns in `ContractProgressSidebar.html` gained an Ended Early option (reveals Amount Received / Released Y-N fields); `progress_saveMilestones` now writes Notes per-milestone alongside Status.

3. **Dashboard sheet** (new `25_Dashboard.gs` thin-client file + `library/Lib_Dashboard.gs`, "Build/Refresh Dashboard" menu item) -- refresh-on-demand snapshot: KPI tiles pulled from Connects_Helper, a pipeline funnel bar chart, and a Monthly_Performance revenue trend line chart. `ANALYZE_JOB_WORKFLOW` (`07_Workflow_Analyzer.gs`) was refactored to call the same new `readFunnelStagesFromSheets_`/`FFLib.computeFunnelStages` computation the Dashboard uses, so the two reports never disagree. Charts are removed and rebuilt on every refresh (`sheet.getCharts().forEach(...removeChart)`) to avoid duplication on repeat runs. `buildDashboardLayoutSpec()` in the Library holds the actual layout/label knowledge, per Sawandi's explicit ask to keep that in the Library rather than the thin client. Library bumped to **v23**.

**Follow-up refinements after initial confirmation:**
- 3 new guided-tour steps: "Manage Your Portfolio Projects" and "Build Your Dashboard" fire alongside the existing early keyword-strategy tips (`18_Keyword_Strategy.gs`, `GENERATE_KEYWORD_STRATEGY`); "Keep Snapshotting Monthly" fires at the end of `SNAPSHOT_MONTH_END()` (`13_Snapshot.gs`). Tour is now 13 steps total (doc comment in `03_Helpers.gs` updated to match).
- Dashboard now gets wiped by both Reset System depths: `RESET_TO_AFTER_SETUP` calls a new shared `clearDashboardSheet_(sheet)` helper (also used by `BUILD_DASHBOARD` itself, so both stay in sync on what "clear" looks like); `RESET_TO_BEFORE_SETUP` deletes it outright (added to its sheet-delete list + warning text).
- `Dashboard` moved to the first position in `reorderPipelineTabs_`'s tab order (`00_Setup_Wizard.gs`) -- only takes effect on new Setup Wizard runs, so Sawandi manually dragged the tab on his own already-set-up personal copy to match.

**All of it pushed to all 4 deployment targets** (Production Master, Gumroad copy, personal copy, Loom Demo Copy) via the swap-`.clasp.json`-scriptId-and-push pattern, restored to the Production Master default afterward each time. Library pushed once at v23 (no further Library version bump needed for the tour-step/reset/tab-order follow-ups -- pure client-side changes). Committed and pushed to `visualkirby/Upwork-Acquisition-Pipeline` (`f73d1a4`).

### Key Notes
- Confirmed `REPAIR_FORMULAS` (`15_Formula_Fixes.gs`) already calls `ensurePipelineSheets_(ss, [])` non-destructively -- this is the existing mechanism that backfills the new Dashboard sheet onto any already-set-up customer copy without a full wizard re-run, same pattern already used for Client_Chat_Log/Contract_Tracker/Milestone_Tracker/Hourly_Log previously.
- Tab reordering only fires inside `wizard_initialize` (Setup Wizard), not on Repair Formulas -- worth remembering that any future tab-order change needs a manual drag on already-set-up copies (personal copy, and eventually the other 3 deployment targets if any of them have already run the wizard rather than staying pristine).

### What Is Next
- Run `System Tools > Repair Formulas` on any other already-set-up FreelanceFlow copy (not just the personal one) to retrofit the new Dashboard sheet
- Manually drag the Dashboard tab to position 1 on any other already-set-up copy -- the new tab-order code only applies to fresh wizard runs
- Everything else still open from earlier 2026-07-08 sessions, untouched this session: OBS demo recording, Gmail "Send mail as" confirmation, LinkedIn beta tester re-verification, Pro tier Cal.com event, WooCommerce/Stripe deferred until revenue, Drive cleanup, PipelineIQ crossref, onEdit-dispatch-gap audit

---

## Session: 2026-07-08

### What Was Done

**Guided tour: Step 7 added, then extended to 10 steps total.** Added Step 7 ("Log Your Bids," fires at the end of `RUN_JOB_CLASSIFICATION`) after Sawandi flagged he didn't remember seeing Step 6 -- also baked a 👉 prefix into `showTourStep_` itself so every tour title reads as part of the tour instead of blending into ordinary alerts. Later in the session, Steps 8 ("Import the Chat") and 9 ("Log the Contract") were converted from the two sidebars' old standalone one-time info-box tips (`showWalkthroughOnce_`, removed entirely from `19_Chat_Import.gs`/`20_Contract_Setup.gs` and their HTML) into proper sequential tour steps firing right before their sidebars auto-open on Interview=Y/Hired=Y. Step 10 ("Track This Contract") added, firing the first time a contract is saved, pointing to the new Log Contract Progress sidebar.

**Root-caused and fixed a real AI-token-waste bug in the Log Proposal Bid sidebar, then rebuilt it as a 4-step wizard.** The original single-form sidebar (`22_Proposal_Generator_Sidebar.gs` + `ProposalGeneratorSidebar.html`) resubmitted every field on every save, including fields already saved from a prior visit (the form re-prefills from the sheet after each save) -- and the AI-triggering automation only checked "is this field non-blank," not "did it change," so e.g. adding Notes after bids were already entered silently re-ran the bid-recommendation and proposal-regen AI calls for the exact same values.
- Rebuilt as 4 steps -- Bids / Boost Connects / Additional Questions / Status & Notes -- each saving only its own fields via dedicated functions (`proposal_saveBids`/`proposal_saveBoost`/`proposal_saveQuestions`/`proposal_saveStatusNotes`)
- Extracted the previously-inlined Additional_Questions -> Additional_Answers automation out of `handleEdit` into a reusable `generateAdditionalAnswers_` (`14_Edit_Trigger.gs`), matching the existing `computeBidRecommendation_`/`applyBoostConnects_` pattern, so the new step could call it
- Iterated through 2 rounds of live testing feedback: first added per-field change-detection guards (compare submitted value to the sheet's current value before firing the AI call), then **removed those same guards** once testing showed step-isolation alone already fixes the resend-waste bug and the guards were just blocking legitimate re-triggers (e.g. retyping the identical bid value during a retest showed nothing happening). Final design fires whenever a step's field is provided, matching `handleEdit`'s own per-column-edit semantics -- safe because each step's payload structurally can't carry another step's fields anymore
- Added a Skip button to every step (Step 1 skips to Step 2 without saving bids; Step 4's Skip resets the wizard without saving Status/Notes, since there's no Step 5 to advance to)

**Milestone_Tracker -> Contract_Tracker auto-complete.** Once every milestone belonging to a contract reaches Released, `Contract_Tracker`'s Status now auto-flips to Completed and `Total_Released` rolls up automatically -- new `allMilestonesReleased_`/`autoCompleteContract_` helpers in `03_Helpers.gs`, wired into a new `applyMilestoneStatusEffects_` (extracted from the inline `handleEdit` block, `14_Edit_Trigger.gs`). Previously this was a fully manual step.

**New "Log Contract Progress" sidebar** (`23_Contract_Progress_Sidebar.gs` + `ContractProgressSidebar.html`, new "Log Contract Progress" menu item) -- ongoing tracking, deliberately separate from Log New Contract's one-time setup:
- Fixed contracts: per-milestone Status dropdowns; Save applies changes through the same `applyMilestoneStatusEffects_` automation as a direct cell edit
- Hourly contracts: Date/Hours/Notes entry form; "Add Entry" appends a new row each click (never overwrites), so logging hours across multiple days is just repeated clicks; shows the last 10 entries
- Bug found and fixed mid-testing: `google.script.run` silently drops its *entire* response (no error thrown, just a blank sidebar panel) when the returned object contains a raw `Date` -- fixed by sending `Log_Date` back to the client as an ISO string instead of a Date object
- `contract_saveSetup` (`20_Contract_Setup.gs`) now auto-creates a starter `Hourly_Log` row (pre-linked Discovery_ID/Job_Title) when a contract is set up as Hourly, mirroring the milestone pre-fill Fixed contracts already got

**Hourly_Log Status column added** (requested mid-session so Hourly contracts get the same sheet-column-driven status propagation `Milestone_Tracker` already has for Fixed contracts):
- New Status column (Active/Completed dropdown) added to `Hourly_Log`'s schema (`00_Setup_Wizard.gs`) plus validation
- Retrofit path added to `REPAIR_FORMULAS` (`ensureHourlyLogStatusColumn_` in `15_Formula_Fixes.gs`) -- appends the column and backfills existing rows from `Contract_Tracker`'s current Status for any `Hourly_Log` sheet created before this column existed (like Sawandi's personal test copy)
- New `applyHourlyLogStatusEffects_` (`14_Edit_Trigger.gs`) propagates a Status edit on any `Hourly_Log` row to that contract's `Contract_Tracker` Status, reusing `autoCompleteContract_` for the Completed transition
- The sidebar's "Update Status" control was rewired to write through `Hourly_Log`'s most-recent row instead of `Contract_Tracker` directly, so both entry paths (direct sheet edit or sidebar) go through one single mechanism

**Updated the FreelanceFlow demo video script** (Section 4 of `Benchline_Analytics_Infrastructure_Setup.docx`, edited directly via `python-docx` since the file is synced locally through Google Drive for Desktop at `G:\My Drive\Benchline Analytics\Important_Plans\`) -- extended from 10 to 11 minutes: added a new beat for Log Proposal Bid, renamed/expanded the old "Tracking Contracts and Revenue" beat to cover Contract Setup auto-opening plus Log Contract Progress, added a guided-tour mention to the hook, added a 4.1 prep-checklist item for a sample screening question.

**Incidental Drive finding, not acted on:** a search for `Benchline_Analytics_Infrastructure_Setup` turned up 5 duplicate Google Docs with the identical title, all much smaller than the real 207KB `.docx` and all last-modified within about an hour of each other on 2026-07-07. Read the full content of one -- confirmed it's a near-identical (pre-this-session) copy of the real doc's content, just with some markdown-artifact formatting differences, consistent with an auto-generated conversion rather than an independent edit. Safe to delete, not touched this session.

### Key Notes
- **All of this session's code changes are uncommitted and deployed only to the personal test copy.** 13 files changed/added locally (11 modified, 2 new), pushed live via the swap-`.clasp.json`-scriptId-and-push pattern to the personal copy only, restored to the Production Master default afterward each time -- never committed to git, never pushed to GitHub, never pushed to the other 3 deployment targets (Production Master, Gumroad copy, Loom Demo Copy). Sawandi is resetting and running a full test session on the personal copy as of Session End; **this is the single most important follow-up once that test passes.**
- `google.script.run` gotcha worth remembering for future sidebar work: returning a raw `Date` object (even nested inside a larger response object) silently drops the entire response client-side -- no error thrown, nothing in withFailureHandler, just nothing happens. Always serialize dates to ISO strings before returning from a server function.
- The "resend wastes AI tokens" root cause was really two separate things that looked like one: (1) a single combined form resubmitting untouched fields -- fixed by splitting into steps with structurally disjoint payloads; (2) an unrelated "did it change" guard that seemed like a reasonable extra safeguard but actually just broke legitimate re-triggers once (1) had already fixed the real problem. Worth remembering not to over-fix a problem that a simpler structural change already solved.

### What Is Next
- **Top priority**: once Sawandi's in-progress reset-and-test session passes, commit this session's changes to git and push to GitHub, then clasp-push to the other 3 deployment targets (Production Master, Gumroad copy, Loom Demo Copy) -- currently only the personal copy has any of this session's work
- Run `System Tools > Repair Formulas` on any *other* already-set-up FreelanceFlow copy (not just the personal one) once deployed, to retrofit the new Hourly_Log Status column
- Record the updated ~11-minute OBS demo per the revised Section 4 script, upload to YouTube as Unlisted, add the link to `page-freelanceflow.php` and the Gumroad listing (long-open item, script now current)
- Delete the 5 stale duplicate `Benchline_Analytics_Infrastructure_Setup` Google Docs found this session
- Everything else still open from 2026-07-07 untouched this session: Gmail "Send mail as" confirmation, LinkedIn beta tester re-verification, Pro tier Cal.com event (blocked on Gumroad's 30-day restriction), WooCommerce/Stripe deferred until real revenue, Drive cleanup of old throwaway test copies, PipelineIQ crossref, and whether the "Sheets sometimes never dispatches onEdit for a specific cell during rapid entry" root cause affects other `14_Edit_Trigger.gs` blocks

---

## Session: 2026-07-07

### What Was Done

**Root-caused and fixed why a real customer's fresh copy of FreelanceFlow never showed the System Tools menu -- two separate, stacked bugs, only fully diagnosed via a long live debugging chain with Sawandi testing each step in the browser.**

1. **`onOpen()` auto-launch bug.** `00_Setup_Wizard.gs`'s `onOpen()` called `buildSystemMenu_()` then `PropertiesService.getScriptProperties()` to auto-open the wizard on first load. Since `onOpen` is a simple trigger and a fresh copy has never been authorized by its new owner, touching `PropertiesService` there causes the whole trigger to be blocked by Apps Script's authorization gate -- not just that line. Fixed by stripping `onOpen()` down to just `buildSystemMenu_()` and dropping the auto-launch entirely; the Setup Guide already documented opening the wizard manually, so nothing was actually lost.
2. **FreelanceFlow-Library not shared publicly (the real blocker).** Even after the `onOpen()` fix, the menu still didn't appear on fresh copies. Diagnostic chain: confirmed the Gumroad-linked Sheet ID was correct, confirmed via `clasp pull` that the live code matched the fix exactly, ruled out a Google account "Verify it's you" security checkpoint, then had Sawandi check the copy's own Apps Script Executions log (0 executions logged -- the trigger wasn't even attempting to run) and manually run `onOpen` from the script editor, which surfaced the real error: `Library with identifier FFLib is missing (perhaps it was deleted, or you don't have read access?)`. Checked the Library's Drive permissions directly -- shared with literally nobody but the owner (`skirby@visualdreamland.com`). Since Apps Script must resolve every manifest-declared library dependency before any function in a project can run, this blocked the entire script for any non-owner account, which is also why the simple trigger never logged an execution attempt. **This would have blocked every real customer, not just the `onOpen` bug.** Fixed by Sawandi sharing the Library "Anyone with the link" / Viewer via the Share dialog.

**Pushed the `onOpen()` fix to all 4 deployment targets** (Production Master, Gumroad copy, personal copy, Loom demo copy) via the swap-`.clasp.json`-scriptId-and-push pattern, restored `.clasp.json` to the Production Master default afterward. Committed and pushed to `visualkirby/Upwork-Acquisition-Pipeline` (`c97834c`), bundled with the 2026-07-06 session log entry that had been sitting uncommitted.

**Created `FreelanceFlow_Setup_Guide (Updated)`** (new Google Doc, same limitation as before -- no tool exists to edit an existing Google Doc in place), merging the existing guide with two additions: a new "Getting Your Copy" section covering the full path from the Gumroad receipt email through "File > Make a copy," and an authorization-consent step inserted into "Opening the Setup Wizard" (Google's OAuth screen on first menu click, expected and one-time). Sawandi swapped the Gumroad Content tab link to the new doc and deleted the old one.

**Demo video plan changed from Loom to OBS Studio + YouTube (Unlisted).** Loom's free tier caps recordings at ~5 minutes; the drafted script runs ~10 minutes. Updated `Benchline_Analytics_Infrastructure_Setup.docx`: renamed Section 4's heading from "FreelanceFlow Loom Demo Video Script" to "FreelanceFlow Demo Video Script," changed the 4.3 checklist item from "Upload to Loom" to "Upload the final video to YouTube as Unlisted," and cleaned the now-inconsistent "for the Loom demo" wording out of the section's intro line.

**Reviewed the Infrastructure Setup doc with Sawandi and corrected a real discrepancy:** Section 3.6 showed the Pro tier ($127) listing as fully published, but the live site marks it "Coming Soon" -- Pro is built on Gumroad but not published, gated by Gumroad's 30-day account-age restriction. Split the single "Publish both listings" checkbox into a checked "Publish Template listing" line and an unchecked "Publish Pro listing (blocked by...)" line so the doc reflects reality.

**Appended a "Week 2 Check-In: July 8-10, 2026" section to `Benchline_Product_And_Job_Search_Plan_2026-06-30.docx`** with a status note (FreelanceFlow live, the two bugs above, Pro tier gated, WooCommerce/Stripe intentionally deferred until real revenue) and a day-by-day 3-day plan: Wed (record + upload demo, Gmail "Send mail as" check), Thu (embed demo link, register remaining Phase 1 staffing agencies), Fri (LinkedIn recruiter outreach, Green Belt/SNHU status check-in).

**Sawandi got an organic beta-tester lead via LinkedIn** (Bhagyasree Mallavarapu, a senior data analyst with 10 years' experience who's hit her own Upwork friction) and offered her free access via a Gumroad 100%-off discount code once the demo/guide/site polish is done. Advised him to run the full fresh-account signup flow through himself one more time before sending it to her, given the exact bug found this session, and to prepare 2-3 specific feedback questions rather than an open-ended ask.

### Key Notes
- The Library-sharing bug is the more consequential of the two fixes -- the `onOpen` fix alone would not have unblocked real customers, since the Library dependency failure blocks the entire script regardless of trigger code
- `clasp pull` into a scratch directory (rather than trusting `git`'s local state) was the key step that ruled out a stale deployment as the cause and pointed the investigation toward the Library instead
- No tool exists to change Drive sharing permissions programmatically -- the Library share fix had to be done by Sawandi manually

### What Is Next
- Verify the full customer signup flow end-to-end one more time (fresh Google account, copy, menu, authorization prompt, wizard) before sending access to the LinkedIn beta tester
- Record the OBS demo, upload to YouTube (Unlisted), embed the link in `page-freelanceflow.php` and the Gumroad listing
- Create the "FreelanceFlow Setup Call" Cal.com event type once the Pro tier's 30-day Gumroad restriction lifts
- Decide on WooCommerce + Stripe (Section 5 of the Infrastructure doc) only once FreelanceFlow has real revenue -- explicitly deferred for now

---

## Session: 2026-07-06

### What Was Done

**Gumroad product listing content drafted (not yet published) -- FreelanceFlow $47 template tier:**
- Wrote Summary + 5 feature bullets (AI Job Scoring, AI-Generated Proposals, Connect Spend Tracking, Contract and Revenue Tracking, Guided Setup Wizard) for the existing Name/Description fields
- Confirmed with Sawandi: $47 flat only for now (no Pro tier bundled -- Gumroad blocks combining a product with a service call until the account is 30 days old), no refund policy (digital template)
- Drafted receipt fields: button text "Access Your Template" (21/26 chars, chosen over "Download" since delivery is a Drive link/copy, not a file) and a custom thank-you message referencing the "File > Make a copy" + Setup Wizard flow, per the delivery method confirmed from `Benchline_Analytics_Infrastructure_Setup.docx`

**Pro tier ($127) content drafted for later, once the 30-day account restriction lifts:**
- Name/Description/Summary/feature bullets (same 5 as base plus "1:1 Setup Call"), confirmed scope is just the base template + a setup call, nothing else added
- Separate receipt custom message including a `[Cal.com booking link]` placeholder
- New Cal.com event type spec drafted for this call: "FreelanceFlow Setup Call," 30 min, Google Meet, `freelanceflow-setup` slug -- distinct from the existing Discovery Call/Product Walkthrough sales-funnel events, since this is post-purchase onboarding, not a sales call. Not yet created in Cal.com (Cal.com itself isn't set up yet).

**Shipped a real feature: Projects sheet for portfolio management.** Sawandi noticed the Setup Wizard's portfolio step (Step 6) had no field for a project description, and portfolio data lived only as scattered `Portfolio_N`/`Portfolio_N_Keywords` key-value rows in Settings.
- New **Projects** sheet (`Project_Name`, `Description`, `Keywords`) is now the single source of truth for portfolio data, created and populated by the wizard, inserted right before Settings in `reorderPipelineTabs_`'s tab order
- `SetupWizard.html` Step 6 gained a Description textarea per project row
- `00_Setup_Wizard.gs`: new `initProjectsSheet_()`; `initSettingsSheet_` no longer writes `Portfolio_N` rows; `getPortfolioMapFromSettings_` rewritten as `getPortfolioMapFromProjects_`, header-mapped off the new sheet (feeds Proposal_Generator's Portfolio_Project matching formula and Keyword Strategy generation)
- `05_AI_Context.gs`: `Portfolio_All` (the string fed into every AI proposal prompt) is now built from Projects as `Name: Description` pairs instead of just names off Settings -- the actual point of adding descriptions
- `18_Keyword_Strategy.gs` repointed to the renamed function
- `library/Lib_ProposalGenerator.gs`: fallback message updated to reference the Projects sheet instead of Settings' old Portfolio_1-5 rows -- shipped as **Library v22**
- Caught and fixed a bug this same change introduced: `RESET_TO_BEFORE_SETUP`'s hardcoded sheet-delete list and both reset functions' alert text didn't mention the new Projects sheet (`04_Reset.gs`) -- fixed so factory reset actually clears it and the soft reset's "kept" list is accurate

**Also committed leftover uncommitted work from the 2026-07-05 session** (`01_Menu.gs` menu reorder, `08_Keyword_Mining.gs` Drop Keywords purge fix) that had been clasp-pushed live back then but never made it into git -- found via `git status` showing unexpected pre-existing diffs, committed separately from today's own work.

**Deployment incident -- discovered `.clasp.json` drift, then discovered the Production Master Sheet was in Trash, then fully recreated all 4 deployment targets.**
1. Before the first push, cross-checked `apps-script/.clasp.json`'s scriptId against memory and found it had drifted back to `1bYZnJPfm5nqnfzFcZjhRym7HryuHdYP_uqjz75mko75EoOoUDx0sUAuK` -- the script ID for Sawandi's real, in-use personal Upwork pipeline (the exact mixup documented from 2026-07-05), not the Production Master. Caught before any push landed there; had Sawandi pull the 3 correct script IDs directly from each Sheet's Project Settings instead of trusting stored memory.
2. Pushed the Projects-sheet feature, then the Library v22 bump, then the reset-function fix to all 3 known targets (Production Master, personal copy, Loom demo copy) across several rounds, swapping `.clasp.json`'s scriptId each time and restoring it to the Production Master default afterward -- all confirmed successful.
3. While getting a view-only Drive share link for the Production Master (for the Gumroad Content tab), Sawandi found that Sheet sitting in his Trash -- not something either of us did directly; most likely swept up by accident during the 2026-07-05 test-copy cleanup, since several similarly-named copies were flagged for manual deletion around then.
4. Sawandi didn't trust a restore-from-trash given the accumulated confusion over which files were current, and asked for a full recreation instead of a restore.
5. Recreated all 4 deployment targets from scratch via `clasp create --type sheets`: Production Master Template (hidden, dev-only), Gumroad copy ("FreelanceFlow"), personal copy ("FreelanceFlow - Sawandi's Upwork Pipeline"), and Loom Demo Copy. Current code (including Library v22 pin) pushed to all 4, each verified live via `get_file_metadata` (correct title, not trashed) before being trusted. New IDs recorded in memory `project_freelanceflow_production_master.md`; old IDs (including the trashed old Production Master) kept in that same memory for Sawandi's manual deletion once he's spot-checked the new files.
6. New feedback memory saved (`feedback_recreate_over_restore.md`): when file/deployment state is uncertain or has drifted before, Sawandi prefers full recreation over restoring/trusting an existing file, even if that file checks out fine on inspection.

**Created two new Google Docs in `G:\My Drive\FreelanceFlow-Template\`** (no tool exists to edit the old Setup Guide doc in place, only create new ones):
- **Setup Guide (Updated)** -- full refresh of the existing guide: added the Step 6 Description field and its own note about editing the Projects sheet post-setup, updated the "what happens on launch" list and the full sheet table to current state (16 sheets), removed the stale reference to Followup_Tracker (removed back on 2026-07-04), added a line pointing to the new User Manual
- **User Manual** (new doc) -- full System Tools menu + all 16 sheets, using the 6-step guided tour's popup content as the spine for the "Your First Session" section, with additional sections for contracts/chat, AI automation, performance analysis, keyword management, portfolio management, and system maintenance (reset, diagnostics, API key) that the tour itself doesn't cover
- Old Setup Guide doc and the "(Updated)" suffix on the new one both still need manual cleanup (delete old, rename new) -- no Drive delete/rename tool available

### What Is Next
- Delete the old FreelanceFlow_Setup_Guide doc (https://docs.google.com/document/d/19dugWMeRv0nUJNJaHFfW_fce2IP-xDkPyrs9Au4vxYg) and rename "FreelanceFlow_Setup_Guide (Updated)" to the correct name
- Drive cleanup: delete the 4 superseded 2026-07-05-batch spreadsheets (old Production Master -- was in Trash, old Gumroad/personal/Loom copies; full old IDs in memory `project_freelanceflow_production_master.md`) once the new ones are spot-checked, plus the still-unconfirmed 10 throwaway test/verification copies flagged back on 2026-07-05 (never confirmed deleted)
- Get the view-only Drive share link (Share > Anyone with the link > Viewer) for the NEW Production Master (Sheet ID `17x3oS3OLoEhuaWzN5UbgXHNUOYeDnfG0aJjOiAR7OEM`) and paste it into the Gumroad Content tab with the "File > Make a copy, then follow the Setup Wizard" note
- Finish and publish the Gumroad $47 FreelanceFlow listing using this session's drafted Summary/feature bullets/button text/custom message
- Save the drafted $127 Pro tier listing for later -- launch once the Gumroad account passes 30 days old
- Create the "FreelanceFlow Setup Call" Cal.com event type (30 min, Google Meet, `freelanceflow-setup` slug) once Cal.com itself is set up (blocked on the separate, still-pending Cal.com setup in `Benchline_Analytics_Infrastructure_Setup.docx`)
- Run the Setup Wizard on the new personal copy (Sheet ID `1oxQ5nmykAtvlbsGkkmeTfxVIwcbVK3NgbpbpRYLxAM4`) with Sawandi's real info
- Record the 10-minute Loom demo walkthrough on the new Loom Demo Copy (Sheet ID `1lu7bQn2lE2_UeWxEj4_NegvCvIjlosJFRsWrV6pn-a4`)
- Update the 3 placeholder `https://gumroad.com` links in `page-freelanceflow.php` once the Gumroad listing is live
- Build the pre-launch / launch-day / post-launch plan for LinkedIn, r/freelance, and r/upwork (append to `G:\My Drive\Benchline Analytics\Important_Plans\Benchline_Product_And_Job_Search_Plan_2026-06-30.md`)
- Cross-reference the contract/chat/revenue logic against the PipelineIQ SaaS app before that build starts (`project_freelanceflow_pipelineiq_crossref.md`)
- Open item, status unconfirmed: "Bake FILTER pull architecture into Library + template" -- check with Sawandi whether still open
- Not audited: whether the "Sheets sometimes never dispatches onEdit for a specific cell during rapid entry" root cause (found in Hourly_Log) also affects other `14_Edit_Trigger.gs` blocks

---

## Session: 2026-07-05

### What Was Done

**Feature batch shipped (paused mid-build from prior session), then pushed to the live personal pipeline via git + clasp:**
- Setup Wizard auto-runs `GENERATE_KEYWORD_STRATEGY()` right after finishing (wrapped in try/catch, before `FF_SETUP_COMPLETE` is set)
- 6-step guided first-session walkthrough tour (`showTourStep_` in `03_Helpers.gs`, standalone alerts, Cancel on any step sets `FF_TOUR_SKIPPED` and silences the rest)
- Per-session popups: AI Fit Notes after each job logged, session-yield summary at End Session
- Log New Job sidebar overhaul: auto-opens after Start Session, live countdown ("N of M logged"), form clears after save, N/A placeholder on Client Name, halfway-point keyword-switch popup, auto-closes when yield target reached
- New Proposal_Generator input sidebar (`22_Proposal_Generator_Sidebar.gs` + `ProposalGeneratorSidebar.html`) -- row picker + Bid_1st-4th/Boost_Connects/Proposal_Status/Notes, since Job_Scoring needs no manual input anymore
- Pre-existing bug caught and fixed along the way: `END_SESSION`'s `prop.deleteAllProperties()` wiped every script property (API key included) every session end -- replaced with scoped `deleteProperty()` calls for only the 8 session-specific keys

**Discovered `freelanceflow-template` (the folder used as "the template" in prior sessions) is actually bound to Sawandi's own real, in-use "FreelanceFlow | Upwork Bidding Pipeline" spreadsheet, not a dedicated template.** Created a genuine dedicated **Production Master Template** via `clasp create` (Sheet ID `1pQ8AfB4I9Pjb6QCUAGvBRBvFTgQsdAvC4ezK1wQLXCs`, script ID `1wZZ6hIG-RNeD3kwPAYvXBe0f_Bk1wlqu5JaCu8mDJkSN2OMbrB_0qtP8`, local folder `freelanceflow-production-master/`). This is now the canonical source for template copies and clasp pushes going forward, alongside the personal-pipeline `freelanceflow-template/` (see memory `project_freelanceflow_production_master.md`).

**Drove the full manual 10-job walkthrough test myself** (persona Jordan Ellis, bookkeeping niche) via Playwright browser automation against a fresh copy, following `FreelanceFlow_Operations_and_Walkthrough.docx`'s Part B script exactly -- Setup Wizard, Start Session, all 10 jobs logged, scoring/classification, proposals sent, all 4 contract paths (Fixed Completed, Fixed Ended Early, Hourly Completed, Hourly Ended Early). Mid-test correction from the user: all 10 jobs had blank `Job_Link`, which gates the deterministic scoring formulas entirely -- fixed by batch-typing fake Upwork URLs into the column, confirming Job_Link truly is the trigger for `Discovery_Action`/scoring.

**Friction log compiled from the walkthrough, 4 findings:**
1. Tour Step 6 fired mid-Setup-Wizard (stacked with the keyword-strategy summary dialog) -- the wizard's own sheet creation counts as a "selection change" onto Proposal_Generator
2. Proposal_Tracker: rapid-fire "Sent" edits across multiple rows lost a row (two edits computed the same "first empty row" before either wrote) and double-incremented `MTD_Proposals_Sent`
3. Hourly_Log: rapid Tab-across-row entry left `Amount` blank and no revenue recorded for every row in the batch
4. Contract_Tracker "Ended Early" appeared to double-count revenue already recognized via a Released milestone (flagged as a code-reading concern, confirmed via live test this session)

**All 4 bugs fixed, in stages, each verified by targeted live re-tests on fresh Production Master copies (not the full 10-job script -- Setup Wizard + 2-4 seeded jobs targeting just the fixed code path):**

- **Ended Early double-count** -- added `getContractRecognizedRevenue_(ss, discoveryId, contractType)` in `03_Helpers.gs` (sums Released milestones for Fixed, all Hourly_Log entries for Hourly). The Ended Early handler now treats the entered amount as the contract's final total and only adds the delta beyond what's already recognized. Completed branch refactored to call the same helper. **Verified correct on first attempt**: released a $300 milestone, then Ended Early with $300 again -- MTD_Revenue stayed at $300, not $600.

- **Tour Step 6 premature firing** -- `onSelectionChange` now checks `FF_SETUP_COMPLETE === 'true'` before firing, since the wizard's own sheet creation was triggering it. **Verified correct on first attempt**: no premature popup during wizard init; fired correctly once setup finished and Proposal_Generator was genuinely reached.

- **Proposal_Tracker race, attempt 1 (LockService)** -- wrapped the exists-check + find-empty-row + write block in `LockService.getScriptLock()`, plus locked `incrementConnectsHelperMetric_`'s own read-modify-write. **First re-verification round found this did NOT work** -- added temporary debug logging (`console.log` calls, since removed... actually left in the throwaway test copies, not in canonical source) directly in the Apps Script web editor and proved two concurrent "Sent" executions both computed `nextTrackerRow=3` independently, ~1 second apart, despite the lock.

- **Proposal_Tracker race, attempt 2 (appendRow)** -- replaced the manual "find first empty row + 22 separate `setCellValue_` calls" with `tracker.appendRow(ptRowValues)`, where `ptRowValues` is built by looking up each field's column index via the existing header map (still no hardcoded columns). This removes the race at its root since Sheets determines the true current last row server-side at write time. **Re-verified: all 3 concurrent "Sent" edits correctly created their own tracker row, no lost rows** -- but surfaced a narrower, related bug: `MTD_Proposals_Sent`/`MTD_Connects_Used` only reflected 1 of 2 successful row creations.

- **Proposal_Tracker metric-increment race (SpreadsheetApp.flush)** -- root cause: Apps Script batches pending spreadsheet writes rather than committing them immediately, so releasing a lock right after `setValue()`/`appendRow()` doesn't guarantee the write actually landed before the next execution acquires the lock and reads. Added `SpreadsheetApp.flush()` before releasing the lock in `incrementConnectsHelperMetric_` and right after `tracker.appendRow()`. **Verified: 3 concurrent "Sent" edits gave exactly correct MTD_Proposals_Sent=4 and MTD_Connects_Used=24** (both matched baseline + 3 expected increments).

- **Hourly_Log race, attempt 1 (column-range check)** -- rewrote the handler to check whether the edited range's columns included `Hours_Logged`, looping over every row in the range and reading current sheet values instead of trusting `e.value`. **Re-verification found this still failed** -- added top-level debug logging (`console.log` of every `handleEdit` invocation's sheet/row/col/A1 range) and proved Google Sheets sometimes genuinely never dispatches `onEdit` at all for a specific cell during a fast Tab-across-row entry (Discovery_ID and Hours_Logged were silently dropped while Job_Title/Log_Date fired normally for the same rows). No in-code range check can compensate for a trigger that's never invoked.

- **Hourly_Log race, real fix (decouple recompute from edited column)** -- removed the "only if Hours_Logged in edited range" gate entirely. Now *any* edit anywhere on Hourly_Log recomputes Amount for every touched row from whatever Hours_Logged currently holds, since at least one cell per row reliably fires. **Verified: both rows in a rapid Tab-across-row batch computed correctly on the first attempt** (4hrs*$50=$200, 3hrs*$50=$150), revenue tracked correctly too ($350, no drops or double-counts).

**All fixes committed to git (5 commits: `93783bc` Ended Early + both first-attempt race fixes, `03e4874` Proposal_Tracker appendRow, `7049829` Hourly_Log decouple, `556c4df` SpreadsheetApp.flush) and pushed to `visualkirby/Upwork-Acquisition-Pipeline`. Each round also clasp-pushed to both `freelanceflow-template` (personal live pipeline) and `freelanceflow-production-master`.**

**Verification method note:** re-enabled `mcp__playwright__*` (moved deny -> allow in `.claude/settings.local.json`) multiple times this session for live browser-driven testing, reverting to `deny` each time per standing instruction. Each fresh test copy required a full Google OAuth re-authorization click-through (per-copy, since Apps Script authorization doesn't carry over from the source spreadsheet). Two throwaway test copies from this session ("FreelanceFlow - Bugfix Verification Test 2026-07-05" and "...Round 2") still exist in Drive with some temporary debug `console.log` lines added directly via the Apps Script web editor (never synced back to canonical source) -- candidates for manual deletion, no Drive-delete tool available to do it directly.

### What Is Next
- Two throwaway verification test copies in Drive still need manual deletion (see note above)
- Record a 10-minute Loom demo walkthrough (still open from 2026-07-04)
- Build the pre-launch / launch-day / post-launch plan for LinkedIn, r/freelance, r/upwork (append to `G:\My Drive\Benchline Analytics\Important_Plans\Benchline_Product_And_Job_Search_Plan_2026-06-30.md`)
- Cross-reference this and the 2026-07-04 session's contract/chat/revenue logic against the PipelineIQ SaaS app before that build starts (`project_freelanceflow_pipelineiq_crossref.md`)
- Confirm with Sawandi whether "Bake FILTER pull architecture into Library + template" (open item from 2026-07-04) is still open or was folded into this session's work
- Consider whether the same "Sheets sometimes never dispatches onEdit for a specific cell during rapid multi-cell entry" root cause affects any OTHER handleEdit blocks in `14_Edit_Trigger.gs` beyond Hourly_Log (Job_Discovery, Contract_Tracker, Proposal_Generator) -- not audited this session, only found via targeted testing of the two reported bugs

---

## Session: 2026-07-05 (cont.)

### What Was Done

**Menu reorder + sheet tab reorder + new Drop Keywords feature, per Sawandi's exact spec, shipped just before cutting the Gumroad production copy:**
- `System Tools` menu fully reordered (`01_Menu.gs`): FreelanceFlow Setup first, then Start/End Session, then logging items (Log New Job/Log Proposal Bid/Import Client Chat/Log New Contract), then analysis items (Run Job Classification/Run AI Proposals/Analyze Job Workflow/Analyze Session Patterns/Analyze Bid Patterns -- kept per Sawandi's confirmation, his list omitted it by accident), then keyword items (Generate Keyword Strategy/Mine Keywords/**Drop Keywords**, new), Snapshot Month End, API key items, diagnostics, and Reset System last
- New `reorderPipelineTabs_(ss)` in `00_Setup_Wizard.gs`, called at the end of `wizard_initialize` -- moves every tab into a fixed final order (Job_Discovery through Monthly_Performance per Sawandi's spec) independent of creation order, then deletes the leftover default "Sheet1". Lower-risk than reordering the actual sheet-creation/dependency sequence.
- New "Drop" dropdown column on `Keyword_Strategy` (`applyKeywordStrategyValidation_`) -- a manual user override, separate from the AI-computed `Recommended_Action` column. New `purgeDroppedKeywords_(ss)` + `DROP_KEYWORDS()` in `18_Keyword_Strategy.gs` remove matching rows from `Keyword_Search_List` only; the `Keyword_Strategy` row itself (and its Drop marker) is never touched, per Sawandi's confirmed choice.
- **Found and fixed a pre-existing dead-code bug along the way**: `MINE_KEYWORDS()` already had a "purge dropped keywords first" block, but it checked `Recommended_Action === "Drop"` -- a value that formula (`Lib_KeywordStrategy.gs`'s `buildKeywordStrategyFormula`) can never actually produce (only "Complete"/"Keep Testing"/"Avoid"). That purge had silently never fired since it was written. Replaced with a shared `purgeDroppedKeywords_()` call keyed off the new manual Drop column, so Mine Keywords' auto-purge-before-mining now actually works.

**All of it verified live before shipping** (fresh copy off the updated Production Master, via Playwright): full menu order confirmed, all 15 tabs confirmed in the exact requested order with Sheet1 deleted, Drop dropdown confirmed selectable, Drop Keywords confirmed removing exactly the right row count from Keyword_Search_List while leaving the Keyword_Strategy row untouched. Pushed via git + clasp to both `freelanceflow-template` and `freelanceflow-production-master` before the verification copy was even made.

**Created the actual launch artifacts:**
- **Gumroad production copy**, named plainly **"FreelanceFlow"** (Sheet ID `1Xv2Dt_nc55F8eOVULxIZbOAJsqca8upzZVzzFuzsDuQ`) -- untouched, wizard not run, this is the file customers get via "File > Make a copy"
- **Sawandi's own personal copy**, "FreelanceFlow - Sawandi's Upwork Pipeline" (Sheet ID `12CdpGyMx9KRE8InOxGrBqos3Y87eLTf2Kjmd3dXTInI`) -- untouched, Sawandi will run the Setup Wizard himself with his real info rather than Claude driving it
- **Loom demo copy**, "FreelanceFlow - Loom Demo Copy" (Sheet ID `12Yk8N6g44Z5zfeTBBSGDF5-_1thlkIz-wp6tN93pveY`) -- untouched, ready for Sawandi to record the demo walkthrough on

**Compiled a full list of every throwaway FreelanceFlow test/verification copy in Drive (10 files spanning this session back through 2026-07-03's E2E testing) via `search_files`, since no Drive-delete tool exists in this environment** -- reported all 10 with links for Sawandi to delete manually (Bugfix Verification Test x2, Menu-Order Verify, Walkthrough Test (Jordan Ellis) in 4 variants across sessions, Master Template v2, E2E Test Copy x2). Confirmed what to keep: the 3 new copies above, the dev "Production Master Template" (still needed for future clasp pushes), and "FreelanceFlow | Upwork Bidding Pipeline" (Sawandi's real, currently-in-use personal pipeline with live data -- not a test copy).

### What Is Next
- Sawandi to manually delete the 10 throwaway test copies listed above (links given, no tool available to do it directly)
- Sawandi to run the Setup Wizard on "FreelanceFlow - Sawandi's Upwork Pipeline" with his real info, replacing his old live pipeline (`FreelanceFlow | Upwork Bidding Pipeline`) as his working copy going forward
- Record the 10-minute Loom demo walkthrough using "FreelanceFlow - Loom Demo Copy" (long-open item, copy is now ready)
- Set up Gumroad listing pointing at the new "FreelanceFlow" production copy; update the 3 placeholder `https://gumroad.com` links in `page-freelanceflow.php` once live
- Build the pre-launch / launch-day / post-launch plan for LinkedIn, r/freelance, r/upwork (append to `G:\My Drive\Benchline Analytics\Important_Plans\Benchline_Product_And_Job_Search_Plan_2026-06-30.md`)
- Cross-reference this and the 2026-07-04 session's contract/chat/revenue logic against the PipelineIQ SaaS app before that build starts (`project_freelanceflow_pipelineiq_crossref.md`)
- Confirm with Sawandi whether "Bake FILTER pull architecture into Library + template" (open item from 2026-07-04) is still open or was folded into this session's work
- Consider whether the "Sheets sometimes never dispatches onEdit for a specific cell during rapid multi-cell entry" root cause (found in Hourly_Log) affects other `14_Edit_Trigger.gs` blocks -- not audited

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

---

## Session: 2026-07-04

### What Was Done

**Fixed the single-question Additional_Answers bug (two root causes).**
1. **Trigger race**: the handler was literally named `onEdit(e)`, which Apps Script auto-registers as a restricted "simple trigger" (can't call `UrlFetchApp`) *in addition to* the properly authorized installable trigger already registered via `registerEditTrigger_()`. Both fired on every edit; the simple trigger's permissions-error output could race the installable trigger's correct AI output and overwrite it. Fixed by renaming the handler to `handleEdit(e)` throughout `14_Edit_Trigger.gs` (no bare `onEdit` exists anymore, so Apps Script never auto-registers the restricted version) and re-pointing `registerEditTrigger_()` at the new name. Folded into `REPAIR_FORMULAS()` so existing customers get the fix automatically.
2. **AI hallucination**: `generateAdditionalAnswers` in `library/Lib_ProposalGenerator.gs` invented 5 extra Q&A pairs when given only 1 real question. Fixed by tightening the prompt: "Answer ONLY the exact question(s) listed below -- do not invent, add, or answer any question that is not explicitly listed... If only one question is listed, return only that one question and its answer." Shipped as Library v18.

**Verified Proposal_Templates populates correctly from niche** -- confirmed the template-selection logic correctly keys off the Setup Wizard's niche field before moving on.

**Full Proposal_Tracker build-out** (MTD-metrics accumulation): added `Discovery_ID` as Proposal_Tracker's new first column (it had never been retrofitted with this cross-sheet join key despite every other pipeline sheet having it -- sourced from `Proposal_Generator`'s row at Sent-time), plus Viewed/Interview/Hired columns feeding `Connects_Helper`'s `MTD_Replies`/`MTD_Interviews`/`MTD_Hires`.

**Replaced Followup_Tracker with a full contract-management build** (large feature request covering interview chat logging, milestone/contract tracking, and hourly logging):
- **Client_Chat_Log** (new sheet): triggered when `Proposal_Tracker.Interview` is set to "Y" -- opens a Chat Import sidebar (`19_Chat_Import.gs` + `ChatImportSidebar.html`, same look/pattern as the existing Setup Wizard sidebar) where the user pastes a raw copy-pasted Upwork message transcript; a new Library function `parseChatTranscript` (`library/Lib_ChatParser.gs`, gpt-4o-mini, temp 0.2, strict "preserve every message verbatim" prompt) parses it into one row per message, replacing any prior parse for that Discovery_ID (filter-out-then-rewrite).
- **Contract_Tracker** + **Milestone_Tracker** (new sheets): triggered when `Proposal_Tracker.Hired` is set to "Y" -- opens a Contract Setup sidebar (`20_Contract_Setup.gs` + `ContractSetupSidebar.html`) to choose Fixed or Hourly and set up milestones (Fixed) or an hourly rate (Hourly). Milestone_Tracker rows move Funded -> Delivered -> Released (each stamps its own date); Contract_Tracker's Status=Completed computes a `Total_Released` rollup (sum of Released milestones, or sum of Hourly_Log for hourly contracts); Status=Ended Early runs two sequential `ui.prompt()` dialogs (amount actually funded/received, then Y/N on whether it was released) and writes `Total_Released` accordingly.
- **Hourly_Log** (new sheet): per-entry `Hours_Logged` input; script computes `Amount = Hours_Logged * Contract_Tracker.Hourly_Rate` via a new `getContractHourlyRate_()` helper, feeding revenue by delta (`e.value`/`e.oldValue`, never a live re-read -- see bug note below).
- **Revenue-timing redesign** (explicit correction mid-session): `Proposal_Tracker.Revenue` no longer auto-adds to `Connects_Helper` when Hired flips to "Y" -- it's now purely informational. Revenue is recognized only at `Milestone_Tracker.Status="Released"`, at each `Hourly_Log` entry (progressive), and at `Contract_Tracker.Status="Ended Early"` (manual reconciliation amount). `Status="Completed"` computes `Total_Released` for display but deliberately does **not** push additional revenue, to avoid double-counting what Released milestones/Hourly_Log entries already contributed.
- New generic helper `incrementConnectsHelperMetric_(ss, metricName, amount)` in `03_Helpers.gs`, reused across every revenue-accumulation trigger point.
- Followup_Tracker and Followup_Templates fully removed: sheet defs, `initFollowupTemplates_`, all edit-trigger write code, and the Reset menu's clearByHeaders block for it.

**UX upgrades:**
- **Job_Discovery guided sidebar** (`21_Job_Discovery_Sidebar.gs` + `LogJobSidebar.html`): same setup-wizard-style popup pattern for logging a new job. Its `job_saveEntry(data)` had to inline every side effect (`cleanJobText_`, date/session stamps, `AI_Fit_Notes` generation, dupe-link check) that a direct cell paste would normally trigger via `handleEdit`, since script-driven `setValues()` does not fire `onEdit`.
- **First-session walkthrough popups**: new `showWalkthroughSeen_(key)` / `showWalkthroughOnce_(key, title, message)` helpers in `03_Helpers.gs` (one-time script-property flags, same `FF_`-prefix convention as `FF_SETUP_COMPLETE`). Wired in after the first Job_Discovery entry, after the first APPLY auto-proposal, after the first proposal Sent, and the first time Client_Chat_Log / Contract_Tracker sidebars open.

**Two library/formula bugs fixed along the way:**
- **Revenue delta double-counting from live re-reads**: two rapid edits to the same cell (e.g., delete then retype) queue two async trigger executions; a handler that re-reads the *current* cell value instead of using the event's own `e.value`/`e.oldValue` can see both executions land on the same final value paired against different old-values, doubling the recorded delta. Fixed everywhere delta-tracking happens (Hourly_Log built with this pattern from the start).
- **Header-lookup violation, self-caught**: first draft of `chat_parseTranscript` built new Client_Chat_Log rows as positional arrays assuming fixed column order, violating the standing header-based-mapping rule. Rewrote to resolve every column via `getCol_` before building each row.
- **REPAIR_FORMULAS never created new sheets**: existing customers had no path to receive a template update's brand-new sheets (Client_Chat_Log, Contract_Tracker, Milestone_Tracker, Hourly_Log) short of destructively re-running the whole Setup Wizard. Fixed by adding `ensurePipelineSheets_(ss, [])` as the first line of `REPAIR_FORMULAS()` (only creates missing sheets, never touches existing ones).

**E2E Phase 4 (Send Diagnostic Report) and Phase 5 (version-pin bug/repair cycle) both verified** on the E2E test copy (Sheet ID `1H9VaAlJUrW_o-S1wGI6_e_5NYeV-oapka4KZErXVFr0`). Library progressed v17 -> v18 (Additional_Answers prompt) -> v19 (Lib_ChatParser added) -> v20 (deliberate test bug: flipped `<=` to `>=` in Final_Decision's connects gate, confirmed zero effect on the E2E copy until its `appsscript.json` pin was manually bumped) -> v21 (real fix, confirmed restored via Repair Formulas; also confirmed as the final good state pinned in both the production template and the E2E clone).

**Git repo was badly stale** -- `apps-script/` (the git-tracked thin client) was missing files 18-21 and several new HTML sidebars entirely. Fully resynced from the live `freelanceflow-template/` source, committed (`2f3f416`, "Sync apps-script thin client with active template; add contract/chat/hourly tracking", 30 files changed) and pushed to `visualkirby/Upwork-Acquisition-Pipeline`.

**Master Template v2 created** (pristine, never had the Setup Wizard run on it): Sheet ID `1EiKLjfpM_TYT-cUDu2W9GODct1lb1sAK0bUVPnJwBS8`, script ID `1dqxJ9mir_UnRHEKUO2xGbpzxjVlvGUFjXr14prHcOuTdBVZk3eDERs-k`, pinned to Library v21. A separate duplicate (via File > Make a copy, which also duplicates the bound script) was created and renamed "FreelanceFlow - Walkthrough Test (Jordan Ellis)" (Sheet ID `1VvQlwOVW1MZmutLq3paNq11EsHkWlUlPGLyVz_W3YLY`) for the user's own manual walkthrough test -- also never had the wizard run on it, so it will auto-launch fresh.

**Pivoted from further Claude-led E2E testing to a manual, human-run test.** Built `FreelanceFlow_Operations_and_Walkthrough.docx` (python-docx, saved to `G:\My Drive\Benchline Analytics\Important_Plans\`) with:
- **Part A -- Operations SOP**: Gumroad order fulfillment via a view-only Drive link to the Master Template (customer does their own "File > Make a copy" -- no per-order manual work), a diagnostic-report triage checklist with a ready-to-paste Claude Code prompt template, the proven push-the-fix version-pin cycle as a repeatable checklist, and a `Support_Tickets` sheet column spec (Date, Customer, Product Tier, Issue Summary, Diagnostic Report, Status, Fix Version Shipped, Resolution Notes, Follow-up Date) for logging/handling tickets without Claude available.
- **Part B -- manual 10-job walkthrough script**: persona "Jordan Ellis" (bookkeeping niche, Aggressive scoring profile, Connect Balance 100), a 10-job overview table + full job-data card per job engineered against the real deterministic scoring formulas (confirmed by reading `Lib_JobDiscoveryFormulas.gs`, `Lib_JobScoringFormulas.gs`, `Lib_WizardFormulas.gs`) so outcomes land reliably: 4 jobs reach a contract (2 Fixed: one Completed, one Ended Early/partial payment; 2 Hourly: one Completed, one Ended Early/no payment), the rest a mix of Review/Ready in Discovery and APPLY/HOLD/SKIP in Scoring. Contract-stage steps spelled out for all 4 contract jobs. Ends with a blank 10-row Friction Log table for capturing UX issues found.
- Verified: 119 paragraphs, 15 tables, 36 checkbox lines, all headings present in correct order.

**New note for the future SaaS buildout**: all of this session's updated logic (contract management, chat parsing, revenue timing, walkthrough UX) needs to be cross-referenced against the PipelineIQ/FreelanceFlow SaaS web app before that build starts (see memory `project_freelanceflow_pipelineiq_crossref.md`) -- timed for after the E2E test and Master Template v2 are both fully done.

### What Is Next
- Waiting on the user to run the manual walkthrough themselves against "FreelanceFlow - Walkthrough Test (Jordan Ellis)" using the docx script, and log anything found in its Friction Log table
- Complete "Post-test: Build Master Template v2" -- the pristine master + validation duplicate exist and are correctly configured, but the true first-time-customer confirmation pass (running the Setup Wizard fresh, unassisted) was deferred to the user's manual test; re-scope this step once the user returns from that
- Record a 10-minute Loom demo walkthrough
- Build a pre-launch / launch-day / post-launch plan for LinkedIn, r/freelance, and r/upwork (to be appended to the master product plan doc, not this repo)
- Cross-reference this session's contract/chat/revenue logic against the PipelineIQ SaaS app before that build starts (per `project_freelanceflow_pipelineiq_crossref.md`)
- Loose end: leftover unused `library/Lib_Helpers.gs` already removed this session (done, no longer open)
