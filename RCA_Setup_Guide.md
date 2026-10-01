# RCA cycle task tracker — Microsoft 365 setup

Prepared for Jeffrey Cox. This is an offline deployment package, not a deployed Microsoft 365 solution. It contains a native Power BI project, Microsoft Lists provisioning script and task data, and Power Apps control source. The Power Apps source must be pasted into a canvas app; it is not an importable .msapp. Native Power BI Desktop, Power Apps Studio and your tenant were not available for end-to-end testing. Review the validation file and run the acceptance checks below before sharing.

## What is included

- `data/RCA_Records.json` and `.csv`: all 2026 cycle records from the tracker, including task descriptions and checklist items.
- `lists/Create-RCAList.ps1` and `columns.json`: create one Microsoft List, seed missing records and create parent-task views for all teams and each individual team. Rerunning preserves existing task progress.
- `powerbi/RCA.pbip`: overview and three team pages, Year/Cycle/Phase/Business Group/Team slicers, completed/remaining/total/percentage cards, charts and task table. Overview opens without an example filter selection.
- `powerapps/RCA_Controls.yaml`: team tabs, filters, editable task details, completion checkbox and the Task 3 checklist inside the task detail section.
- Corrected Excel tracker is supplied separately as a working reference.

## Task IDs and counting

The visible format is **Team.BusinessGroup.Phase.Task**. Example: `DI.FC.2.1` means Design & Insights, FCM, Phase 2, Task 1. The first number is always the phase. `TaskKey` adds year and cycle, for example `DI.FC.2.1.2026.CY2`, so recurring tasks remain unique. Checklist IDs add `.C01` through `.C13` before the year and cycle and link through `ParentKey`.

The data contains 1,568 parent task occurrences across four 2026 cycles and 364 checklist records, totaling 1,932 list rows. There are 392 parent tasks per cycle. Data & Reporting Phase 1 Task 3 has 13 checklist items for each of seven business groups in each cycle. Checklist items are stored in the same list and shown inside their parent task in the app; they are excluded from dashboard task totals.

Cycles run Jan 1–Mar 31, Apr 1–Jun 30, Jul 1–Sep 30 and Oct 1–Dec 31. A parent task counts as complete only when its own checkbox is checked and every required checklist item is checked. Unchecking a checklist item in the app reopens its parent. Existing 2026 data only is seeded; future years require new records with unique year/cycle TaskKeys and adjusted cycle dates, plus adding the year to the app dropdown.

## 1. Preview Power BI offline

1. Extract the entire ZIP into a local folder; keep its subfolders together.
2. Open `powerbi/RCA.pbip` in an up-to-date Power BI Desktop that supports Power BI projects and PBIR reports. Enable those formats in Desktop options if required by your version.
3. Refresh. The default `UseSharePoint=false` loads an embedded snapshot; no work connection is needed for this preview. Editing the CSV does not change the embedded snapshot.
4. Import `powerbi/RCA_Theme.json` through View > Themes > Browse for themes. Check all four report pages and the slicers.

## 2. Create the Microsoft List at work

Use your organization's approved PowerShell/PnP setup and an approved Entra application client ID for interactive PnP authentication. The script does not register an application, install software or change tenant permissions. A site owner or authorized colleague may need to run it.

From PowerShell 7, run:

```powershell
./lists/Create-RCAList.ps1 -SiteUrl 'https://YOURTENANT.sharepoint.com/sites/YOURSITE' -ClientId 'YOUR_APPROVED_ENTRA_APPLICATION_CLIENT_ID'
```

The default list name is `RCA_Tracker`; retain it for the supplied app formulas. Record the list GUID printed at the end. The script adds missing keys and does not overwrite already seeded task progress. Version history is enabled. AssignedTo is a text field, not a directory Person field. Standard task views hide checklist rows; the app provides the parent/child editing experience. Direct list edits do not enforce the app's checklist guard, although Power BI still excludes incomplete checklists from completion totals.

## 3. Connect the report to the list

In Power BI Desktop, use Transform data > Edit parameters:

- `UseSharePoint`: true
- `SiteUrl`: your SharePoint site URL
- `ListId`: the GUID printed by the script (without braces)

Sign in with your work organizational account and refresh. The supplied query uses SharePoint internal field names. If your connector returns display names instead, map them to the names in `lists/columns.json` before the SelectColumns step; do not silently accept blank required columns. Verify 1,932 rows initially and 1,568 parent tasks before publishing. Publish to an authorized workspace and configure data-source credentials and scheduled refresh there. Power BI uses imported data: saving in the app updates the list immediately, but report totals change after a semantic-model refresh. Do not rely on PowerBIIntegration.Refresh for this SharePoint import model.

## 4. Create the editable Power Apps screen

1. In Power Apps Studio create a blank canvas app with a tablet/custom 1366 × 900 layout. Connect the SharePoint list `RCA_Tracker` as a data source.
2. Copy the complete contents of `powerapps/RCA_Controls.yaml`. Paste the control source into the blank screen using Studio's supported code paste feature. This is control YAML, not an entire app import. Studio may require control-version/property adjustments; resolve any App Checker errors before publishing.
3. Set the screen's `OnVisible` formula to the contents of `powerapps/Screen.OnVisible.fx`.
4. Play the app. All teams and all filters start selected as “All.” Click a team tab or choose Year, Cycle, Phase and Business Group, then select a task. Change fields, check Task complete and choose Save task. Checklist checkboxes save their individual changes immediately.
5. For Data & Reporting Phase 1 Task 3, the checklist appears below its parent details. All 13 items must be complete before saving the parent as complete. No separate checklist tab exists.
6. Resolve delegation warnings against SharePoint before broad use. Parent browsing uses Filter/SortByColumns; only the selected task's 13 checklist rows are collected locally. Do not replace the parent query with an unbounded local collection.
7. Save, publish and share the app with the team. SharePoint list permissions and Power Apps permissions must both be granted. Concurrent users should refresh before editing the same task; this starter uses last-saved field updates, not a custom conflict resolution workflow.

## 5. Put editing beside the report

For embedded editing, add a Power Apps visual to the published report, pass the RCA `TaskKey` field and use its Create new app workflow so the PowerBIIntegration object is present. Add the list connection and paste the same controls into that generated app as in step 4. Alternatively configure an existing compatible app through the visual. Use a dedicated report page or resize the task table to make room for the editor.

Add a button labeled “Load selected report task” to the app and set its `OnSelect` to `powerapps/LoadSelectedReportTask.OnSelect.fx`. Select exactly one parent task in the report table, then click that button to load it. Configure visual interactions so the task table filters the Power Apps visual. Pass a single selected TaskKey to avoid the visual's 1,000-row transfer limit; the app retrieves its checklist directly from the list. Publish and share the embedded app separately from the report. This final embedding step needs your work app/report IDs and is not pre-bound in the offline report.

The full app also works standalone with its own filters. Use Power BI's slicers for reporting and the app's controls for list editing; changing the app's dropdowns does not change Power BI slicers automatically.

## 6. SharePoint and acceptance checks

Embed the published report using SharePoint's Power BI web part and, if desired, embed the standalone app using the Power Apps web part. Use your tenant's private sharing and licensing arrangements. A SharePoint page alone does not grant report, app or list access.

Before sharing:

- Confirm the overview defaults to all data, and each team page restricts that team.
- Filter Year 2026, Cycle 2, Design & Insights and FCM; inspect IDs/descriptions and reset filters afterward.
- Verify the ID suffix numbers mean Phase then Task, using a case where they differ.
- Save a normal task as complete; reload the app and confirm it persisted. Refresh Power BI and verify completed rises by one and remaining falls by one.
- On Task 3, leave one check incomplete and verify the app prevents parent completion. Complete all 13, save the parent, refresh the report and confirm it contributes one task only. Uncheck one child and verify the parent reopens.
- Check all date fields, comments, team tabs, task table scrolling and access with a team member's account.
- Verify no duplicate TaskKeys after rerunning the list script, and that existing progress was preserved.

## Microsoft references

- Power BI projects: https://learn.microsoft.com/power-bi/developer/projects/projects-overview
- Report project format: https://learn.microsoft.com/power-bi/developer/projects/projects-report
- Power Apps control code: https://learn.microsoft.com/power-apps/maker/canvas-apps/code-view
- Power Apps visual and refresh limitations: https://learn.microsoft.com/power-apps/maker/canvas-apps/powerapps-custom-visual
- SharePoint delegation: https://learn.microsoft.com/power-apps/maker/canvas-apps/connections/connection-sharepoint-online
- PnP interactive connection: https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html
