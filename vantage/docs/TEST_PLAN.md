# Vantage 1.0: tenant test plan

Run this on a **non-production** site first. Use three test accounts:
- **Owner:** a site owner.
- **Member:** added as a workspace Contributor.
- **Reader:** added as a workspace Reader.

## 1. Package and deploy
- [ ] Upload `releases/vantage.sppkg` to the App Catalog. It shows as **Vantage** with its icon and screenshots.
- [ ] Choose "Enable this app and add it to all sites", then add the **Vantage** web part (Advanced group) to a page.

## 2. Setup
- [ ] As Owner, select **Set up Vantage**. Seven `VNT_*` lists are created and their internal names match ARCHITECTURE.md (check under List settings → column URL `Field=`).
- [ ] Reload the page: it opens straight into the app.
- [ ] Run setup again (the Repair or Upgrade button, or temporarily delete one control list): it completes without duplicating anything.
- [ ] *(0.2.0 sites only)* After the upgrade the page shows **Upgrade now**. Selecting it keeps existing workspaces and archive history.

## 3. Workspaces and permissions
- [ ] Create a workspace. It shows as **Secured**; the three `VNT_<key>_*` groups exist and the Owner is in Admins.
- [ ] Both workspace lists have unique permissions: site Owners and Admins have Full Control, Members Contribute, Readers Read.
- [ ] Add Member and Reader through **Access**, leaving the site-access option ticked.
- [ ] Member can create and edit tasks, comments and checklists. Member doesn't see the Archive or Access tabs.
- [ ] Reader sees tasks read-only.
- [ ] A user in neither group can't open the workspace lists by URL (access denied).
- [ ] Workspace B's users can't see workspace A in the switcher.
- [ ] Removing a member takes effect after a page reload.

## 4. Work management
- [ ] Create boards; a duplicate board name is rejected.
- [ ] Create a task with an assignee (people picker), a due date and a priority. The assignee's bell shows "You were assigned…".
- [ ] Editing the status writes an activity entry. Completed sets the Completed Date and 100% progress; reopening clears it.
- [ ] Comments, checklist (tick and untick, remove) and dependencies (no self or duplicate dependency) all work.
- [ ] Pressing Escape with unsaved changes asks before discarding them.
- [ ] Deleting a task removes its comments, checklist items, activity and notifications, plus dependency rows on other tasks that point at it.
- [ ] Filters (text, status, assigned to me) work. Overdue dates show in red. At phone width the list shows as cards.
- [ ] The Overview figures match the data.

## 5. Archive
- [ ] Create about 50 tasks: some Completed or Cancelled with old Completed Dates, some open. Add comments and checklists, attach small files to a few, and attach one file larger than 25 MB.
- [ ] Save a **date** policy and run it. Eligible tasks and all their related records move to the archive. Attachments are present in the archive items. Open tasks stay in the active list.
- [ ] The task with the large attachment is **Skipped**, and nothing for it appears in the archive list.
- [ ] Job history and the audit log show the run. `VNT_ArchiveLocks` is empty afterwards.
- [ ] Run a **count** policy with a low maximum. The oldest eligible tasks are archived; if not enough are eligible, the job reports how many records remain over the limit.
- [ ] Start a run in two browsers at once: the second one is refused.
- [ ] Close the browser during a run. After 30 minutes a new run starts and the old job shows **Stopped**. A re-run doesn't create duplicate archive copies.
- [ ] Edit a task (in another tab) while it is being archived: it is kept, and the failure is logged.

## 6. Restore
- [ ] Restore an archived task. It returns with a new ID, its comments, checklist, activity and attachments, and the archive copies are removed.
- [ ] Restoring while an archive run is in progress is refused.

## 7. Scale and resilience
- [ ] Use a workspace with more than 5,000 records in the active list. Boards, the dashboard and an archive run keep working (indexed filters).
- [ ] When throttled (429), requests retry with back-off and don't create duplicate items.
- [ ] Check the dark theme and section background variants, keyboard-only navigation, and a screen reader on the task panel.
