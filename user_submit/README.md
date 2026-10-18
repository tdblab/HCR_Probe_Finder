# Submitted probe sets

Drop submission forms (filled copies of `templates/HCR_probe_submission.xlsx`) into this folder. Every push rebuilds
the site, and each form's genes appear on the **Lab probe sets** tab under the submitter's name.

To add one from the GitHub website: open this folder, click **Add file → Upload files**, drop the .xlsx, and click
**Commit changes**. The site updates about two minutes later (watch the Actions tab for the green tick).

- Name files `Firstname_Lastname_YYYY-MM-DD.xlsx` so they are easy to find.
- To replace a set, upload the corrected file and delete the old one.
- Files whose names start with `_` are ignored, which is a quick way to hold one back without deleting it.
- Pairs that fail the checks (wrong letters, unknown or mixed initiators, missing halves) are left out. The Actions log
  (open the run, then the **build** step) lists every file read and any rows it skipped.
- Set IDs come from the initials on the form (for example JT-001). Each person's files are numbered in date order. If
  two people choose the same initials, the second gets a number added (JT2-001).
