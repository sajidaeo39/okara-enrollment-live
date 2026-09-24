# Okara Live Enrollment Dashboard

This repository turns the existing Okara SIS enrollment Python report into a daily GitHub-hosted dashboard.

## What happens automatically

- Every day at **07:00 Pakistan Standard Time (UTC+5)**, GitHub Actions runs `update_enrollment.py`.
- The script connects to SIS, gets all three Okara tehsils, all markazes and all schools.
- It writes `data/Okara_Latest_Enrollment_Gender_Latest.xlsx`.
- A dated Excel archive is also stored in `data/`.
- The dashboard `index.html` automatically loads the latest workbook; no manual upload is required.
- Anyone with the GitHub Pages link can view the live dashboard and download the latest Excel.
- The dashboard still supports the existing School Wise / Markaz Wise / Tehsil Wise / Wing Wise / Top 10 / Last 10 / Low Performing views and filtered Excel/PDF downloads.

## Important completeness rule

The Python job compares the discovered school list with the enrollment result rows and also checks for failed school requests. If any school request fails, the run stops before replacing the published `Latest` workbook. This prevents a partial SIS result from being presented as a complete daily report.

The existing script excludes the same explicitly excluded school(s) as the supplied version. It does not silently add other exclusions.

## One-time GitHub setup

1. Create a GitHub repository, for example `okara-live-enrollment`.
2. Upload the contents of this folder to the repository's `main` branch.
3. In GitHub: **Settings → Pages → Source: GitHub Actions**.
4. Open the Pages URL shown by GitHub.

After that, normally you only open the Pages link. The scheduled workflow keeps the data updated.

## Manual refresh

GitHub Actions also has **Run workflow** for an on-demand update. This is useful if SIS was temporarily unavailable at 07:00.

## Note about the 07:00 time

The schedule is `02:00 UTC`, which corresponds to `07:00 PKT`. GitHub Actions can start a scheduled workflow a little after the exact minute because GitHub schedules are not guaranteed to execute at the exact second.
