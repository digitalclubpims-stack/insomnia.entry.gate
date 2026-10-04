# INSOMNIA Entry Portal — Phone-Only Google Sheets + Apps Script

This version intentionally removes Firebase, Cloud Functions, Firebase Auth, GitHub Pages dependencies, npm, CLI deployment, and separate frontend folders.

## What is included
- `apps-script/Code.gs` — complete backend and Google Sheet database logic.
- `apps-script/Index.html` — single responsive website containing Home, Get Pass, CR Verification, Gate Scanner and Admin.
- `SETUP.md` — phone-only setup instructions.

## New architecture
Google Form → Google Sheet → Apps Script Web App → CR/Admin/Student/Scanner views.

The Google Sheet is the live source of truth. Payment verification creates the ticket. The QR token is random and is not the visible ticket ID. Entry confirmation uses `LockService` so two scanners cannot normally consume the same night twice simultaneously.

## Removed as unnecessary
Firebase config, Firestore rules/indexes, Cloud Functions, Firebase Auth, GitHub Pages frontend modules, `src/firebase.js`, Firebase deployment files and separate admin/CR/scanner/ticket HTML folders.

## Important limitation
This is intentionally simpler and phone-manageable, but Google Sheets + Apps Script is less robust than Firestore under extreme concurrent traffic. For the planned ~1,000 attendees, 4 scanners and online-only entry, it is a practical lightweight solution if the venue internet is good.
