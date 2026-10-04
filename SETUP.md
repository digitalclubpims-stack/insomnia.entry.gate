# INSOMNIA — phone-only setup

## 1. Create the Google Sheet
Create one blank Google Sheet. Open **Extensions → Apps Script**.

## 2. Add the two files
Create/replace `Code.gs` with the contents of `apps-script/Code.gs`.
Create an HTML file named `Index` and paste `apps-script/Index.html` into it.

## 3. Run setup once
In Apps Script, select `setupSystem` and press Run. Approve the Google permissions. This creates:
- Registrations
- Tickets
- EntryLogs
- Users
- Config

The first admin password is `CHANGE-ME-1234`. Change it immediately in the Config sheet.

## 4. Connect the Google Form
Keep your existing Google Form and response sheet. In Apps Script, create an installable trigger:
- Function: `onFormSubmit`
- Event source: From spreadsheet
- Event type: On form submit

The script expects these form headings (minor spelling variations are supported): Email, Name of attendee, Roll number, Phone number, Batch, Money paid to, UTR number, Payment screenshot, Amount.

## 5. Deploy the website
Apps Script → Deploy → New deployment → Web app.
- Execute as: **Me**
- Who has access: **Anyone**

Copy the Web App URL. This is the single INSOMNIA website.

## 6. CR access
The seeded CR usernames are:
- CR-2021-1 — Kashish Mahajan
- CR-2022-1 — Rhythm Gupta
- CR-2023-1 — Gurman Singh Bhatia
- CR-2024-1 — Nishant Mittal
- CR-2025-1 — Ishan

Their access code is initially blank. On first login, the first code entered becomes their code. Give each CR a private code. CR 2 slots and 2026 can be added from Admin.

## 7. Scanner access
The four seeded scanner usernames are:
- BOYS-01
- BOYS-02
- GIRLS-01
- GIRLS-02

Set an access code for each. Their first login sets their code if blank. Each scanner is gate-limited.

## 8. Student flow
Student pays CR → submits Form → CR verifies UTR → ticket is automatically generated → student opens Get Pass and enters mobile + UTR → QR appears.

## 9. Gate flow
Scanner signs in → chooses the current night → scans QR → scanner displays ticket details → volunteer checks college ID → volunteer presses **CONFIRM ENTRY** → that night's status becomes USED.

The same ticket can be used once on each of Nights 1, 2 and 3.

## 10. Night control
Admin can set the current night from the Admin page. The scanner also requires a night selection so the operator cannot accidentally consume the wrong night.

## 11. Do not delete these columns
Do not rename or delete columns in the five system sheets after setup. The script uses fixed column positions.
