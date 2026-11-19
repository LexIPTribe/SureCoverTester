# SureCover Tester

A Windows demo app that generates fictional flight delay claim documents and prepares an email in classic Outlook for review.

## Requirements

- Windows with Windows PowerShell 5.1
- Microsoft Outlook classic configured on the computer

## Run

Double-click `Start SureCover Tester.bat`. The app opens a form with randomized fictional test data. Choose a scenario, edit the fields as needed, and select **Generate + email** to create sample documents and open a prefilled message in Outlook.

Tick **Add frustration and urgency to email** to add a frustrated, please-hurry paragraph to the email body. SureCover then flags the claim as **Urgent**.

The app saves generated files under `Documents\SureCover tests`. Use **Generate files only** if you do not want to open an email. The **Send immediately (skip review)** option is off by default; leave it off to review the message in Outlook before sending.

All generated documents are marked as fictional demonstration data. Do not enter real claimant or policy information into public demos.
