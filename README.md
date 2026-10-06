# Student Search Website

This project creates a simple public website that searches student records from an Excel spreadsheet.

## Features
- loads the spreadsheet from GitHub
- searches by student name
- displays matching rows in a table
- ready to deploy on Vercel

## Setup

```bash
npm install
npm run dev
```

Then open http://localhost:3000

## Spreadsheet URL
The app currently loads this workbook:

https://raw.githubusercontent.com/kohyupingjanice-stack/Test-1/main/2026%20TG%20allocation_Sem%201_9%20Jan%20(2).xlsx

If the file moves or is renamed, update `SPREADSHEET_URL` in `app/page.js`.

## Deployment
Deploy this app to Vercel and it will get HTTPS automatically.

If you want a custom `.io` domain, connect it in the Vercel dashboard after deployment.
