RAJ Agencies V7 — GitHub Pages + Apps Script backend

1) Apps Script: replace Code.gs with this Code.gs. Keep your existing Spreadsheet Script Property. Save.
2) Deploy > Manage deployments > Edit > New version > Deploy. Keep Execute as Me / Anyone.
3) GitHub raj-feedback: replace index.html, add CNAME, and upload assets/RAJ_GROUP_Logo.png.
4) Settings > Pages: main / root. Custom domain feedback.rajagencies.info.
5) Test one submission and confirm a new Testimonials row. Admin Notes stores the browser reference.

IMPORTANT: V7 uses a no-CORS background POST so customers do not navigate to Apps Script or hit Google account/session UI. The success screen uses a browser reference (RAJ-WEB-...), while the Sheet keeps the official RAJ-T-###### ID.
