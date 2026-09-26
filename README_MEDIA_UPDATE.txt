RAJ Agencies Feedback — Optional Photo / Video Update

WHAT STAYS SAME
- Existing 4-step UI, branding, background, ratings, sounds, animations and success screen.
- Existing custom domain: feedback.rajagencies.info
- Existing GitHub Pages setup.

WHAT IS NEW
- Step 4 has optional Photo / Video Testimonial.
- Mobile can use camera/gallery through the device file picker.
- One photo OR one video per feedback.
- Photo max 5 MB; video max 15 MB (30–45 sec recommended).
- Media permission checkbox appears only when a file is selected.
- Media is stored privately in Google Drive.
- Sheet gets: Media Type, Media Drive Link, Media File ID, Media File Name, Media Consent.

IMPORTANT BACKEND UPDATE
1. Open the Apps Script project currently serving the feedback API.
2. Back up the current Code.gs first.
3. Replace Code.gs with the Code.gs included in this ZIP.
4. Project Settings > Script Properties: confirm SPREADSHEET_ID already contains your testimonial spreadsheet ID.
5. Save, then Deploy > Manage deployments > Edit > New version > Deploy.
6. Execute as: Me. Access: Anyone.
7. Run healthCheck() once from the editor and authorize Drive access when Google asks. It will create/find:
   RAJ Agencies - Customer Testimonial Media
8. Keep the existing Web App URL if editing the existing deployment. index.html already points to the current API URL.

GITHUB UPDATE
- Replace index.html in PxPixaro/raj-feedback with the index.html in this ZIP.
- Keep assets/RAJ_GROUP_Logo.png and CNAME as included.
- Do NOT change GoDaddy DNS.

TEST
1. Submit normal feedback without media — it should work exactly as before.
2. Submit one small photo — verify new Sheet row and Drive link.
3. Submit one short video under 15 MB — verify Drive file and Sheet link.

NOTE
GitHub Pages cannot store uploads. The browser sends the optional media to Apps Script, which saves it in your Google Drive. Large mobile videos are deliberately blocked to keep uploads reliable.
