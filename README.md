SPO Reading View makes Word and Excel documents stored in SharePoint Online and OneDrive for work or school open in Reading View by default, instead of Editing View.

Why?
Opening a document in Editing View can cause accidental edits, co-authoring presence and unnecessary "modified" activity. Most of the time you just want to read.

How it works
- Document links: the link is rewritten locally so the document opens in Reading View.
- Sharing links (for example the short "/:w:/s/..." or "/:x:/s/..." links you receive in chat or email): the extension asks the SharePoint site you are visiting which document the link points to, then opens that document in Reading View.
- Works when you click links in web pages, paste them into the address bar, or open them from other apps.
- You can switch to editing at any time with the Edit button inside Word or Excel for the web; the extension will not switch you back.

Features
- Works on *.sharepoint.com and *.sharepoint.cn
- Word and Excel documents
- Toolbar popup with simple on/off switches
- Optional: also override links that explicitly request edit mode (off by default)
- Optional debug log for troubleshooting (does not log site or file names)

Privacy
- No data collection, no analytics, no ads, no remote code.
- The only network request the extension makes is a document lookup sent to the same SharePoint site you are already visiting, using your existing session. Nothing is sent to the developer or any third party.
- Preferences are stored locally on your device.

Limitations
- PowerPoint for the web ignores the Reading View setting, so PowerPoint files are not supported.

This is an independent open-source project and is not affiliated with, endorsed by, or sponsored by Microsoft. SharePoint, OneDrive, Word and Excel are trademarks of the Microsoft group of companies.
