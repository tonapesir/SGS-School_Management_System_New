SGS School Management System — V19.57

१) Apps Script (आधी): Spreadsheet ची प्रत घ्या → apps_script/Code.gs paste करा → Deploy > Manage deployments > Edit > New version.
   नंतर Editor मध्ये एकदा diagRosterCacheSize चालवा (View > Logs) — विद्यार्थी-यादी Cache मध्ये बसते का ते कळेल.
२) GitHub: github_pages मधील index.html, script.js, style.css, sw.js replace करा (बाकी फाइल्स न बदललेल्या).
   sw.js मधील cache नाव sgs-v19.57-1 आहे; जुना cache आपोआप जाईल.
३) tests/ फोल्डर: चाचणी-साधन (mock + Playwright). Repository मध्ये ठेवू शकता; खऱ्या Sheet वर चालत नाही.
