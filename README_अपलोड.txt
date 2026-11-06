SGS School Management System — V19.64 (V19.63 + ब-११ गोपनीयता/डेटा-धारण + ब-१२ वापरकर्ता-मार्गदर्शिका)

Deploy क्रम:
1) Apps Script: apps_script/Code.gs paste → Deploy → Manage deployments → ✏️ → "New version" → Deploy. (SERVER_VERSION = V19.64)
   नवे Editor-फंक्शन: retentionReport() — केवळ वाचन (काहीही पुसत/लिहित नाही). Trigger नाही.
2) GitHub Pages: github_pages/ मधील index.html, script.js, style.css, sw.js, privacy-policy.html आणि नवा फोल्डर help/ (५ चित्रे) अपलोड.
3) privacy-policy.html हा मसुदा आहे — मुख्याध्यापक/कायदेशीर सल्लागाराच्या मंजुरीनंतरच प्रकाशित/App Store लिंक म्हणून वापरा.
4) SGS_वापरकर्ता_मार्गदर्शिका.pdf — शिक्षकांना वाटण्यासाठी (A4, मराठी फॉन्ट embed).

मदत-मजकूर बदलायचा असल्यास: tools/help_content.html संपादित करा →
   node tools/build_help.js                                     (index.html मधील ❓ मदत पान अद्ययावत)
   node tools/build_help.js --pdf out.pdf --font-dir tools/fonts (PDF)
   node tools/make_help_shots.js                                (बनावट डेटासह स्क्रीनशॉट; ॲप बदलल्यास)
(Node + Playwright + Chromium लागते.)

चाचण्या (tests/ मध्ये टाका):
  node tests/test_static_refs.js ; node tests/test_ui_contrast_all.js ; node tests/test_privacy_policy.js
  node tests/test_server_v1961.js ; SGS_GS_BASE=<V19.61 Code.gs> node tests/test_server_v1962.js ; SGS_GS_BASE=<V19.62 Code.gs> node tests/test_server_v1963.js ; node tests/test_server_v1964.js
  node tests/test_ui_v1961.js ; node tests/test_ui_v1962.js ; node tests/test_ui_v1963.js ; SGS_PDF=<guide.pdf> node tests/test_ui_v1964.js
सर्व बनावट (mock) Apps Script / बनावट backend वर — खऱ्या Sheet/Drive/Apps Script/फोनवर चाचणी झालेली नाही.
