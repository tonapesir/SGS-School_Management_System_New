SGS School Management System — V19.63 (V19.62 + ब-९ वर्ष-अखेर मार्गदर्शक + ब-१० रोजचा सारांश)

Deploy क्रम:
1) Apps Script: apps_script/Code.gs paste → Deploy → Manage deployments → ✏️ → "New version" → Deploy. (SERVER_VERSION = V19.63)
   नवे Editor-फंक्शन Run करण्याची गरज नाही (V19.62 चे setup* आधी केले असतील). नवे Trigger नाही.
2) GitHub Pages: github_pages/ मधील index.html, script.js, style.css, sw.js अपलोड (APP_VERSION = V19.63, sw = sgs-v19.63-1).

नवे: डेटा तपासणी पानात "🎓 वर्ष-अखेर मार्गदर्शक" टॅब (Super); डॅशबोर्डवर "📋 रोजचा सारांश" कार्ड (Super + Master). ईमेल पर्याय नाही (निर्णय).

चाचण्या (tests/ मध्ये टाका):
  node tests/test_static_refs.js
  node tests/test_ui_contrast_all.js
  node tests/test_server_v1961.js
  SGS_GS_BASE=<V19.61 Code.gs> node tests/test_server_v1962.js
  SGS_GS_BASE=<V19.62 Code.gs> node tests/test_server_v1963.js   (BASE ऐच्छिक: "विद्यमान कृती बदलल्या नाहीत" ही तुलना-चाचणी)
  node tests/test_ui_v1961.js ; node tests/test_ui_v1962.js ; node tests/test_ui_v1963.js
सर्व बनावट (mock) Apps Script / बनावट backend वर — खऱ्या Sheet/Drive/Apps Script/फोनवर चाचणी झालेली नाही.
