SGS School Management System — V19.61 (अ-२ ते अ-६)

Deploy क्रम:
1) Apps Script: apps_script/Code.gs paste करा → Deploy → Manage deployments → ✏️ Edit → Version: "New version" → Deploy.
   (SERVER_VERSION = V19.61.) एकदा Editor मधून setupClientErrorPurgeTrigger चालवा (दर रविवारी ९० दिवसांपेक्षा जुन्या त्रुटी पुसतो). पर्यायी: auditGrowthReport (केवळ वाचन).
2) GitHub Pages: github_pages/ मधील index.html, script.js, style.css, sw.js अपलोड करा (APP_VERSION = V19.61, sw = sgs-v19.61-1).
   manifest.json, privacy-policy.html, icons बदललेले नाहीत.

चाचण्या (तुमच्या tests/ फोल्डरमध्ये नवीन फाइल्स टाका; run_all.sh test_*.js आपोआप घेत असल्यास तसेच चालेल):
  node tests/test_static_refs.js          (अ-२)
  node tests/test_ui_contrast_all.js      (अ-३; Playwright + Chromium)
  node tests/test_server_v1961.js         (अ-४/अ-६ सर्व्हर; tests/lib/gas_mock.js लागते — बनावट Apps Script)
  node tests/test_ui_v1961.js             (अ-४/अ-५ क्लायंट; Playwright)
env: SGS_WEB = index.html चा फोल्डर, SGS_GS = Code.gs चा path.
सर्व बनावट (mock) वातावरणावर — खऱ्या Sheet/Drive/Apps Script/फोनवर चाचणी झालेली नाही.
