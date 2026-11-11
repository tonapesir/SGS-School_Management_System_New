SGS School Management System — V19.64.1 (फक्त क्लायंट दुरुस्ती; Apps Script पुन्हा Deploy करण्याची गरज नाही)

दुरुस्ती: डॅशबोर्डवर "मदत" कार्ड चुकून "डेटा तपासणी" कार्डाच्या आत घुसले होते (V19.64 मधील माझी चूक — बिल्ड-टूलचा anchor चुकीचा होता).
परिणाम: कार्डे एकमेकांत अडकली; (index_1.html मधील हाताने केलेल्या दुरुस्तीनंतर) <div> संतुलन बिघडून आजची हजेरी/वाढदिवस/सूचना कार्डे आयकॉनच्या ग्रिडमध्ये घुसली.

अपलोड (GitHub Pages): github_pages/index.html, script.js, sw.js — फक्त हे ३. (style.css, privacy-policy.html, help/ बदललेले नाहीत; Code.gs V19.64 तसाच.)
नंतर ॲप Refresh करा (PWA नवी आवृत्ती घेईल: APP_VERSION = V19.64.1, sw = sgs-v19.64.1-1).

नवी चाचणी: tests/test_html_structure.js — <div>/टॅग संतुलन, सर्व .page एकाच पालकाखाली, .dcard एकमेकांत नाहीत, widgets ग्रिडमध्ये नाहीत.
  node tests/test_html_structure.js          (SGS_INDEX=<file> ने दुसरी फाइल तपासता येते)
tools/build_help.js: मदत-कार्ड आता संतुलित </div> नंतर घालतो.
