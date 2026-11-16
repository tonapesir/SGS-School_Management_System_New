SGS School Management System — V19.65 (V19.64.1 + शैक्षणिक वर्ष बॅज, Students column B जुळवणी, बढतीपूर्वी Backup-पडताळणी)

Deploy क्रम:
1) Apps Script: apps_script/Code.gs paste → Deploy → Manage deployments → ✏️ → "New version" → Deploy. (SERVER_VERSION = V19.65)
   (सर्व्हर आधी, मग GitHub — ping मधील वर्ष जुन्या सर्व्हरवर नसेल तर बॅज "—" दिसेल, त्रुटी नाही.)
2) GitHub Pages: index.html, script.js, style.css, sw.js, privacy-policy.html अपलोड. (help/ बदललेला नाही.) PDF पुन्हा वाटा (नवी ओळी जोडल्या).
3) ॲप Refresh करा.

नवे:
 • डॅशबोर्ड (व वर्ग शिक्षक पान) च्या उजव्या कोपऱ्यात ठळक "📅 शैक्षणिक वर्ष: 2026-27" बॅज. Super/Master साठी तो दाबल्यावर पॅनेल; वेगळे वर्ष असल्यास ⚠️ लाल.
 • Students चा column B: Save करताना आपोआप चालू वर्ष; जुने/वेगळे वर्ष Super "जुळवा" (SYNC) ने — आधी Students चा Backup (पडताळणीसह), फक्त column B बदलतो, बदलांच्या इतिहासात नोंद.
 • बढती: नवीन वर्षाची पडताळणी (स्वरूप/क्रम), आणि बढतीपूर्वी Students चा Backup खरोखर पूर्ण आहे का ते तपासूनच बढती; Backup चे नाव/ठिकाण संदेशात. (बढती मिळालेल्यांचा column B = पुढचे वर्ष — आधीपासून होते, आता चाचणीने सिद्ध.)
 • वर्ष-अखेर पायरी ९: जुने वर्ष असलेले विद्यार्थी असल्यास "📅 वर्ष जुळवा…".

चाचण्या: tests/ मध्ये टाका; नवीन: test_server_v1965.js, test_ui_v1965.js (आणि V19.64.1 पासून test_html_structure.js).
सर्व बनावट (mock) Apps Script / बनावट backend वर — खऱ्या Sheet/Drive/Apps Script/फोनवर चाचणी झालेली नाही.
