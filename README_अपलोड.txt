SGS School Management System — V19.56 (अंतिम)

१) Apps Script (Code.gs)  — आधी हे करा
   - Spreadsheet ची प्रत (Make a copy) घ्या.
   - Extensions > Apps Script मध्ये जुना Code.gs बदलून apps_script/Code.gs चा मजकूर paste करा.
   - Deploy > Manage deployments > Edit (पेन्सिल) > Version: New version > Deploy.
   - Web app URL तोच राहतो.

२) GitHub (फक्त github_pages फोल्डरमधील फाइल्स)
   - Repository मध्ये या फाइल्स जुन्या फाइल्सवर replace करा:
     index.html, script.js, style.css   (हे तीन बदललेले आहेत)
   - sw.js, manifest.json, privacy-policy.html, icon-192.png, icon-512.png बदललेले नाहीत
     (सोयीसाठी दिल्या आहेत; आधीच्या तशाच ठेवल्या तरी चालते).
   - Commit केल्यावर GitHub Pages ~१-२ मिनिटांत अपडेट होते.
   - अॅपमध्ये Hard refresh (Ctrl+Shift+R) / PWA बंद करून पुन्हा उघडा. sw.js नेटवर्क-आधी असल्याने नवी आवृत्ती आपोआप येते.

३) तपासणी: Footer/आवृत्ती V19.56 दिसते का; Profile मध्ये "फी व प्रवास" व "सर्व वर्षांचा इतिहास"; Super/Master ला मेनूत "📦 Archive".
