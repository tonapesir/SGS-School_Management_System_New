SGS School Management System — V19.60

नवे (सर्व वाचन-आधारित; एकमेव नवे लिखाण = Student_Audit Sheet, Save कधीही अडवत नाही)
1. 🔎 डेटा तपासणी (Super): दुबार Reg No, जन्मतारीख रिकामी/चुकीची, WhatsApp रिकामा/१० अंकी नाही, वय-वर्ग न जुळणे, फोटो नाही. वर्गनिहाय + यादी + CSV + Profile.
2. 🕑 बदलांचा इतिहास (Profile, Super/Master): कोणी, कोणता फील्ड, कधी, जुने->नवे. आधार फक्त शेवटचे ४ अंक. + Activity Log + बढती.
   फील्ड-स्तरीय इतिहास V19.60 पासून पुढचे बदलच नोंदवतो (त्यापूर्वीचे फक्त Activity Log मध्ये "कृती" म्हणून दिसतात).
3. 📊 Archive -> "तुलना": दोन वर्षांची फी (वसुली %) व उपस्थिती % वर्गनिहाय, फरकासह, CSV.
4. 🗓️ वर्ष-अखेर पूर्वतपासणी (Super): वर्गनिहाय बढती-आढावा, थकबाकी, LC नोंद/बाकी, डेटा-त्रुटी, Backup स्थिती, इशारे.
5. वेळ-मोजणी: Apps Script Editor मध्ये diagTimings चालवा -> Logs मध्ये प्रत्येक मुख्य action किती ms. (दोनदा चालवा; पहिली cold असते)
6. 📏 Sheet आकार: डेटा तपासणी -> Sheet आकार; मुख्यपृष्ठावर ७०%/८५% इशारा (Super/Master). रोज ईमेलसाठी setupCapacityWatchTrigger एकदा Run करा.
7. 📴 ऑफलाइन: वर्ग-यादी ७ दिवस (आधी २४ तास) जतन; ऑफलाइन पट्टी + "शेवटचे अपडेट"; Login ऑफलाइन असल्यास स्पष्ट संदेश.
   + शिक्षकाच्या वर्ग-यादीत नाव/Reg No आता escape होतात.

Deploy
1) Spreadsheet ची प्रत -> apps_script/Code.gs paste -> Deploy > Manage deployments > Edit > New version.
2) GitHub: github_pages मधील index.html, script.js, style.css, sw.js replace करा.
3) चाचण्या: bash tests/run_all.sh
