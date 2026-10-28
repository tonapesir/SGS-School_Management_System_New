SGS School Management System — V19.58

बदल (V19.57.1 नंतर)
1. LC: "शै. वर्ष सुरुवात दिनांक" -> "या वर्गात कधीपासून शिकत होता (महिना / वर्ष)": English महिन्याचा dropdown + manual वर्ष.
   छापील: "इयत्ता 7 त जून 2023 पासून" (मराठी महिना). Word डेटामध्ये खोटा "01" दिवस नाही — "जून 2023".
2. LC: Conduct = चांगला, Progress = चांगली. (आधीचा दोष: या दोन चौकटींना चुकीचा/डुप्लिकेट id होता, म्हणून टाकलेले मूल्य वाचले जात नव्हते.
   आता टाकलेले मूल्य छापले जाते व सर्व्हरला जाते. विद्यार्थी बदलल्यावर ही मूल्ये व "कधीपासून" रीसेट होतात.)
3. आवृत्ती-जुळणी: सर्व्हर/ॲप आवृत्ती जुळत नसल्यास सूचना (ॲप जुना -> Refresh; सर्व्हर जुना -> फक्त Super/Master ला "Deploy करा").
4. सुरक्षा (escape): नोटीस (title/message/postedBy), Diary remark, हजेरी कारण, विद्यार्थ्यांची नावे/Reg No (dropdown, याद्या) आता escape होतात.
5. tests/ + .github/workflows/tests.yml: प्रत्येक push वर आपोआप चाचण्या.

Deploy
1) Spreadsheet ची प्रत -> apps_script/Code.gs paste -> Deploy > Manage deployments > Edit > New version.
2) GitHub: github_pages मधील index.html, script.js, style.css, sw.js replace करा (बाकी फाइल्स न बदललेल्या).
3) (ऐच्छिक) tests/ फोल्डर व .github/ फोल्डर repository च्या root मध्ये ठेवा. स्थानिक चाचणी: bash tests/run_all.sh
   (Node + Playwright हवे; खऱ्या Sheet/Drive वर चालत नाही.)
