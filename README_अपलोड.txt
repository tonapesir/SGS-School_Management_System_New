SGS School Management System — V19.59

बदल: LC पान
1. Reg No शोधल्यावर "मूळ विद्यार्थी माहिती" आधीप्रमाणे Students मधून येते.
   त्याचबरोबर "LC" Sheet मध्ये त्या Reg No ची नोंद शोधली जाते:
   - सापडल्यास: LC विशेष माहिती + "इतर शैक्षणिक माहिती" चे सर्व फील्ड आपोआप भरतात (शेवटची नोंद; एकाहून अधिक असल्यास संख्या सांगते).
   - न सापडल्यास: LC-विशेष माहिती रिकामी. (मागील विद्यार्थ्याची मूल्ये आता राहत नाहीत.)
     प्रवेश दिनांक / प्रवेशावेळी इयत्ता Students मधून येतात; Conduct/Progress = चांगला/चांगली (फक्त डीफॉल्ट).
2. Clear: संपूर्ण फॉर्म रिकामा; फक्त राष्ट्रीयत्व=भारतीय, मातृभाषा=मराठी, Conduct=चांगला, Progress=चांगली हे डीफॉल्ट परत येतात.
   Clear नंतर उशिरा आलेले LC-उत्तर फॉर्म पुन्हा भरत नाही.

Deploy (दोन्ही आवश्यक — नवा सर्व्हर action getLCByReg)
1) Spreadsheet ची प्रत -> apps_script/Code.gs paste -> Deploy > Manage deployments > Edit > New version.
2) GitHub: github_pages मधील index.html, script.js, style.css, sw.js replace करा.
3) चाचण्या: bash tests/run_all.sh
