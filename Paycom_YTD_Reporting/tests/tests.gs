var testManagersSheetName = "Test_PaycomManagersByEmailFromTracker",
    testPaycomAnalysisSheetName = "Test_PaycomYTDAnalysis",
    configSheetName = "Config";

function test_sendReportToManagers() {
  sendReportToManagers(testManagersSheetName, testPaycomAnalysisSheetName);
}

function test_sendReportToInterns() {
  sendReportToInterns(testPaycomAnalysisSheetName)
}
