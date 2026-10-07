// ============================================================
// V12AuditTests.js - NON-DESTRUCTIVE AUDIT LOG WRAPPERS
// ============================================================
// These wrappers do not deploy, mutate client data, or change registry state.
// They make returned audit objects visible in the Apps Script execution log.
// ============================================================

function testV12ReleaseCompletenessLog(businessId) {
  var result = auditV12ReleaseCompleteness(businessId);
  Logger.log('V12 RELEASE COMPLETENESS: ' + JSON.stringify(result));
  return result;
}

function testV12DeploymentRuntimeAuditLog(businessId) {
  var result = auditV12ClientDeploymentRuntime(businessId);
  Logger.log('V12 DEPLOYED RUNTIME AUDIT: ' + JSON.stringify(result));
  return result;
}
