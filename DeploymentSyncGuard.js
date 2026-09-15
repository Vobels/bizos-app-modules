// ============================================================
// DeploymentSyncGuard.js
// Ensures the master database is repaired before upgrade/payment work.
// ============================================================

function ensureBizOSReady_() {
  try {
    var result = ensureBizOSMasterDatabase();
    console.log('BizOS database check:', JSON.stringify(result));
    return result;
  } catch (e) {
    console.error('BizOS database check failed:', e);
    return {success:false, message:e.message};
  }
}
