// ============================================================
// ProvisioningLock.js
// One lock owner for client provisioning.
// Internal provisioning helpers must not acquire the Script Lock.
// ============================================================

function withClientProvisioningLock_(requestId, callback) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return {
      success: false,
      code: 'PROVISIONING_BUSY',
      message: 'Your workspace is already being set up. Please wait a moment and try again.'
    };
  }

  try {
    return callback();
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}
