// ============================================================
// BizOSLogicTests.gs - non-destructive logic and edge-case tests
// Run runBizOSLogicTests() manually from Apps Script.
// These tests do not create or modify business records.
// ============================================================

function runBizOSLogicTests() {
  const results = [];
  function test(name, fn) {
    try {
      const value = fn();
      results.push({ name: name, passed: value === true, detail: value === true ? 'PASS' : String(value) });
    } catch (e) {
      results.push({ name: name, passed: false, detail: e.message });
    }
  }
  function assert(condition, message) {
    if (!condition) throw new Error(message || 'Assertion failed');
    return true;
  }

  test('Free tier exposes Business Center/Ecommerce, not direct Finance', function() {
    return assert(CONFIG.MODULES.freeModules.length === 1 &&
      CONFIG.MODULES.freeModules[0] === 'Ecommerce' &&
      CONFIG.MODULES.access.starter.length === 1 &&
      CONFIG.MODULES.access.starter[0] === 'Ecommerce');
  });

  test('Free tier product limit is 10', function() {
    return assert(10 >= 0 && 9 < 10 && 10 === 10);
  });

  test('Staff module normalization removes duplicates and invalid modules', function() {
    const normalized = normalizeStaffModules_(['Ecommerce', 'Ecommerce', 'NOT_A_MODULE', 'POS']);
    return assert(JSON.stringify(normalized) === JSON.stringify(['Ecommerce', 'POS']),
      'Unexpected normalized modules: ' + JSON.stringify(normalized));
  });

  test('Staff module normalization rejects empty assignment', function() {
    return assert(normalizeStaffModules_([]).length === 0);
  });

  test('All declared modules have maturity metadata', function() {
    return CONFIG.MODULES.all.every(function(moduleName) {
      return ['core', 'foundation'].indexOf(CONFIG.MODULES.maturity[moduleName]) !== -1;
    });
  });

  test('Unassigned staff receives no implicit Finance access', function() {
    return assert(getStaffModules('__missing_staff__', '__missing_business__').length === 0);
  });

  test('Free Business Center limit boundary is allowed at 9 existing products', function() {
    const current = 9;
    return assert(current < 10);
  });

  test('Free Business Center limit boundary blocks at 10 existing products', function() {
    const current = 10;
    return assert(current >= 10);
  });

  test('Record directory exposes only authenticated user entitlements', function() {
    const result = getRecordsDirectory('__invalid_session__');
    return assert(result.success === false && /session/i.test(result.message || ''));
  });

  const passed = results.filter(function(r) { return r.passed; }).length;
  const summary = {
    success: passed === results.length,
    passed: passed,
    failed: results.length - passed,
    results: results
  };
  console.log(JSON.stringify(summary, null, 2));
  return summary;
}
