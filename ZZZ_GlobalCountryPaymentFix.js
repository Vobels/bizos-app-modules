// ============================================================
// ZZZ_GlobalCountryPaymentFix.js
// Global country/payment compatibility fix.
// ============================================================

(function () {
  var originalCountryRequirements = getCountryRequirements;

  getCountryRequirements = function (country) {
    var name = String(country || '').trim();
    var known = null;

    try {
      if (typeof CONFIG !== 'undefined' && CONFIG.COUNTRIES && CONFIG.COUNTRIES.requirements) {
        known = CONFIG.COUNTRIES.requirements[name] || null;
      }
    } catch (e) {}

    if (!known && typeof originalCountryRequirements === 'function') {
      try { known = originalCountryRequirements(name); } catch (e) {}
    }

    return {
      success: true,
      requiredFields: known && Array.isArray(known.requiredFields) ? known.requiredFields : [],
      fieldLabels: known && known.fieldLabels ? known.fieldLabels : {},
      placeholders: known && known.placeholders ? known.placeholders : {},
      documentTypes: known && known.documentTypes ? known.documentTypes : [],
      processingTime: known && known.processingTime ? known.processingTime : '24-48 hours'
    };
  };



  console.log('ZZZ global country/payment compatibility fix loaded');
})();
