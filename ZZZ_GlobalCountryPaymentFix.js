// ============================================================
// ZZZ_GlobalCountryPaymentFix.js
// Global country/payment compatibility fix.
// ============================================================
// Unknown countries must not be rejected merely because BizOS does
// not have country-specific registration-field definitions yet.
// Payment falls back to USD when no local payment configuration exists.
// ============================================================

(function () {
  var originalCountryRequirements = getCountryRequirements;
  var originalPricing = getLocalizedPricing;

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

    if (known) {
      return {
        success: true,
        requiredFields: Array.isArray(known.requiredFields) ? known.requiredFields : [],
        fieldLabels: known.fieldLabels || {},
        placeholders: known.placeholders || {},
        documentTypes: known.documentTypes || [],
        processingTime: known.processingTime || '24-48 hours'
      };
    }

    // Global fallback: country is valid even when BizOS has no
    // country-specific compliance fields configured yet.
    return {
      success: true,
      requiredFields: [],
      fieldLabels: {},
      placeholders: {},
      documentTypes: [],
      processingTime: '24-48 hours'
    };
  };

  getLocalizedPricing = function (country) {
    var name = String(country || '').trim();
    var paymentConfig = null;

    try {
      if (typeof PAYMENT_COUNTRIES !== 'undefined' && PAYMENT_COUNTRIES) {
        paymentConfig = PAYMENT_COUNTRIES[name] || PAYMENT_COUNTRIES.default || null;
      }
    } catch (e) {}

    // Keep explicitly configured country pricing when available.
    if (paymentConfig) {
      var rate = Number(paymentConfig.multiplier || 1);
      var exchangeRate = 1500;
      var currency = String(paymentConfig.currency || 'USD').toUpperCase();
      var symbol = paymentConfig.symbol || (currency === 'NGN' ? '₦' : '$');
      return {
        sovereign: { price: Math.round(499 * exchangeRate * rate) / 100, currency: symbol, code: currency },
        enterprise: { price: Math.round(2990 * exchangeRate * rate) / 100, currency: symbol, code: currency }
      };
    }

    // Universal fallback: Paystack checkout uses USD for countries for which
    // BizOS has no local-currency configuration.
    return {
      sovereign: { price: 499, currency: '$', code: 'USD' },
      enterprise: { price: 2990, currency: '$', code: 'USD' }
    };
  };

  console.log('ZZZ global country/payment compatibility fix loaded');
})();
