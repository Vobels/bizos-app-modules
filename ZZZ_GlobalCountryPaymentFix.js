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

  getLocalizedPricing = function (country) {
    var name = String(country || '').trim();

    // Nigeria has an explicit local price in central configuration.
    if (name === 'Nigeria' && typeof CONFIG !== 'undefined' && CONFIG.PRICING && CONFIG.PRICING.sovereign) {
      var naira = Number(CONFIG.PRICING.sovereign.naira || 0);
      if (naira > 0) {
        return {
          sovereign: { price: naira, currency: '₦', code: 'NGN' },
          enterprise: { price: 2990, currency: '$', code: 'USD' }
        };
      }
    }

    // All other countries are accepted. Where no local currency/payment
    // configuration exists, charge the global USD price rather than rejecting
    // the country.
    return {
      sovereign: { price: 500, currency: ' },
      enterprise: { price: 2990, currency: '$', code: 'USD' }
    };
  };

  console.log('ZZZ global country/payment compatibility fix loaded');
})();
, code: 'USD' },
      enterprise: { price: 2990, currency: '$', code: 'USD' }
    };
  };

  console.log('ZZZ global country/payment compatibility fix loaded');
})();
