// config.gs
// ============================================================
// 📁 Config.gs - CENTRALIZED BACKEND CONFIGURATION
// ============================================================

var CONFIG = {
  // ============================================================
  // 🏢 APP IDENTITY
  // ============================================================
  APP: {
    name: "BizOS",
    tagline: "Business Operating System",
    domain: "bizos.higroups.com",
    supportEmail: "bizgroups@gmail.com",
    supportPhone: "+234 123 456 7890",
    year: new Date().getFullYear(),
    
    colors: {
      primary: "#5D2A86",
      primaryLight: "#7C3AED",
      primaryDark: "#4A1F6B",
      secondary: "#10B981",
      warning: "#F59E0B",
      danger: "#EF4444",
    }
  },

  // ============================================================
  // 🔐 AUTHENTICATION
  // ============================================================
  AUTH: {
    demo: {
      email: "demo@bizos.com",
      password: "demo123",
      business: "Demo Business",
      isReadOnly: true,
    },
    sessionTimeout: 3600, // 1 hour
    passwordMinLength: 6,
  },

  // ============================================================
  // 💰 PRICING - One Time Payment Only
  // ============================================================
  PRICING: {
    sovereign: {
      // USD is the single authoritative base price. Customer-facing
      // currencies are calculated from the admin-managed FX table below.
      usd: 500,
      description: "Lifetime access + Custom Domain + White Label",
      includes: [
        "All 12 Business Modules",
        "White Label + Custom Domain",
        "Your Own Google Sheet Backend",
        "Lifetime Updates",
        "24/7 Email Support"
      ]
    },
    services: {
      backup: 5,
      maintenance: 5,
      customization: 5,
      support: 5,
      default: 5,
      description: "Optional service requests. Only pay when you need help."
    },
    payment: {
      provider: "Paystack",
      baseCurrency: "USD",
      // Gateway capabilities are resolved separately from pricing.
      // Paystack USD remains feature-flagged until USD settlement is enabled.
      paystackUsdEnabledProperty: "PAYSTACK_USD_ENABLED",
      defaultCurrency: "USD",
      countryDefaultCurrencies: { Nigeria: "NGN" },
      currencyCountryRestrictions: { NGN: ["Nigeria"] },
      flutterwaveSupportedCurrencies: [
        "USD", "GBP", "EUR", "CAD", "XAF", "COP", "EGP", "GHS",
        "KES", "INR", "NGN", "RWF", "SLL", "ZAR", "TZS", "UGX", "XOF", "ZMW"
      ]
    },
    // Rates are expressed as 1 USD = rate units of the target currency.
    // Admin can update these occasionally; checkout snapshots the rate used.
    exchangeRates: {
      USD: { rate: 1, name: "US Dollar", symbol: "$", decimals: 2, enabled: true },
      NGN: { rate: 1500, name: "Nigerian Naira", symbol: "₦", decimals: 0, enabled: true },
      GBP: { rate: 0.76, name: "British Pound", symbol: "£", decimals: 2, enabled: true },
      EUR: { rate: 0.86, name: "Euro", symbol: "€", decimals: 2, enabled: true },
      CAD: { rate: 1.38, name: "Canadian Dollar", symbol: "C$", decimals: 2, enabled: true }
    }
  },

  // ============================================================
  // 📦 MODULES
  // ============================================================
  MODULES: {
    all: [
      "Finance", "Sales", "Ecommerce", "CRM", "HR", 
      "Logistics", "Tax", "Agro", "Productivity", 
      "POS", "Attendance", "Warehouse"
    ],
    freeModules: ["Finance", "Ecommerce"],
    access: {
      demo: ["Finance", "Ecommerce", "Sales", "CRM", "HR", "Logistics", "Tax", "Agro", "Productivity", "POS", "Attendance", "Warehouse"],
      free: ["Finance", "Ecommerce"],
      sovereign: ["Finance", "Ecommerce", "Sales", "CRM", "HR", "Logistics", "Tax", "Agro", "Productivity", "POS", "Attendance", "Warehouse"],
    },
    icons: {
      Finance: "fa-calculator",
      Sales: "fa-chart-line",
      Ecommerce: "fa-cart-shopping",
      CRM: "fa-users",
      HR: "fa-briefcase",
      Logistics: "fa-truck",
      Tax: "fa-file-invoice",
      Agro: "fa-seedling",
      Productivity: "fa-clock",
      POS: "fa-cash-register",
      Attendance: "fa-fingerprint",
      Warehouse: "fa-warehouse",
    },
    colors: {
      Finance: "from-purple-500 to-purple-700",
      Sales: "from-blue-500 to-blue-700",
      Ecommerce: "from-orange-500 to-orange-700",
      CRM: "from-green-500 to-green-700",
      HR: "from-red-500 to-red-700",
      Logistics: "from-teal-500 to-teal-700",
      Tax: "from-yellow-500 to-yellow-700",
      Agro: "from-emerald-500 to-emerald-700",
      Productivity: "from-indigo-500 to-indigo-700",
      POS: "from-pink-500 to-pink-700",
      Attendance: "from-cyan-500 to-cyan-700",
      Warehouse: "from-slate-500 to-slate-700",
    },
    maturity: {
      Finance: 'core',
      Ecommerce: 'core',
      POS: 'core',
      Sales: 'foundation',
      CRM: 'foundation',
      HR: 'foundation',
      Logistics: 'foundation',
      Tax: 'foundation',
      Agro: 'foundation',
      Productivity: 'foundation',
      Attendance: 'foundation',
      Warehouse: 'foundation'
    },
    labels: {
      Finance: "Financial Management",
      Sales: "Sales Pipeline",
      Ecommerce: "E-commerce",
      CRM: "Customer Relations",
      HR: "Human Resources",
      Logistics: "Logistics",
      Tax: "Tax Compliance",
      Agro: "Agriculture",
      Productivity: "Productivity Suite",
      POS: "Point of Sale",
      Attendance: "Staff Attendance",
      Warehouse: "Warehouse Management",
    },
  },

  // ============================================================
  // 👤 TIERS
  // ============================================================
  TIERS: {
    demo: {
      name: "Demo",
      isReadOnly: true,
      maxRecords: 50,
      badge: "(Demo)",
      upgradeMessage: "Demo Mode - Read Only",
    },
    free: {
      name: "Free",
      isReadOnly: false,
      maxRecords: 500,
      badge: "Free",
      upgradeMessage: "Upgrade to Sovereign for lifetime access",
    },
    sovereign: {
      name: "Sovereign",
      isReadOnly: false,
      maxRecords: null,
      badge: "⭐ Sovereign",
      upgradeMessage: "",
    },
  },

  // ============================================================
  // 🌍 COUNTRY CONFIGURATION
  // ============================================================
  COUNTRIES: {
    supported: [
      "Nigeria", "Kenya", "Ghana", "SouthAfrica", "Egypt",
      "USA", "Canada", "UK", "Germany", "France", 
      "Turkey", "India", "Philippines", "Australia"
    ],
    requirements: {
      Nigeria: {
        requiredFields: ["cacNumber", "taxId"],
        fieldLabels: { cacNumber: "CAC Registration Number", taxId: "TIN" },
        placeholders: { cacNumber: "RC1234567", taxId: "12345678-1234" },
        documentTypes: ["CAC Certificate", "Tax Certificate"],
        processingTime: "24-48 hours"
      },
      Kenya: {
        requiredFields: ["businessRegNumber", "kraPin"],
        fieldLabels: { businessRegNumber: "Business Registration Number", kraPin: "KRA PIN" },
        placeholders: { businessRegNumber: "PVT-1234567", kraPin: "A123456789B" },
        documentTypes: ["Business Registration Certificate", "KRA PIN Certificate"],
        processingTime: "24-48 hours"
      },
      USA: {
        requiredFields: ["ein"],
        fieldLabels: { ein: "EIN (Employer Identification Number)" },
        placeholders: { ein: "12-3456789" },
        documentTypes: ["IRS EIN Letter"],
        processingTime: "24-48 hours"
      },
      UK: {
        requiredFields: ["companyNumber", "utr"],
        fieldLabels: { companyNumber: "Company Registration Number", utr: "UTR" },
        placeholders: { companyNumber: "12345678", utr: "1234567890" },
        documentTypes: ["Companies House Certificate", "HMRC UTR Letter"],
        processingTime: "24-48 hours"
      },
      Germany: {
        requiredFields: ["handelsregisterNumber", "taxNumber"],
        fieldLabels: { handelsregisterNumber: "Handelsregister Number", taxNumber: "Tax Number" },
        placeholders: { handelsregisterNumber: "HRB 12345", taxNumber: "123/456/78901" },
        documentTypes: ["Handelsregister Extract", "Tax Registration Certificate"],
        processingTime: "24-48 hours"
      },
      France: {
        requiredFields: ["siret"],
        fieldLabels: { siret: "SIRET Number" },
        placeholders: { siret: "123 456 789 00012" },
        documentTypes: ["SIRET Registration", "Kbis Extract"],
        processingTime: "24-48 hours"
      },
      Turkey: {
        requiredFields: ["taxNumber", "mersis"],
        fieldLabels: { taxNumber: "Tax Number", mersis: "MERSIS Number" },
        placeholders: { taxNumber: "1234567890", mersis: "1234567890" },
        documentTypes: ["Tax Certificate", "Trade Registry Gazette"],
        processingTime: "24-48 hours"
      },
      India: {
        requiredFields: ["gst"],
        fieldLabels: { gst: "GST Number" },
        placeholders: { gst: "22AAAAA0000A1Z5" },
        documentTypes: ["GST Certificate"],
        processingTime: "24-48 hours"
      },
      Philippines: {
        requiredFields: ["secNumber", "tin"],
        fieldLabels: { secNumber: "SEC Registration Number", tin: "TIN" },
        placeholders: { secNumber: "CS202412345", tin: "123-456-789-000" },
        documentTypes: ["SEC Certificate", "BIR Form 2303"],
        processingTime: "24-48 hours"
      },
      Australia: {
        requiredFields: ["abn"],
        fieldLabels: { abn: "ABN (Australian Business Number)" },
        placeholders: { abn: "12 345 678 901" },
        documentTypes: ["ABN Registration"],
        processingTime: "24-48 hours"
      },
      Canada: {
        requiredFields: ["businessNumber"],
        fieldLabels: { businessNumber: "Business Number" },
        placeholders: { businessNumber: "123456789RC0001" },
        documentTypes: ["Business Registration Certificate"],
        processingTime: "24-48 hours"
      },
      SouthAfrica: {
        requiredFields: ["registrationNumber"],
        fieldLabels: { registrationNumber: "CIPC Registration Number" },
        placeholders: { registrationNumber: "2019/123456/07" },
        documentTypes: ["CIPC Certificate"],
        processingTime: "24-48 hours"
      },
      Egypt: {
        requiredFields: ["taxNumber"],
        fieldLabels: { taxNumber: "Tax Number" },
        placeholders: { taxNumber: "123-456-789" },
        documentTypes: ["Tax Card"],
        processingTime: "24-48 hours"
      },
      Ghana: {
        requiredFields: ["businessRegistrationNumber", "taxIdentificationNumber"],
        fieldLabels: { businessRegistrationNumber: "Business Registration Number", taxIdentificationNumber: "Tax Identification Number" },
        placeholders: { businessRegistrationNumber: "BN123456789", taxIdentificationNumber: "C0000000000" },
        documentTypes: ["Business Registration Certificate", "Tax Identification Card"],
        processingTime: "24-48 hours"
      }
    },
    default: "Nigeria",
  },

  // ============================================================
  // 📨 EMAIL
  // ============================================================
  EMAIL: {
    fromName: "BizOS",
    fromEmail: "bizgroups@gmail.com",
    supportEmail: "bizgroups@gmail.com",
  },

  // ============================================================
  // 🔗 URLS
  // ============================================================
  URLS: {
    base: "https://bizos.higroups.com",
    support: "mailto:bizgroups@gmail.com",
  },

  // ============================================================
  // ⚙️ FEATURE FLAGS
  // ============================================================
  FEATURES: {
    enableDemo: true,
    enableWhiteLabel: true,
    enableServiceRequests: true,
    enableCharts: true,
    enableNotifications: true,
    enableExport: true,
    enableImport: true,
    enableTeamManagement: true,
    enableUpgrade: true,
  },

  // ============================================================
  // 📊 DEFAULTS
  // ============================================================
  DEFAULTS: {
    recordsPerPage: 10,
    dateFormat: "YYYY-MM-DD",
    currency: "USD",
    currencySymbol: "$",
    defaultTier: "free",
    defaultCountry: "Nigeria",
  },

  // ============================================================
  // 🆕 CLIENT CONFIGURATION (MOVED INSIDE CONFIG)
  // ============================================================
  CLIENT: {
    defaultPrimaryColor: '#5D2A86',
    defaultTier: 'sovereign',
    moduleAccess: {
      free: ['Finance', 'Ecommerce'],
      sovereign: ['Finance', 'Ecommerce', 'Sales', 'CRM', 'HR', 'Logistics', 'Tax', 'Agro', 'Productivity', 'POS', 'Attendance', 'Warehouse']
    },
    // Client provisioning settings
    provisioning: {
      autoCreateSheet: true,
      autoDeployLanding: true,
      sendWelcomeEmail: true,
      defaultFolderName: 'BizOS_Client_Workspaces'
    }
  }
};

// ============================================================
// 🛠 HELPER FUNCTIONS
// ============================================================

function getConfig() {
  return CONFIG;
}

function getAppName() {
  return CONFIG.APP.name;
}

function getAppDomain() {
  return CONFIG.APP.domain;
}

function getSupportEmail() {
  return CONFIG.EMAIL.supportEmail;
}

function getSovereignPrice() {
  return CONFIG.PRICING.sovereign;
}

function getPaymentExchangeRates_() {
  return (CONFIG.PRICING && CONFIG.PRICING.exchangeRates) || {};
}

function getPaymentCurrencyConfig_(currency) {
  var code = String(currency || '').trim().toUpperCase();
  var rates = getPaymentExchangeRates_();
  var entry = rates[code];
  if (!entry || entry.enabled === false) return null;
  var rate = Number(entry.rate);
  if (!isFinite(rate) || rate <= 0) return null;
  return {
    code: code,
    rate: rate,
    name: entry.name || code,
    symbol: entry.symbol || code,
    decimals: Number.isFinite(Number(entry.decimals)) ? Number(entry.decimals) : 2,
    enabled: true
  };
}

function convertBizOSBasePriceToCurrency_(currency) {
  var code = String(currency || '').trim().toUpperCase();
  var config = getPaymentCurrencyConfig_(code);
  if (!config) return {success:false,message:'This currency is not currently available for BizOS checkout.'};
  var baseAmount = Number(CONFIG.PRICING.sovereign.usd);
  if (!isFinite(baseAmount) || baseAmount <= 0) {
    return {success:false,message:'BizOS base pricing is not configured correctly.'};
  }
  var raw = baseAmount * config.rate;
  var factor = Math.pow(10, Math.max(0, config.decimals));
  var amount = Math.round(raw * factor) / factor;
  return {
    success:true,
    baseAmount:baseAmount,
    baseCurrency:'USD',
    rate:config.rate,
    rateSource:'admin',
    rateUpdatedAt:CONFIG.PRICING.exchangeRatesUpdatedAt || '',
    currency:config.code,
    price:amount,
    symbol:config.symbol,
    name:config.name,
    decimals:config.decimals
  };
}

function getPaymentPricingConfigForAdmin() {
  return {
    success:true,
    baseCurrency:'USD',
    basePrice:Number(CONFIG.PRICING.sovereign.usd),
    exchangeRates:getPaymentExchangeRates_(),
    updatedAt:CONFIG.PRICING.exchangeRatesUpdatedAt || ''
  };
}

function updatePaymentExchangeRates(sessionId, rates, basePrice) {
  requireAdminSession_(sessionId);
  try {
    rates = rates || {};
    var current = getPaymentExchangeRates_();
    var next = {};
    Object.keys(current).forEach(function(code) {
      var existing = current[code] || {};
      var incoming = rates[code];
      var value = incoming === undefined || incoming === '' ? existing.rate : Number(incoming);
      if (code === 'USD') value = 1;
      if (!isFinite(value) || value <= 0) throw new Error('Invalid exchange rate for ' + code + '.');
      next[code] = {
        rate:value,
        name:existing.name || code,
        symbol:existing.symbol || code,
        decimals:Number.isFinite(Number(existing.decimals)) ? Number(existing.decimals) : 2,
        enabled:existing.enabled !== false
      };
    });

    var parsedBase = basePrice === undefined || basePrice === '' ? Number(CONFIG.PRICING.sovereign.usd) : Number(basePrice);
    if (!isFinite(parsedBase) || parsedBase <= 0) throw new Error('Base USD price must be greater than zero.');

    CONFIG.PRICING.sovereign.usd = parsedBase;
    CONFIG.PRICING.exchangeRates = next;
    CONFIG.PRICING.exchangeRatesUpdatedAt = new Date().toISOString();

    // Keep the editable config persistent without requiring a code deployment.
    PropertiesService.getScriptProperties().setProperty(
      'BIZOS_PAYMENT_PRICING_CONFIG',
      JSON.stringify({
        basePrice:parsedBase,
        exchangeRates:next,
        updatedAt:CONFIG.PRICING.exchangeRatesUpdatedAt
      })
    );

    return getPaymentPricingConfigForAdmin();
  } catch (error) {
    console.error('updatePaymentExchangeRates error:', error);
    return {success:false,message:error.message || 'Payment pricing could not be updated.'};
  }
}

function loadPersistedPaymentPricingConfig_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('BIZOS_PAYMENT_PRICING_CONFIG');
    if (!raw) return;
    var saved = JSON.parse(raw);
    if (!saved || !saved.exchangeRates) return;
    var base = Number(saved.basePrice);
    if (isFinite(base) && base > 0) CONFIG.PRICING.sovereign.usd = base;
    Object.keys(saved.exchangeRates).forEach(function(code) {
      if (!CONFIG.PRICING.exchangeRates[code]) return;
      var savedEntry = saved.exchangeRates[code];
      var rate = Number(savedEntry.rate);
      if (!isFinite(rate) || rate <= 0) return;
      CONFIG.PRICING.exchangeRates[code].rate = code === 'USD' ? 1 : rate;
      if (savedEntry.enabled !== undefined) CONFIG.PRICING.exchangeRates[code].enabled = savedEntry.enabled !== false;
    });
    CONFIG.PRICING.exchangeRatesUpdatedAt = saved.updatedAt || '';
  } catch (error) {
    console.error('loadPersistedPaymentPricingConfig_ error:', error);
  }
}

loadPersistedPaymentPricingConfig_();

function getServiceFee() {
  return CONFIG.PRICING.services.default;
}

function getFreeModules() {
  return CONFIG.MODULES.freeModules;
}

function getAllModules() {
  return CONFIG.MODULES.all;
}

function normalizeSubscriptionTier(tier) {
  var raw = String(tier || 'free').trim().toLowerCase();
  if (raw === 'sovereign' || raw === 'tier2' || raw === 'professional' || raw === 'enterprise') return 'sovereign';
  return 'free';
}

function getModuleAccess(tier) {
  return CONFIG.MODULES.access[normalizeSubscriptionTier(tier)] || CONFIG.MODULES.access.free;
}

function getCountryRequirements(country) {
  return CONFIG.COUNTRIES.requirements[country] || null;
}

function getSupportedCountries() {
  return CONFIG.COUNTRIES.supported;
}

function isFeatureEnabled(feature) {
  return CONFIG.FEATURES[feature] !== false;
}

function isDemoMode() {
  return CONFIG.FEATURES.enableDemo;
}

function getDefaultTier() {
  return CONFIG.DEFAULTS.defaultTier;
}

function getDefaultCountry() {
  return CONFIG.DEFAULTS.defaultCountry;
}

// ============================================================
// 🆕 CLIENT HELPER FUNCTIONS
// ============================================================

function getClientDefaultColor() {
  return CONFIG.CLIENT.defaultPrimaryColor;
}

function getClientDefaultTier() {
  return CONFIG.CLIENT.defaultTier;
}

function getClientModuleAccess(tier) {
  return CONFIG.CLIENT.moduleAccess[tier] || CONFIG.CLIENT.moduleAccess.sovereign;
}

function getClientProvisioningConfig() {
  return CONFIG.CLIENT.provisioning;
}

// ============================================================
// 📦 EXPOSE TO FRONTEND (via scriptlets)
// ============================================================

function getFrontendConfig() {
  return {
    app: CONFIG.APP,
    modules: {
      all: CONFIG.MODULES.all,
      free: CONFIG.MODULES.freeModules,
      labels: CONFIG.MODULES.labels,
      maturity: CONFIG.MODULES.maturity,
      icons: CONFIG.MODULES.icons,
      colors: CONFIG.MODULES.colors,
    },
    pricing: {
      sovereign: CONFIG.PRICING.sovereign,
      services: CONFIG.PRICING.services,
    },
    tiers: CONFIG.TIERS,
    features: CONFIG.FEATURES,
    defaults: CONFIG.DEFAULTS,
    urls: CONFIG.URLS,
    client: {
      defaultPrimaryColor: CONFIG.CLIENT.defaultPrimaryColor,
      defaultTier: CONFIG.CLIENT.defaultTier,
    }
  };
}

console.log('✅ Config.gs loaded successfully');

// ============================================================
// MASTER SPREADSHEET CONTEXT
// ============================================================
// Time-based provisioning workers do not have an active spreadsheet
// context. Store the authoritative master spreadsheet ID once so
// background provisioning can always reach the same database.
function getBizOSMasterSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var spreadsheetId = String(
    props.getProperty('BIZOS_MASTER_SPREADSHEET_ID') || ''
  ).trim();

  if (spreadsheetId) {
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      console.error('Stored BizOS master spreadsheet could not be opened:', error);
    }
  }

  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    try {
      props.setProperty('BIZOS_MASTER_SPREADSHEET_ID', active.getId());
    } catch (ignore) {}
    return active;
  }

  throw new Error(
    'BizOS master spreadsheet is not configured for background provisioning. ' +
    'Open the master spreadsheet-bound Apps Script project once and authorize the provisioning worker.'
  );
}

function configureBizOSMasterSpreadsheet_(spreadsheetId) {
  spreadsheetId = String(spreadsheetId || '').trim();
  if (!spreadsheetId) {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) spreadsheetId = active.getId();
  }
  if (!spreadsheetId) {
    return {
      success: false,
      code: 'MASTER_SPREADSHEET_ID_REQUIRED',
      message: 'Open the master spreadsheet-bound Apps Script project and run this setup again.'
    };
  }

  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    PropertiesService.getScriptProperties().setProperty('BIZOS_MASTER_SPREADSHEET_ID', ss.getId());
    return {success:true,spreadsheetId:ss.getId(),spreadsheetName:ss.getName()};
  } catch (error) {
    return {
      success:false,
      code:'MASTER_SPREADSHEET_INVALID',
      message:error && error.message ? error.message : 'The master spreadsheet could not be opened.'
    };
  }
}
