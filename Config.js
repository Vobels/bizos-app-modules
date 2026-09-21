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
      usd: 800,
      naira: 1200000,
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
      currency: "USD",
      convertToLocal: true,
    },
    exchangeRateFallback: 1500,
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
    freeModules: ["Ecommerce"],
    access: {
      demo: ["Finance", "Ecommerce", "Sales", "CRM", "HR", "Logistics", "Tax", "Agro", "Productivity", "POS", "Attendance", "Warehouse"],
      starter: ["Ecommerce"],
      sovereign: ["Finance", "Ecommerce", "Sales", "CRM", "HR", "Logistics", "Tax", "Agro", "Productivity", "POS", "Attendance", "Warehouse"],
      enterprise: ["Finance", "Ecommerce", "Sales", "CRM", "HR", "Logistics", "Tax", "Agro", "Productivity", "POS", "Attendance", "Warehouse"],
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
    starter: {
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
    enterprise: {
      name: "Enterprise",
      isReadOnly: false,
      maxRecords: null,
      badge: "🏢 Enterprise",
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
    defaultTier: "starter",
    defaultCountry: "Nigeria",
  },

  // ============================================================
  // 🆕 CLIENT CONFIGURATION (MOVED INSIDE CONFIG)
  // ============================================================
  CLIENT: {
    defaultPrimaryColor: '#5D2A86',
    defaultTier: 'sovereign',
    moduleAccess: {
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

function getServiceFee() {
  return CONFIG.PRICING.services.default;
}

function getFreeModules() {
  return CONFIG.MODULES.freeModules;
}

function getAllModules() {
  return CONFIG.MODULES.all;
}

function getModuleAccess(tier) {
  return CONFIG.MODULES.access[tier] || CONFIG.MODULES.access.starter;
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