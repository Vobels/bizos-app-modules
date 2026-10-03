//dashboard.gs
// ==================== DASHBOARD FUNCTIONS (OPTIMIZED) ====================

// Helper function to safely get user or return null
function safeGetUser(sessionId) {
  try {
    return getUserFromSession(sessionId);
  } catch(e) {
    console.log("safeGetUser error:", e.message);
    return null;
  }
}

// Main entry point - automatically chooses fast or accurate based on period
function getDashboardKPIs(period, sessionId, forceAccurate = false) {
  // Handle case where parameters are swapped
  if (typeof period === 'string' && period.length > 30 && sessionId === undefined) {
    sessionId = period;
    period = 'this_month';
  }
  
  try {
    const user = safeGetUser(sessionId);
    if (!user) {
      console.log("getDashboardKPIs: No valid session, returning zeros");
      return { revenue: 0, expenses: 0, profit: 0, profitMargin: 0 };
    }
    
    if (period === 'this_month' && !forceAccurate) {
      return getDashboardKPIsFast(sessionId);
    }
    
    return getDashboardKPIsAccurate(sessionId, period);
  } catch (error) {
    console.error('Error getting KPIs:', error);
    return { revenue: 0, expenses: 0, profit: 0, profitMargin: 0 };
  }
}

// FAST VERSION - For dashboard (last 30 days only, cached)
function getDashboardKPIsFast(sessionId) {
  const user = safeGetUser(sessionId);
  if (!user) return { revenue: 0, expenses: 0, profit: 0, profitMargin: 0 };
  
  const cache = CacheService.getScriptCache();
  const cacheKey = `dashboard_fast_${sessionId}`;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);
  
  try {
    const workspace = getWorkspaceFile(sessionId);
    const sheet = workspace.getSheetByName('Financial_Data');
    if (!sheet) return { revenue: 0, expenses: 0, profit: 0, profitMargin: 0 };
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { revenue: 0, expenses: 0, profit: 0, profitMargin: 0 };
    
    const startRow = Math.max(1, lastRow - 2000);
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn()).getValues();
    const headers = data[0];
    
    const dateCol = headers.indexOf('Date');
    const typeCol = headers.indexOf('Type');
    const amountCol = headers.indexOf('Amount');
    
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    thirtyDaysAgo.setHours(0, 0, 0, 0);
    
    let revenue = 0, expenses = 0;
    
    for (let i = 1; i < data.length; i++) {
      const rowDate = new Date(data[i][dateCol]);
      if (rowDate < thirtyDaysAgo) continue;
      
      const amount = parseFloat(data[i][amountCol]) || 0;
      const type = data[i][typeCol];
      
      if (type && type.toLowerCase() === 'revenue') {
        revenue += amount;
      } else if (type && type.toLowerCase() === 'expense') {
        expenses += Math.abs(amount);
      }
    }
    
    const profit = revenue - expenses;
    const profitMargin = revenue > 0 ? (profit / revenue) * 100 : 0;
    const result = { revenue, expenses, profit, profitMargin };
    
    cache.put(cacheKey, JSON.stringify(result), 120);
    return result;
  } catch(e) {
    console.error("getDashboardKPIsFast error:", e);
    return { revenue: 0, expenses: 0, profit: 0, profitMargin: 0 };
  }
}

// ACCURATE VERSION - For reports
function getDashboardKPIsAccurate(sessionId, period) {
  const user = safeGetUser(sessionId);
  if (!user) return { revenue: 0, expenses: 0, profit: 0, profitMargin: 0 };
  
  try {
    const workspace = getWorkspaceFile(sessionId);
    const sheet = workspace.getSheetByName('Financial_Data');
    if (!sheet) return { revenue: 0, expenses: 0, profit: 0, profitMargin: 0 };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const dateCol = headers.indexOf('Date');
    const typeCol = headers.indexOf('Type');
    const amountCol = headers.indexOf('Amount');
    
    const now = new Date();
    let revenue = 0, expenses = 0;
    let startDate = null;
    
    switch(period) {
      case 'today':
        startDate = new Date(now.setHours(0, 0, 0, 0));
        break;
      case 'this_week':
        startDate = new Date(now);
        startDate.setDate(now.getDate() - now.getDay());
        startDate.setHours(0, 0, 0, 0);
        break;
      case 'this_month':
        startDate = new Date(now.getFullYear(), now.getMonth(), 1);
        break;
      case 'this_year':
        startDate = new Date(now.getFullYear(), 0, 1);
        break;
      case 'last_month':
        startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
        const endDate = new Date(now.getFullYear(), now.getMonth(), 0);
        for (let i = 1; i < data.length; i++) {
          const rowDate = new Date(data[i][dateCol]);
          if (rowDate >= startDate && rowDate <= endDate) {
            const amount = parseFloat(data[i][amountCol]) || 0;
            if (data[i][typeCol] === 'revenue') revenue += amount;
            else if (data[i][typeCol] === 'expense') expenses += Math.abs(amount);
          }
        }
        const profitLastMonth = revenue - expenses;
        return { revenue, expenses, profit: profitLastMonth, profitMargin: revenue > 0 ? (profitLastMonth / revenue) * 100 : 0 };
      default:
        startDate = null;
    }
    
    for (let i = 1; i < data.length; i++) {
      if (startDate) {
        const rowDate = new Date(data[i][dateCol]);
        if (rowDate < startDate) continue;
      }
      
      const amount = parseFloat(data[i][amountCol]) || 0;
      const type = data[i][typeCol];
      
      if (type && type.toLowerCase() === 'revenue') {
        revenue += amount;
      } else if (type && type.toLowerCase() === 'expense') {
        expenses += Math.abs(amount);
      }
    }
    
    const profit = revenue - expenses;
    const profitMargin = revenue > 0 ? (profit / revenue) * 100 : 0;
    
    return { revenue, expenses, profit, profitMargin };
  } catch(e) {
    console.error("getDashboardKPIsAccurate error:", e);
    return { revenue: 0, expenses: 0, profit: 0, profitMargin: 0 };
  }
}

// Optimized Enhanced Dashboard Data
function getEnhancedDashboardData(period, sessionId) {
  // Handle case where parameters are swapped
  if (typeof period === 'string' && period.length > 30 && sessionId === undefined) {
    sessionId = period;
    period = 'this_month';
  }
  
  // Return empty dashboard template
  const emptyDashboard = {
    success: true,
    dashboard: {
      kpis: {
        revenue: { value: 0, change: 0 },
        expenses: { value: 0, change: 0 },
        netProfit: { value: 0, change: 0 },
        profitMargin: { value: 0, change: 0 }
      },
      trends: { categoryBreakdown: {}, daily: {}, weekly: {} },
      moduleSummary: {},
      health: { overall: 0, metrics: [], recommendations: [] },
      charts: {
        revenueVsExpenses: { labels: [], revenue: [], expenses: [] },
        expenseBreakdown: { labels: [], values: [] }
      },
      recentActivity: [],
      alerts: []
    }
  };
  
  try {
    const user = safeGetUser(sessionId);
    if (!user) {
      console.log("getEnhancedDashboardData: No valid session, returning empty dashboard");
      return emptyDashboard;
    }
    
    const workspace = getWorkspaceFile(sessionId);
    const useAccurate = (period === 'this_year' || period === 'last_month');
    const financialMetrics = getDetailedFinancialMetrics(workspace, period, useAccurate);
    const moduleSummary = getModuleSummary(sessionId);
    const businessSnapshot = getDashboardBusinessSnapshot(workspace, moduleSummary);
    const businessHealth = calculateBusinessHealth(workspace, financialMetrics, businessSnapshot);
    const chartData = getChartDataOptimized(workspace, period);
    const recentActivity = getRecentActivityAcrossModules(workspace, 10, moduleSummary);
    const alerts = generateBusinessAlerts(workspace, financialMetrics, moduleSummary, businessSnapshot);
    
    return {
      success: true,
      dashboard: {
        kpis: financialMetrics.kpis,
        trends: financialMetrics.trends,
        moduleSummary: moduleSummary,
        health: businessHealth,
        charts: chartData,
        recentActivity: recentActivity,
        alerts: alerts,
        businessSnapshot: businessSnapshot
      }
    };
    
  } catch (error) {
    console.error('Dashboard error:', error);
    return emptyDashboard;
  }
}

// Optimized Recent Transactions
function getRecentTransactions(limit, sessionId) {
  if (typeof limit === 'string' && limit.length > 30 && sessionId === undefined) {
    sessionId = limit;
    limit = 10;
  }
  
  const user = safeGetUser(sessionId);
  if (!user) return [];
  
  try {
    const workspace = getWorkspaceFile(sessionId);
    const sheet = workspace.getSheetByName('Financial_Data');
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    
    const startRow = Math.max(1, lastRow - 100);
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn()).getValues();
    const headers = data[0];
    
    const transactions = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
    
    return transactions.sort((a, b) => new Date(b.Date) - new Date(a.Date)).slice(0, limit);
  } catch (error) {
    console.error('Error getting transactions:', error);
    return [];
  }
}

// Optimized Chart Data
function getChartDataOptimized(workspace, period) {
  try {
    const cache = CacheService.getScriptCache();
    const workspaceId = workspace.getId();
    const cacheKey = `chart_data_${workspaceId}_${period}`;
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
    
    const sheet = workspace.getSheetByName('Financial_Data');
    if (!sheet) return { revenueVsExpenses: { labels: [], revenue: [], expenses: [] }, expenseBreakdown: { labels: [], values: [] } };
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { revenueVsExpenses: { labels: [], revenue: [], expenses: [] }, expenseBreakdown: { labels: [], values: [] } };
    
    const startRow = Math.max(1, lastRow - 2000);
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn()).getValues();
    const headers = data[0];
    
    const dateCol = headers.indexOf('Date');
    const typeCol = headers.indexOf('Type');
    const amountCol = headers.indexOf('Amount');
    const categoryCol = headers.indexOf('Category');
    
    const now = new Date();
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(now.getDate() - 30);
    thirtyDaysAgo.setHours(0, 0, 0, 0);
    
    const last30Days = [];
    for (let i = 29; i >= 0; i--) {
      const date = new Date(now);
      date.setDate(now.getDate() - i);
      last30Days.push(date.toISOString().split('T')[0]);
    }
    
    const revenueByDay = {};
    const expensesByDay = {};
    const expenseByCategory = {};
    
    for (let i = 1; i < data.length; i++) {
      const rowDate = new Date(data[i][dateCol]);
      if (rowDate < thirtyDaysAgo) continue;
      
      const dateStr = rowDate.toISOString().split('T')[0];
      const amount = parseFloat(data[i][amountCol]) || 0;
      const type = data[i][typeCol];
      const category = data[i][categoryCol] || 'other';
      
      if (type === 'revenue') {
        revenueByDay[dateStr] = (revenueByDay[dateStr] || 0) + amount;
      } else if (type === 'expense') {
        const absAmount = Math.abs(amount);
        expensesByDay[dateStr] = (expensesByDay[dateStr] || 0) + absAmount;
        expenseByCategory[category] = (expenseByCategory[category] || 0) + absAmount;
      }
    }
    
    const result = {
      revenueVsExpenses: {
        labels: last30Days,
        revenue: last30Days.map(d => revenueByDay[d] || 0),
        expenses: last30Days.map(d => expensesByDay[d] || 0)
      },
      expenseBreakdown: {
        labels: Object.keys(expenseByCategory),
        values: Object.values(expenseByCategory)
      }
    };
    
    cache.put(cacheKey, JSON.stringify(result), 300);
    return result;
  } catch(e) {
    console.error("getChartDataOptimized error:", e);
    return { revenueVsExpenses: { labels: [], revenue: [], expenses: [] }, expenseBreakdown: { labels: [], values: [] } };
  }
}


/**
 * ==================== DASHBOARD SUPPORT HELPERS ====================
 * These helpers are intentionally kept in Dashboard.js because the
 * enhanced dashboard depends on them at runtime.
 */

function dashboardPeriodBounds_(period, now) {
  const current = now ? new Date(now) : new Date();
  current.setHours(0, 0, 0, 0);

  let start = new Date(current);
  let end = new Date(current);
  let previousStart = new Date(current);
  let previousEnd = new Date(current);

  switch (period) {
    case 'today':
      end.setDate(end.getDate() + 1);
      previousStart.setDate(previousStart.getDate() - 1);
      previousEnd = new Date(current);
      break;

    case 'this_week': {
      const day = current.getDay();
      start.setDate(start.getDate() - day);
      end = new Date(start);
      end.setDate(end.getDate() + 7);
      previousEnd = new Date(start);
      previousStart = new Date(start);
      previousStart.setDate(previousStart.getDate() - 7);
      break;
    }

    case 'this_year':
      start = new Date(current.getFullYear(), 0, 1);
      end = new Date(current.getFullYear() + 1, 0, 1);
      previousStart = new Date(current.getFullYear() - 1, 0, 1);
      previousEnd = new Date(current.getFullYear(), 0, 1);
      break;

    case 'last_month':
      start = new Date(current.getFullYear(), current.getMonth() - 1, 1);
      end = new Date(current.getFullYear(), current.getMonth(), 1);
      previousStart = new Date(current.getFullYear(), current.getMonth() - 2, 1);
      previousEnd = new Date(current.getFullYear(), current.getMonth() - 1, 1);
      break;

    case 'this_month':
    default:
      start = new Date(current.getFullYear(), current.getMonth(), 1);
      end = new Date(current.getFullYear(), current.getMonth() + 1, 1);
      previousStart = new Date(current.getFullYear(), current.getMonth() - 1, 1);
      previousEnd = new Date(current.getFullYear(), current.getMonth(), 1);
      break;
  }

  return { start, end, previousStart, previousEnd };
}

function dashboardReadFinancialPeriod_(sheet, start, end) {
  const result = { revenue: 0, expenses: 0, count: 0, categories: {} };
  if (!sheet || sheet.getLastRow() <= 1) return result;

  const data = sheet.getDataRange().getValues();
  if (!data.length) return result;

  const headers = data[0];
  const dateCol = headers.indexOf('Date');
  const typeCol = headers.indexOf('Type');
  const amountCol = headers.indexOf('Amount');
  const categoryCol = headers.indexOf('Category');

  if (dateCol < 0 || typeCol < 0 || amountCol < 0) return result;

  for (let i = 1; i < data.length; i++) {
    const rawDate = data[i][dateCol];
    const rowDate = rawDate instanceof Date ? new Date(rawDate) : new Date(rawDate);
    if (isNaN(rowDate.getTime()) || rowDate < start || rowDate >= end) continue;

    const amount = Math.abs(parseFloat(data[i][amountCol]) || 0);
    const type = String(data[i][typeCol] || '').toLowerCase();
    result.count++;

    if (type === 'revenue') {
      result.revenue += amount;
    } else if (type === 'expense') {
      result.expenses += amount;
      const category = String(data[i][categoryCol] || 'other');
      result.categories[category] = (result.categories[category] || 0) + amount;
    }
  }

  return result;
}

function calculatePercentageChange(current, previous) {
  const currentValue = Number(current) || 0;
  const previousValue = Number(previous) || 0;
  if (previousValue === 0) return currentValue === 0 ? 0 : 100;
  return ((currentValue - previousValue) / Math.abs(previousValue)) * 100;
}

function getDailyTrends(workspace, period) {
  const sheet = workspace && workspace.getSheetByName('Financial_Data');
  const bounds = dashboardPeriodBounds_(period, new Date());
  const now = new Date();
  const start = new Date(now);
  start.setHours(0, 0, 0, 0);
  start.setDate(start.getDate() - 29);

  const labels = [];
  const revenue = {};
  const expenses = {};

  for (let i = 0; i < 30; i++) {
    const d = new Date(start);
    d.setDate(start.getDate() + i);
    const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    labels.push(key);
    revenue[key] = 0;
    expenses[key] = 0;
  }

  if (sheet && sheet.getLastRow() > 1) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const dateCol = headers.indexOf('Date');
    const typeCol = headers.indexOf('Type');
    const amountCol = headers.indexOf('Amount');

    if (dateCol >= 0 && typeCol >= 0 && amountCol >= 0) {
      for (let i = 1; i < data.length; i++) {
        const rowDate = data[i][dateCol] instanceof Date ? new Date(data[i][dateCol]) : new Date(data[i][dateCol]);
        if (isNaN(rowDate.getTime()) || rowDate < start || rowDate > now) continue;

        const key = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const amount = Math.abs(parseFloat(data[i][amountCol]) || 0);
        const type = String(data[i][typeCol] || '').toLowerCase();

        if (type === 'revenue' && revenue[key] !== undefined) revenue[key] += amount;
        if (type === 'expense' && expenses[key] !== undefined) expenses[key] += amount;
      }
    }
  }

  return {
    labels: labels,
    revenue: labels.map(key => revenue[key]),
    expenses: labels.map(key => expenses[key])
  };
}

function getWeeklyTrends(workspace, period) {
  const daily = getDailyTrends(workspace, period);
  const weekly = {};

  daily.labels.forEach((label, i) => {
    const date = new Date(label + 'T00:00:00');
    const weekStart = new Date(date);
    weekStart.setDate(date.getDate() - date.getDay());
    const key = Utilities.formatDate(weekStart, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    if (!weekly[key]) weekly[key] = { revenue: 0, expenses: 0 };
    weekly[key].revenue += daily.revenue[i] || 0;
    weekly[key].expenses += daily.expenses[i] || 0;
  });

  return weekly;
}

function getPreviousPeriodMetricsOptimized(workspace, period) {
  const sheet = workspace && workspace.getSheetByName('Financial_Data');
  const bounds = dashboardPeriodBounds_(period, new Date());
  return dashboardReadFinancialPeriod_(sheet, bounds.previousStart, bounds.previousEnd);
}

function getDetailedFinancialMetrics(workspace, period, useAccurate) {
  const sheet = workspace && workspace.getSheetByName('Financial_Data');
  const bounds = dashboardPeriodBounds_(period || 'this_month', new Date());
  const current = dashboardReadFinancialPeriod_(sheet, bounds.start, bounds.end);
  const previous = getPreviousPeriodMetricsOptimized(workspace, period || 'this_month');

  const currentProfit = current.revenue - current.expenses;
  const previousProfit = previous.revenue - previous.expenses;
  const currentMargin = current.revenue > 0 ? (currentProfit / current.revenue) * 100 : 0;
  const previousMargin = previous.revenue > 0 ? (previousProfit / previous.revenue) * 100 : 0;

  const daily = getDailyTrends(workspace, period);
  const weekly = getWeeklyTrends(workspace, period);

  return {
    kpis: {
      revenue: { value: current.revenue, change: calculatePercentageChange(current.revenue, previous.revenue) },
      expenses: { value: current.expenses, change: calculatePercentageChange(current.expenses, previous.expenses) },
      netProfit: { value: currentProfit, change: calculatePercentageChange(currentProfit, previousProfit) },
      profitMargin: { value: currentMargin, change: calculatePercentageChange(currentMargin, previousMargin) }
    },
    trends: {
      categoryBreakdown: current.categories,
      daily: daily,
      weekly: weekly
    },
    current: current,
    previous: previous,
    accurate: !!useAccurate
  };
}

function calculateBusinessHealth(workspace, financialMetrics, businessSnapshot) {
  const kpis = financialMetrics && financialMetrics.kpis ? financialMetrics.kpis : {};
  const revenue = Number(kpis.revenue?.value) || 0;
  const expenses = Number(kpis.expenses?.value) || 0;
  const margin = Number(kpis.profitMargin?.value) || 0;
  const snapshot = businessSnapshot || {};
  const moduleActivity = Number(snapshot.activeRecords) || 0;
  const activeModules = Number(snapshot.activeModules) || 0;
  const modulesWithData = Number(snapshot.modulesWithData) || 0;

  const expenseScore = revenue > 0
    ? Math.max(0, Math.min(100, (1 - (expenses / revenue)) * 100))
    : (expenses === 0 ? 100 : 0);
  const marginScore = Math.max(0, Math.min(100, margin * 5));
  const activityScore = activeModules > 0
    ? Math.min(100, (modulesWithData / activeModules) * 100)
    : (moduleActivity > 0 ? 50 : 0);

  const overall = Math.round((Math.max(0, expenseScore) + marginScore + activityScore) / 3);
  const metrics = [
    { name: 'Revenue', current: formatDashboardHealthValue_(revenue), score: revenue > 0 ? 100 : 0, target: 'Positive revenue' },
    { name: 'Expense control', current: revenue > 0 ? ((expenses / revenue) * 100).toFixed(1) + '% of revenue' : 'No revenue', score: Math.round(expenseScore), target: 'Lower expense ratio' },
    { name: 'Profit margin', current: margin.toFixed(1) + '%', score: Math.round(marginScore), target: '20%+' },
    { name: 'Business activity', current: moduleActivity.toLocaleString() + ' records across ' + modulesWithData + ' active modules', score: Math.round(activityScore), target: 'Use the modules relevant to your business' }
  ];

  const recommendations = [];
  if (revenue === 0) {
    recommendations.push({ priority: 'high', message: 'No revenue recorded for this period.', action: 'Review finance, ecommerce, sales and POS activity.' });
  }
  if (expenses > revenue && expenses > 0) {
    recommendations.push({ priority: 'high', message: 'Expenses exceed recorded revenue for this period.', action: 'Review recent expenses and revenue entries.' });
  }
  if (margin > 0 && margin < 10) {
    recommendations.push({ priority: 'medium', message: 'Profit margin is below 10%.', action: 'Review pricing, costs and operating expenses.' });
  }
  if (activeModules === 0) {
    recommendations.push({ priority: 'medium', message: 'No operational modules contain records yet.', action: 'Start with the modules that match your business workflow.' });
  }

  return { overall: Math.max(0, Math.min(100, overall)), metrics, recommendations };
}
function formatDashboardHealthValue_(value) {
  const n = Number(value) || 0;
  try {
    return typeof formatCurrency === 'function' ? formatCurrency(n) : n.toFixed(2);
  } catch (e) {
    return n.toFixed(2);
  }
}

function getDashboardBusinessSnapshot(workspace, moduleSummary) {
  const modules = {
    Finance: 'Financial_Data',
    Ecommerce: 'Ecommerce_Data',
    Sales: 'Sales_Data',
    CRM: 'CRM_Data',
    HR: 'HR_Data',
    Logistics: 'Logistics_Data',
    Tax: 'Tax_Data',
    Agro: 'Agro_Data',
    Productivity: 'Productivity_Data',
    POS: 'POS_Data',
    Attendance: 'Attendance_Data',
    Warehouse: 'Warehouse_Data'
  };

  const snapshot = {
    activeModules: 0,
    modulesWithData: 0,
    activeRecords: 0,
    commerceRecords: 0,
    teamMembers: 0,
    operational: {}
  };

  // Business Center is an Ecommerce-connected workspace area rather than a
  // separate entitlement/module sheet, so include its operational records
  // explicitly without counting them as another module.
  try {
    const businessCenterSales = workspace.getSheetByName('BusinessCenter_Sales');
    const saleCount = businessCenterSales && businessCenterSales.getLastRow() > 1
      ? businessCenterSales.getLastRow() - 1 : 0;
    if (saleCount > 0) snapshot.modulesWithData++;
    snapshot.activeRecords += saleCount;
    snapshot.commerceRecords += saleCount;

    if (businessCenterSales && saleCount > 0) {
      const values = businessCenterSales.getDataRange().getValues();
      const headers = values[0] || [];
      const totalCol = headers.indexOf('Total');
      const amountPaidCol = headers.indexOf('Amount_Paid');
      const balanceCol = headers.indexOf('Balance_Due');
      snapshot.operational.BusinessCenter = {
        sales: saleCount,
        salesValue: totalCol >= 0 ? values.slice(1).reduce((sum, row) => sum + (parseFloat(row[totalCol]) || 0), 0) : 0,
        cashCollected: amountPaidCol >= 0 ? values.slice(1).reduce((sum, row) => sum + (parseFloat(row[amountPaidCol]) || 0), 0) : 0,
        receivables: balanceCol >= 0 ? values.slice(1).reduce((sum, row) => sum + (parseFloat(row[balanceCol]) || 0), 0) : 0
      };
    }
  } catch (businessCenterError) {
    console.warn('Business Center dashboard snapshot skipped:', businessCenterError.message);
  }

  Object.keys(modules).forEach(moduleName => {
    if (moduleSummary && moduleSummary[moduleName] && moduleSummary[moduleName].isAccessible === false) return;
    try {
      const sheet = workspace.getSheetByName(modules[moduleName]);
      const count = sheet && sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 0;
      snapshot.activeModules++;
      if (count > 0) snapshot.modulesWithData++;
      snapshot.activeRecords += count;
      if (moduleName === 'Ecommerce' || moduleName === 'POS') snapshot.commerceRecords += count;

      const values = sheet && count > 0 ? sheet.getDataRange().getValues() : [];
      const headers = values.length ? values[0] : [];
      const col = name => headers.indexOf(name);
      const sum = name => {
        const i = col(name);
        if (i < 0) return 0;
        return values.slice(1).reduce((total, row) => total + (parseFloat(row[i]) || 0), 0);
      };

      if (moduleName === 'Ecommerce') {
        snapshot.operational.Ecommerce = { orders: count, orderValue: sum('Total_Amount') };
      } else if (moduleName === 'POS') {
        snapshot.operational.POS = { transactions: count, sales: sum('Total_Amount'), items: sum('Quantity') };
      } else if (moduleName === 'Sales') {
        snapshot.operational.Sales = { deals: count, pipelineValue: sum('Total_Amount') };
      } else if (moduleName === 'CRM') {
        snapshot.operational.CRM = { contacts: count };
      } else if (moduleName === 'HR') {
        snapshot.operational.HR = { employees: count, payroll: sum('Salary') };
        snapshot.teamMembers = count;
      } else if (moduleName === 'Logistics') {
        snapshot.operational.Logistics = { shipments: count };
      } else if (moduleName === 'Tax') {
        snapshot.operational.Tax = { records: count, taxAmount: sum('Tax_Amount') };
      } else if (moduleName === 'Agro') {
        snapshot.operational.Agro = { records: count, revenue: sum('Revenue'), costs: sum('Total_Cost') };
      } else if (moduleName === 'Productivity') {
        snapshot.operational.Productivity = { tasks: count, hours: sum('Hours_Spent') };
      } else if (moduleName === 'Attendance') {
        snapshot.operational.Attendance = { records: count };
      } else if (moduleName === 'Warehouse') {
        let lowStock = 0;
        const q = col('Quantity');
        const min = col('Min_Stock');
        if (q >= 0 && min >= 0) {
          values.slice(1).forEach(row => {
            const quantity = parseFloat(row[q]) || 0;
            const minimum = parseFloat(row[min]);
            if (!isNaN(minimum) && quantity < minimum) lowStock++;
          });
        }
        snapshot.operational.Warehouse = {
          products: count,
          stockUnits: sum('Quantity'),
          inventoryValue: values.slice(1).reduce((total, row) => {
            const quantity = q >= 0 ? (parseFloat(row[q]) || 0) : 0;
            const costCol = col('Unit_Cost');
            const unitCost = costCol >= 0 ? (parseFloat(row[costCol]) || 0) : 0;
            return total + (quantity * unitCost);
          }, 0),
          lowStock: lowStock
        };
      } else if (moduleName === 'Finance') {
        snapshot.operational.Finance = { records: count };
      }
    } catch (e) {
      console.log('Dashboard snapshot skipped for ' + moduleName + ': ' + e.message);
    }
  });

  return snapshot;
}

function getDashboardRecentEntries(moduleName, limit, sessionId) {
  const allowed = ['Finance','Ecommerce','Sales','HR'];
  if (typeof limit === 'string' && limit.length > 30 && sessionId === undefined) {
    sessionId = limit;
    limit = 5;
  }
  if (!allowed.includes(moduleName)) return [];

  const user = safeGetUser(sessionId);
  if (!user) return [];

  try {
    const workspace = getWorkspaceFile(sessionId);
    const sheetMap = {
      Finance: 'Financial_Data',
      Ecommerce: 'Ecommerce_Data',
      Sales: 'Sales_Data',
      HR: 'HR_Data'
    };
    const sheet = workspace.getSheetByName(sheetMap[moduleName]);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const rows = values.slice(1);
    const value = (row, name) => {
      const i = headers.indexOf(name);
      return i >= 0 ? row[i] : '';
    };
    const toDate = raw => {
      const d = raw instanceof Date ? new Date(raw) : new Date(raw);
      return isNaN(d.getTime()) ? null : d;
    };

    return rows.map(row => {
      const date = toDate(value(row, 'Date')) || toDate(value(row, 'Created_At'));
      const amount = moduleName === 'Finance'
        ? (parseFloat(value(row, 'Amount')) || 0)
        : (parseFloat(value(row, 'Total_Amount')) || 0);
      let title = '';
      let status = '';
      let id = '';

      if (moduleName === 'Finance') {
        title = value(row, 'Description') || value(row, 'Category') || 'Financial transaction';
        status = value(row, 'Type') || '';
        id = value(row, 'Transaction_ID') || '';
      } else if (moduleName === 'Ecommerce') {
        title = 'Order ' + (value(row, 'Order_ID') || 'record') + (value(row, 'Product_Name') ? ': ' + value(row, 'Product_Name') : '');
        status = value(row, 'Payment_Status') || value(row, 'Order_Status') || '';
        id = value(row, 'Order_ID') || '';
      } else if (moduleName === 'Sales') {
        title = (value(row, 'Customer_Name') || value(row, 'Product_Service') || 'Sales deal');
        status = value(row, 'Stage') || '';
        id = value(row, 'Deal_ID') || '';
      } else if (moduleName === 'HR') {
        title = value(row, 'Name') || 'Employee';
        status = value(row, 'Status') || '';
        id = value(row, 'Employee_ID') || '';
      }

      return {
        module: moduleName,
        date: date ? date.toISOString() : '',
        title: String(title),
        status: String(status),
        id: String(id),
        amount: amount
      };
    }).filter(item => item.date).sort((a,b) => new Date(b.date) - new Date(a.date)).slice(0, Number(limit) || 5);
  } catch (e) {
    console.error('Dashboard recent entries error:', e);
    return [];
  }
}

function getRecentActivityAcrossModules(workspace, limit, moduleSummary) {
  const activities = [];
  const allModules = {
    Finance: 'Financial_Data',
    Ecommerce: 'Ecommerce_Data',
    Sales: 'Sales_Data',
    CRM: 'CRM_Data',
    HR: 'HR_Data',
    Logistics: 'Logistics_Data',
    Tax: 'Tax_Data',
    Agro: 'Agro_Data',
    Productivity: 'Productivity_Data',
    POS: 'POS_Data',
    Attendance: 'Attendance_Data',
    Warehouse: 'Warehouse_Data'
  };

  Object.keys(allModules).forEach(moduleName => {
    if (moduleSummary && moduleSummary[moduleName] && moduleSummary[moduleName].isAccessible === false) return;

    try {
      const sheet = workspace.getSheetByName(allModules[moduleName]);
      if (!sheet || sheet.getLastRow() <= 1) return;

      const values = sheet.getDataRange().getValues();
      const headers = values[0];
      const dateCol = headers.indexOf('Date');
      const createdAtCol = headers.indexOf('Created_At');
      const idCandidates = ['Transaction_ID','Order_ID','Deal_ID','Contact_ID','Employee_ID','Shipment_ID','Tax_ID','Record_ID','Item_ID'];

      for (let i = Math.max(1, values.length - 25); i < values.length; i++) {
        const row = values[i];
        const rawDate = dateCol >= 0 && row[dateCol] ? row[dateCol] : (createdAtCol >= 0 ? row[createdAtCol] : null);
        const date = rawDate instanceof Date ? new Date(rawDate) : new Date(rawDate);
        if (isNaN(date.getTime())) continue;

        let recordId = '';
        for (const id of idCandidates) {
          const idx = headers.indexOf(id);
          if (idx >= 0 && row[idx]) {
            recordId = String(row[idx]);
            break;
          }
        }

        const description = dashboardActivityDescription_(moduleName, row, headers, recordId);
        const daysAgo = Math.max(0, Math.floor((new Date().getTime() - date.getTime()) / 86400000));
        activities.push({ module: moduleName, description, date: date.toISOString(), daysAgo });
      }
    } catch (e) {
      console.log('Recent activity skipped for ' + moduleName + ': ' + e.message);
    }
  });

  return activities.sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, limit || 10);
}

function dashboardActivityDescription_(moduleName, row, headers, recordId) {
  const value = name => {
    const idx = headers.indexOf(name);
    return idx >= 0 ? row[idx] : '';
  };

  switch (moduleName) {
    case 'Finance': return (value('Description') || 'Financial transaction') + (recordId ? ' (' + recordId + ')' : '');
    case 'Ecommerce': return 'Order ' + (value('Order_ID') || recordId || 'record') + (value('Product_Name') ? ': ' + value('Product_Name') : '');
    case 'Sales': return 'Deal ' + (value('Deal_ID') || recordId || 'record') + (value('Customer_Name') ? ': ' + value('Customer_Name') : '');
    case 'CRM': return 'Contact ' + (value('Name') || recordId || 'record') + ' updated';
    case 'HR': return 'Employee ' + (value('Name') || recordId || 'record') + ' updated';
    case 'Logistics': return 'Shipment ' + (value('Shipment_ID') || recordId || 'record') + ' updated';
    case 'Tax': return 'Tax record ' + (value('Tax_ID') || recordId || 'record') + ' updated';
    case 'Agro': return 'Agro record ' + (value('Record_ID') || recordId || 'record') + ' updated';
    case 'Productivity': return 'Task ' + (value('Record_ID') || recordId || 'record') + ' logged';
    case 'POS': return 'POS transaction ' + (value('Transaction_ID') || recordId || 'record') + ' recorded';
    case 'Attendance': return 'Attendance ' + (value('Record_ID') || recordId || 'record') + ' recorded';
    case 'Warehouse': return 'Inventory item ' + (value('Item_Name') || value('Item_ID') || recordId || 'record') + ' updated';
    default: return moduleName + ' activity';
  }
}

function generateBusinessAlerts(workspace, financialMetrics, moduleSummary, businessSnapshot) {
  const alerts = [];
  const kpis = financialMetrics?.kpis || {};
  const revenue = Number(kpis.revenue?.value) || 0;
  const expenses = Number(kpis.expenses?.value) || 0;

  if (revenue === 0 && expenses === 0) {
    alerts.push({ type: 'info', priority: 'low', title: 'No financial activity', message: 'No revenue or expenses were recorded for this period.' });
  } else if (expenses > revenue) {
    alerts.push({ type: 'warning', priority: 'high', title: 'Expenses exceed revenue', message: 'Recorded expenses are higher than recorded revenue for this period.' });
  }

  if (!moduleSummary || moduleSummary.Warehouse?.isAccessible !== false) {
    const sheet = workspace.getSheetByName('Warehouse_Data');
    if (sheet && sheet.getLastRow() > 1) {
      const values = sheet.getDataRange().getValues();
      const headers = values[0];
      const quantityCol = headers.indexOf('Quantity');
      const minCol = headers.indexOf('Min_Stock');
      if (quantityCol >= 0 && minCol >= 0) {
        let lowStock = 0;
        for (let i = 1; i < values.length; i++) {
          const qty = parseFloat(values[i][quantityCol]) || 0;
          const min = parseFloat(values[i][minCol]);
          if (!isNaN(min) && qty < min) lowStock++;
        }
        if (lowStock > 0) {
          alerts.push({ type: 'warning', priority: 'high', title: 'Low stock', message: lowStock + ' warehouse item(s) are below minimum stock.' });
        }
      }
    }
  }

  if (!moduleSummary || moduleSummary.Ecommerce?.isAccessible !== false) {
    const sheet = workspace.getSheetByName('Ecommerce_Data');
    if (sheet && sheet.getLastRow() > 1) {
      const values = sheet.getDataRange().getValues();
      const headers = values[0];
      const statusCol = headers.indexOf('Payment_Status');
      if (statusCol >= 0) {
        let pending = 0;
        for (let i = 1; i < values.length; i++) {
          if (String(values[i][statusCol] || '').toLowerCase() === 'pending') pending++;
        }
        if (pending > 0) {
          alerts.push({ type: 'info', priority: 'medium', title: 'Pending ecommerce payments', message: pending + ' ecommerce order(s) have a pending payment status.' });
        }
      }
    }
  }

  const snapshot = businessSnapshot || {};
  const pos = snapshot.operational?.POS;
  if (pos && pos.transactions > 0) {
    alerts.push({ type: 'info', priority: 'low', title: 'POS activity', message: pos.transactions + ' POS transaction(s) recorded, totaling ' + formatDashboardHealthValue_(pos.sales) + '.' });
  }

  const attendance = snapshot.operational?.Attendance;
  if (attendance && attendance.records > 0) {
    alerts.push({ type: 'info', priority: 'low', title: 'Staff attendance activity', message: attendance.records + ' attendance record(s) are available.' });
  }

  const logistics = snapshot.operational?.Logistics;
  if (logistics && logistics.shipments > 0) {
    alerts.push({ type: 'info', priority: 'low', title: 'Logistics activity', message: logistics.shipments + ' shipment record(s) are being tracked.' });
  }

  return alerts.slice(0, 8);
}

function recalculateDashboardMetrics(workspaceId) {
  // Dashboard values are calculated from the workspace on demand.
  // The existing short-lived caches expire automatically; this function
  // remains as the post-sync hook used by module auto-sync.
  return { success: true, workspaceId: workspaceId || null };
}

