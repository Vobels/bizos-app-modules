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
    const businessHealth = calculateBusinessHealth(workspace, financialMetrics);
    const chartData = getChartDataOptimized(workspace, period);
    const recentActivity = getRecentActivityAcrossModules(workspace, 10);
    const alerts = generateBusinessAlerts(workspace, financialMetrics);
    
    return {
      success: true,
      dashboard: {
        kpis: financialMetrics.kpis,
        trends: financialMetrics.trends,
        moduleSummary: moduleSummary,
        health: businessHealth,
        charts: chartData,
        recentActivity: recentActivity,
        alerts: alerts
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

// Keep your existing helper functions (calculatePercentageChange, getDailyTrends, getWeeklyTrends, 
// getRecentActivityAcrossModules, generateBusinessAlerts, calculateBusinessHealth, 
// getDetailedFinancialMetrics, getPreviousPeriodMetricsOptimized, recalculateDashboardMetrics)
// They should already exist in your file - keep them as is