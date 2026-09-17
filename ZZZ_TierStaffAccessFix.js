// ============================================================
// BizOS — Tier-based staff module access
// ============================================================
// Staff module entitlement is determined by the business tier.
// Staff_Modules is legacy and is intentionally NOT authoritative.
// Owner/admin permissions remain role-based.
// ============================================================

function getStaffModules(email, businessId) {
  try {
    const businessSheet = getOrCreateBusinessSheet();
    const data = businessSheet.getDataRange().getValues();
    if (data.length < 2) return ['Finance', 'Ecommerce'];

    const headers = data[0];
    const bizIdCol = headers.indexOf('Business_ID');
    const tierCol = headers.indexOf('Subscription_Tier');
    if (bizIdCol === -1) return ['Finance', 'Ecommerce'];

    let tier = 'starter';
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][bizIdCol] || '') === String(businessId || '')) {
        tier = String(tierCol !== -1 ? data[i][tierCol] || 'starter' : 'starter').toLowerCase();
        break;
      }
    }

    return (tier === 'sovereign' || tier === 'enterprise')
      ? ['Finance', 'Ecommerce', 'Sales', 'CRM', 'HR', 'Logistics', 'Tax', 'Agro', 'Productivity', 'POS', 'Attendance', 'Warehouse']
      : ['Finance', 'Ecommerce'];
  } catch (error) {
    console.error('Tier staff access error:', error);
    return ['Finance', 'Ecommerce'];
  }
}
