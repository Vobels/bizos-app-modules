// ============================================================
// ClientDeploymentV6.js - PAID CLIENT DASHBOARD PARITY
// ============================================================
// Paid clients receive the BizOS dashboard/application experience
// against their own isolated workspace. No payment or upgrade UI.
// ============================================================

function generateClientCodeSafelyV6(settings) {
  var base = generateClientCodeSafelyV5(settings);
  if (!base || !base.files) throw new Error('Client deployment package is empty.');

  base.files = base.files.map(function(file) {
    if (file && (file.name === 'client-dashboard' || file.name === 'client-dashboard.html')) {
      file.source = buildClientDashboardHtmlV6_();
      delete file.content;
    }
    return file;
  });

  return base;
}

function buildClientDashboardHtmlV6_() {
  // One outer template literal only. The generated browser JavaScript
  // deliberately uses concatenation instead of nested template literals.
  return `<!doctype html>
<html lang="en">
<head>
<base target="_top">
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title><?= clientName ?> - BizOS</title>
<script src="https://cdn.tailwindcss.com"></script>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<style>
body{font-family:Inter,Arial,sans-serif;background:#f9fafb;margin:0;color:#111827}.sidebar{width:256px}.card{background:#fff;border:1px solid #eaecf0;border-radius:12px;box-shadow:0 1px 3px rgba(16,24,40,.06)}.module-card{transition:.15s}.module-card:hover{transform:translateY(-1px);box-shadow:0 5px 18px rgba(16,24,40,.08)}.mobile-only{display:none}@media(max-width:767px){.sidebar{transform:translateX(-100%);transition:.2s;z-index:40}.sidebar.open{transform:translateX(0)}.mobile-only{display:block}.content{margin-left:0!important}.overlay{display:none}.overlay.show{display:block;position:fixed;inset:0;background:#0006;z-index:30}.desktop-actions{display:none!important}}
</style>
</head>
<body>
<div id="overlay" class="overlay" onclick="closeSidebar()"></div>
<aside id="sidebar" class="sidebar bg-white border-r border-gray-200 fixed top-0 bottom-0 left-0 overflow-y-auto">
<div class="p-5">
<div class="flex items-center gap-3 mb-7"><div class="w-9 h-9 rounded-lg flex items-center justify-center text-white font-bold" style="background:<?= primaryColor ?>"><?= clientName.substring(0,1).toUpperCase() ?></div><div class="font-extrabold text-lg truncate"><?= clientName ?></div></div>
<div class="mb-5 p-3 bg-gray-50 rounded-xl"><div class="font-medium truncate" id="userName">Loading...</div><div class="text-xs text-gray-500 truncate" id="userEmail"></div><div class="text-xs font-semibold mt-2" style="color:<?= primaryColor ?>">Sovereign</div></div>
<nav class="space-y-1">
<button onclick="showHome()" class="w-full text-left px-3 py-2.5 rounded-lg hover:bg-gray-100 flex items-center gap-3"><i class="fas fa-home w-5 text-gray-500"></i>Dashboard</button>
<button onclick="showApps()" class="w-full text-left px-3 py-2.5 rounded-lg hover:bg-gray-100 flex items-center gap-3"><i class="fas fa-th-large w-5 text-gray-500"></i>Apps</button>
<button onclick="showProfile()" class="w-full text-left px-3 py-2.5 rounded-lg hover:bg-gray-100 flex items-center gap-3"><i class="fas fa-user-circle w-5 text-gray-500"></i>Profile</button>
<button onclick="showSettings()" class="w-full text-left px-3 py-2.5 rounded-lg hover:bg-gray-100 flex items-center gap-3"><i class="fas fa-cog w-5 text-gray-500"></i>Settings</button>
<div class="pt-5"><div class="text-xs font-bold uppercase tracking-wider text-gray-400 px-3 mb-2">Applications</div><div id="navModules"></div></div>
</nav></div>
<div class="absolute bottom-0 left-0 right-0 p-5 border-t bg-white"><button onclick="logout()" class="w-full text-left px-3 py-2.5 rounded-lg hover:bg-gray-100 flex items-center gap-3 text-red-600"><i class="fas fa-sign-out-alt w-5"></i>Sign Out</button></div>
</aside>

<main class="content ml-64 min-h-screen">
<header class="bg-white border-b sticky top-0 z-20"><div class="px-4 md:px-6 py-4 flex items-center justify-between"><div class="flex items-center gap-3"><button class="mobile-only text-gray-600" onclick="openSidebar()"><i class="fas fa-bars text-xl"></i></button><div><div class="text-xl font-bold" id="pageTitle">Dashboard</div><div class="text-xs text-gray-500" id="pageSubtitle">Your business at a glance</div></div></div><div class="flex items-center gap-4 desktop-actions"><button class="text-gray-500" title="Refresh" onclick="loadDashboard()"><i class="fas fa-sync-alt"></i></button><button class="text-gray-500" title="Notifications"><i class="fas fa-bell"></i></button></div></div></header>

<section id="home" class="p-4 md:p-6 max-w-7xl mx-auto">
<div class="grid grid-cols-2 lg:grid-cols-4 gap-3 md:gap-5 mb-6"><div class="card p-4 md:p-5"><div class="text-xs text-gray-500">Total Revenue</div><div id="revenue" class="text-xl md:text-3xl font-extrabold mt-2">0</div></div><div class="card p-4 md:p-5"><div class="text-xs text-gray-500">Total Expenses</div><div id="expenses" class="text-xl md:text-3xl font-extrabold mt-2">0</div></div><div class="card p-4 md:p-5"><div class="text-xs text-gray-500">Net Profit</div><div id="profit" class="text-xl md:text-3xl font-extrabold mt-2">0</div></div><div class="card p-4 md:p-5"><div class="text-xs text-gray-500">Profit Margin</div><div id="margin" class="text-xl md:text-3xl font-extrabold mt-2">0%</div></div></div>
<div class="card p-5 mb-6"><div class="flex items-center justify-between mb-4"><div><h2 class="font-bold text-lg">Data Management</h2><p class="text-xs text-gray-500">View, add and manage your business records</p></div><button onclick="showApps()" class="text-sm font-semibold" style="color:<?= primaryColor ?>">View All Apps →</button></div><div id="homeModules" class="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-3"></div></div>
<div class="card p-5 mb-6"><div class="flex items-center justify-between mb-4"><h2 class="font-bold text-lg">Recent Entries</h2><button onclick="openModule('Finance')" class="text-sm font-semibold" style="color:<?= primaryColor ?>">View All →</button></div><div id="recent" class="text-sm text-gray-500">Loading...</div></div>
<div class="card p-5"><div class="mb-4"><h2 class="font-bold text-lg">Featured Apps</h2><p class="text-xs text-gray-500">Quick access to your business tools</p></div><div id="featured" class="grid grid-cols-2 md:grid-cols-4 gap-3"></div></div>
</section>

<section id="apps" class="hidden p-4 md:p-6 max-w-7xl mx-auto"><h2 class="text-2xl font-bold mb-2">Business Applications</h2><p class="text-sm text-gray-500 mb-6">Your enabled BizOS modules, customized for <?= clientName ?>.</p><div id="appsGrid" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4"></div></section>
<section id="profile" class="hidden p-4 md:p-6 max-w-3xl mx-auto"><div class="card p-6"><h2 class="text-2xl font-bold mb-5">Profile</h2><div id="profileBody" class="text-sm"></div></div></section>
<section id="settings" class="hidden p-4 md:p-6 max-w-3xl mx-auto"><div class="card p-6"><h2 class="text-2xl font-bold mb-2">Settings</h2><p class="text-sm text-gray-500 mb-5">Workspace and account settings.</p><div class="p-4 bg-gray-50 rounded-xl"><div class="font-semibold">Workspace</div><div class="text-sm text-gray-500 mt-1"><?= clientName ?></div><div class="text-xs text-gray-400 mt-2">Your BizOS workspace is managed by Vobels.</div></div></div></section>
<section id="module" class="hidden p-4 md:p-6 max-w-7xl mx-auto"><div class="flex items-center justify-between mb-5"><div><button onclick="showHome()" class="text-sm text-gray-500">← Dashboard</button><h2 id="moduleTitle" class="text-2xl font-bold mt-2"></h2></div><button onclick="openAdd()" class="px-4 py-2 rounded-lg text-white font-semibold" style="background:<?= primaryColor ?>"><i class="fas fa-plus mr-2"></i>Add Record</button></div><div class="card overflow-hidden"><div id="moduleTable" class="overflow-x-auto p-3">Loading...</div></div></section>
<div id="modal" class="hidden fixed inset-0 bg-black/50 z-50 items-center justify-center p-4"><div class="bg-white rounded-xl w-full max-w-2xl max-h-[90vh] overflow-y-auto p-5"><div class="flex justify-between items-center mb-4"><h3 id="modalTitle" class="font-bold text-lg">Add Record</h3><button onclick="closeModal()">×</button></div><div id="form" class="grid grid-cols-1 md:grid-cols-2 gap-3"></div><div class="flex justify-end gap-2 mt-5"><button onclick="closeModal()" class="px-4 py-2 rounded-lg border">Cancel</button><button onclick="saveRecord()" class="px-4 py-2 rounded-lg text-white font-semibold" style="background:<?= primaryColor ?>">Save</button></div></div></div>

<script>
const SID=<?= sessionId ? JSON.stringify(sessionId) : 'null' ?>;
const PRIMARY='<?= primaryColor ?>';
const CLIENT='<?= clientName ?>';
const MODULES=['Finance','Ecommerce','Sales','CRM','HR','Logistics','Tax','Agro','Productivity','POS','Attendance','Warehouse'];
const LABELS={Finance:'Financial Management',Ecommerce:'E-commerce',Sales:'Sales Pipeline',CRM:'Customer Relations',HR:'Human Resources',Logistics:'Logistics',Tax:'Tax Compliance',Agro:'Agriculture',Productivity:'Productivity Suite',POS:'Point of Sale',Attendance:'Staff Attendance',Warehouse:'Warehouse Management'};
const ICONS={Finance:'fa-coins',Ecommerce:'fa-shopping-cart',Sales:'fa-chart-line',CRM:'fa-users',HR:'fa-user-tie',Logistics:'fa-truck',Tax:'fa-file-invoice-dollar',Agro:'fa-seedling',Productivity:'fa-list-check',POS:'fa-cash-register',Attendance:'fa-calendar-check',Warehouse:'fa-warehouse'};
let currentModule='',currentRecords=[],editingRow=null;
function run(fn){return google.script.run.withSuccessHandler(fn).withFailureHandler(function(e){alert(e&&e.message?e.message:'Request failed');});}
function hideAll(){['home','apps','profile','settings','module'].forEach(function(id){document.getElementById(id).classList.add('hidden');});}
function showHome(){hideAll();document.getElementById('home').classList.remove('hidden');document.getElementById('pageTitle').textContent='Dashboard';document.getElementById('pageSubtitle').textContent='Your business at a glance';closeSidebar();loadDashboard();}
function showApps(){hideAll();document.getElementById('apps').classList.remove('hidden');document.getElementById('pageTitle').textContent='Apps';document.getElementById('pageSubtitle').textContent='Business applications';renderApps();closeSidebar();}
function showProfile(){hideAll();document.getElementById('profile').classList.remove('hidden');document.getElementById('pageTitle').textContent='Profile';document.getElementById('pageSubtitle').textContent='Account information';var u=JSON.parse(localStorage.getItem('bizos_client_user')||'{}');document.getElementById('profileBody').innerHTML='<b>'+escapeHtml(CLIENT)+'</b><br><span class="text-gray-500">'+escapeHtml(u.email||'')+'</span>';closeSidebar();}
function showSettings(){hideAll();document.getElementById('settings').classList.remove('hidden');document.getElementById('pageTitle').textContent='Settings';document.getElementById('pageSubtitle').textContent='Workspace settings';closeSidebar();}
function openSidebar(){document.getElementById('sidebar').classList.add('open');document.getElementById('overlay').classList.add('show');}
function closeSidebar(){document.getElementById('sidebar').classList.remove('open');document.getElementById('overlay').classList.remove('show');}
function money(v){return Number(v||0).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});}
function loadDashboard(){if(!SID){location.href='?page=login';return;}run(function(r){if(!r||!r.success){showLoadError(r&&r.message);return;}var d=r.dashboard||{},k=d.kpis||{};document.getElementById('revenue').textContent=money(k.revenue);document.getElementById('expenses').textContent=money(k.expenses);document.getElementById('profit').textContent=money(k.profit);document.getElementById('margin').textContent=(k.margin||0)+'%';renderModules(d.modules||[]);renderRecent(d.recentTransactions||[]);}).getClientDashboardData(SID);}
function renderModules(data){var all=MODULES.map(function(n){var x=data.find(function(m){return m.name===n;})||{};return{name:n,label:LABELS[n],count:x.count||0;};});var home='',nav='',featured='';all.forEach(function(x){home+='<button onclick="openModule(\''+x.name+'\')" class="module-card card p-4 text-left"><div class="flex items-center gap-3"><div class="w-10 h-10 rounded-lg flex items-center justify-center text-white" style="background:'+PRIMARY+'"><i class="fas '+ICONS[x.name]+'"></i></div><div class="min-w-0"><div class="font-semibold truncate">'+escapeHtml(x.label)+'</div><div class="text-xs text-gray-500">'+x.count+' records</div></div></div></button>';nav+='<button onclick="openModule(\''+x.name+'\')" class="w-full text-left px-3 py-2 rounded-lg hover:bg-gray-100 text-sm flex justify-between"><span><i class="fas '+ICONS[x.name]+' w-5 text-gray-400"></i>'+escapeHtml(x.label)+'</span><span class="text-xs text-gray-400">'+x.count+'</span></button>';});all.slice(0,4).forEach(function(x){featured+='<button onclick="openModule(\''+x.name+'\')" class="border rounded-xl p-4 text-left hover:bg-gray-50"><i class="fas '+ICONS[x.name]+' text-lg" style="color:'+PRIMARY+'"></i><div class="font-semibold mt-2">'+escapeHtml(x.label)+'</div><div class="text-xs text-gray-500">Open app</div></button>';});document.getElementById('homeModules').innerHTML=home;document.getElementById('navModules').innerHTML=nav;document.getElementById('featured').innerHTML=featured;}
function renderApps(){var html='';MODULES.forEach(function(n){html+='<button onclick="openModule(\''+n+'\')" class="module-card card p-5 text-left"><div class="w-11 h-11 rounded-xl flex items-center justify-center text-white" style="background:'+PRIMARY+'"><i class="fas '+ICONS[n]+'"></i></div><h3 class="font-bold mt-3">'+escapeHtml(LABELS[n])+'</h3><p class="text-xs text-gray-500 mt-1">Manage '+escapeHtml(LABELS[n].toLowerCase())+' records.</p></button>';});document.getElementById('appsGrid').innerHTML=html;}
function renderRecent(rows){if(!rows.length){document.getElementById('recent').innerHTML='<div class="py-8 text-center text-gray-400">No financial records yet. Add your first transaction.</div>';return;}var html='<div class="divide-y">';rows.slice(0,8).forEach(function(r){html+='<div class="py-3 flex justify-between gap-4"><div><div class="font-medium">'+escapeHtml(r.description||'Financial entry')+'</div><div class="text-xs text-gray-500">'+escapeHtml(r.date||'')+'</div></div><div class="font-semibold">'+money(r.amount)+'</div></div>';});document.getElementById('recent').innerHTML=html+'</div>';}
function openModule(name){if(MODULES.indexOf(name)<0)return;currentModule=name;hideAll();document.getElementById('module').classList.remove('hidden');document.getElementById('moduleTitle').textContent=LABELS[name];document.getElementById('pageTitle').textContent=LABELS[name];document.getElementById('pageSubtitle').textContent='Manage your records';closeSidebar();loadModule();}
function loadModule(){document.getElementById('moduleTable').textContent='Loading...';run(function(r){if(!r||!r.success){document.getElementById('moduleTable').innerHTML='<div class="p-6 text-red-600">'+escapeHtml(r&&r.message||'Failed to load module.')+'</div>';return;}currentRecords=r.records||[];var h=r.headers||[];if(!h.length){document.getElementById('moduleTable').innerHTML='<div class="p-6 text-gray-400">No records yet. Click Add Record to create one.</div>';return;}var html='<table class="min-w-full text-sm"><thead><tr>';h.forEach(function(x){html+='<th class="text-left p-3 border-b whitespace-nowrap">'+escapeHtml(x)+'</th>';});html+='<th class="p-3 border-b">Actions</th></tr></thead><tbody>';currentRecords.forEach(function(row,i){html+='<tr>';h.forEach(function(x){html+='<td class="p-3 border-b whitespace-nowrap">'+escapeHtml(row[x])+'</td>';});html+='<td class="p-3 border-b whitespace-nowrap"><button class="text-blue-600 mr-3" onclick="editRecord('+i+')">Edit</button><button class="text-red-600" onclick="deleteRecord('+i+')">Delete</button></td></tr>';});document.getElementById('moduleTable').innerHTML=html+'</tbody></table>';}).getClientModuleSchema(currentModule,SID);}
function openAdd(){editingRow=null;document.getElementById('modalTitle').textContent='Add Record';buildForm({});document.getElementById('modal').classList.remove('hidden');document.getElementById('modal').style.display='flex';}
function editRecord(i){if(!currentRecords[i])return;editingRow=currentRecords[i]._row;document.getElementById('modalTitle').textContent='Edit Record';buildForm(currentRecords[i]);document.getElementById('modal').classList.remove('hidden');document.getElementById('modal').style.display='flex';}
function buildForm(record){var keys=Object.keys(record).filter(function(k){return k!=='_row';});if(!keys.length){document.getElementById('form').innerHTML='<div class="text-gray-500">This module has no fields yet.</div>';return;}var html='';keys.forEach(function(k){html+='<label class="text-sm"><span class="block font-medium mb-1">'+escapeHtml(k)+'</span><input data-key="'+escapeHtml(k)+'" value="'+escapeHtml(record[k]||'')+'" class="w-full border rounded-lg px-3 py-2"></label>';});document.getElementById('form').innerHTML=html;}
function saveRecord(){var record={};document.querySelectorAll('#form [data-key]').forEach(function(i){record[i.getAttribute('data-key')]=i.value;});var done=function(r){if(r&&r.success){closeModal();loadModule();loadDashboard();}else alert(r&&r.message||'Could not save record.');};if(editingRow)run(done).updateClientModuleRecord(currentModule,editingRow,record,SID);else run(done).createClientModuleRecord(currentModule,record,SID);}
function deleteRecord(i){if(!currentRecords[i]||!confirm('Delete this record?'))return;run(function(r){if(r&&r.success){loadModule();loadDashboard();}else alert(r&&r.message||'Could not delete record.');}).deleteClientModuleRecord(currentModule,currentRecords[i]._row,SID);}
function closeModal(){document.getElementById('modal').classList.add('hidden');document.getElementById('modal').style.display='none';}
function logout(){run(function(){localStorage.removeItem('bizos_client_session');localStorage.removeItem('bizos_client_user');location.href='?page=login';}).clientLogout(SID);}
function showLoadError(msg){document.getElementById('home').innerHTML='<div class="card p-8 text-center"><div class="text-red-600 text-xl font-bold mb-2">Dashboard could not load</div><p class="text-gray-500 mb-5">'+escapeHtml(msg||'The client workspace could not be loaded.')+'</p><button onclick="location.reload()" class="px-4 py-2 rounded-lg text-white" style="background:'+PRIMARY+'">Retry</button></div>';}
function escapeHtml(v){return String(v==null?'':v).replace(/[&<>"']/g,function(m){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m];});}
(function(){var u=JSON.parse(localStorage.getItem('bizos_client_user')||'{}');document.getElementById('userName').textContent=u.name||'Business User';document.getElementById('userEmail').textContent=u.email||'';if(!SID){location.href='?page=login';return;}loadDashboard();})();
</script>
</main>
</body>
</html>`;
}
