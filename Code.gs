/*************************************************************************************************
 * FOI PROJECT MASTER REGISTER — "ASTRA" EDITION
 * Turnkey Commercial Interior Fit-out · Project Control System
 * Google Apps Script (V8) · container-bound to a Google Sheet · single backend file
 *------------------------------------------------------------------------------------------------
 * FUNCTION INDEX
 *  ── SETUP & MENU ─────────────────────────────────────────────────────────────────────────────
 *   onOpen()                Spreadsheet menu "FOI Project Control"
 *   menuSetup / menuOpenWebApp / menuLoadSample / menuClearSample / menuBackup / menuTriggers /
 *   menuAddMeAdmin          menu handlers
 *   setupSystem()           creates + styles every sheet, seeds Settings dropdown masters
 *   ensureSheet_(name)      sheet bootstrap            styleSheet_(name, sh)  header styling etc.
 *   seedSettings_()         default dropdown masters + owner as Admin user
 *  ── WEB APP ──────────────────────────────────────────────────────────────────────────────────
 *   doGet()                 serves Index.html (SPA, XFrameOptions ALLOWALL)
 *  ── CONTEXT & ROLES (server-side enforcement) ─────────────────────────────────────────────────
 *   userCtx_()              email / name / role + permission flags
 *   canWrite_() canDelete_() scopeProjectsFor_() checkScope_() sanitizeOut_()
 *  ── DATA UTILS ────────────────────────────────────────────────────────────────────────────────
 *   ss_() sh_() readData_() invalidate_() rowToObj_() objToRow_() normVal_() normOut_()
 *   num_() pad_() todayStr_() nowStr_() dOffset_() diffDays_() TZ_() withLock_() friendly_()
 *   logSystem_() audit_() getSettings_() invalidateSettingsCache_() settingsList_() usersList_()
 *  ── BOOTSTRAP ─────────────────────────────────────────────────────────────────────────────────
 *   getBootstrapData()      ONE call: user, role, settings, lookups, schema, KPIs
 *  ── GENERIC CRUD ENGINE (schema-driven) ───────────────────────────────────────────────────────
 *   listRecords(p)          server-side search / filter / sort / pagination (50 rows/page)
 *   saveRecord(p)           validate → create/update → audit → rollup recompute (LockService)
 *   deleteRecord(p)         soft delete (Projects) / hard delete → audit → rollups
 *   cleanRecord_() validateRecord_() writeRecord_() nextId_() computeRow_() calcCtx_()
 *   getFinSums_() clientsMap_() recomputeProject_() autoDocChecklist_()
 *  ── PROJECT 360 & DASHBOARD ───────────────────────────────────────────────────────────────────
 *   getProjectDetail(id)    project + cross-sheet stats for the detail tabs
 *   buildDashboard_(f) / getDashboardData(f)   KPIs, chart aggregates, watch tables
 *   getNotifications()      bell + daily digest source
 *  ── DRIVE / PDF / EXPORT / EMAIL ──────────────────────────────────────────────────────────────
 *   createProjectFolder(id)         09-folder Drive structure, link saved to the project
 *   generateProjectSummaryPDF(id)   branded one-page PDF to Drive, returns link
 *   exportCSV(p)                    CSV of the current filtered view
 *   emailWeeklyMIS() dailyDigest() dailyBackup()   trigger jobs
 *   misHtml_(payload) projectPdfHtml_(proj, ctx)   HTML builders
 *  ── TRIGGERS ──────────────────────────────────────────────────────────────────────────────────
 *   createTriggers() deleteTriggers()
 *  ── SAMPLE DATA ───────────────────────────────────────────────────────────────────────────────
 *   loadSampleData() clearSampleData()   3 fully linked demo projects
 *************************************************************************************************/

/* ============================================ 1. APP CONFIG ================================== */

var APP = {
  name: 'FOI Project Master Register',
  short: 'FOI PMR',
  pageSize: 50,
  settingsCacheSecs: 21600, // 6h
  backupKeep: 14,
  backupFolder: 'PMR Backups',
  projectsFolder: 'FOI Projects',
  docFolderNames: ['01 Contract', '02 BOQ', '03 Drawings', '04 Approvals', '05 Procurement',
                   '06 DPR & Photos', '07 Billing', '08 Quality & Snag', '09 Handover'],
  terminalStatuses: ['Closed', 'Cancelled'],
  colors: {header: '#1F3864', master: '#1F3864', exec: '#4F46E5', log: '#64748B'}
};

/* Dropdown masters — every list below is seeded into the Settings sheet and is fully
 * editable from the web app (Settings ▸ Dropdown Lists). Nothing is hard-coded in screens. */
APP.seed = {
  'Project Type':        ['Office', 'Retail', 'Hospitality', 'Healthcare', 'Education', 'Industrial', 'Residential', 'Showroom', 'Data Center', 'Cafeteria/Food Court', 'Other'],
  'Contract Type':       ['Turnkey', 'Item Rate', 'Lump Sum', 'Cost Plus', 'PMC', 'Design & Build'],
  'Project Status':      ['Lead', 'Tender/Estimation', 'Awarded', 'Mobilization', 'Design Stage', 'Execution', 'Finishing', 'Snagging', 'Handover', 'Closed', 'On Hold', 'Cancelled'],
  'Live Status':         ['Awarded', 'Mobilization', "Design Stage", 'Execution', 'Finishing', 'Snagging'],
  'Phase':               ['Pre-Site', 'Site Preparation', 'Services', 'Finishes', 'Furniture & FF&E', 'Testing & Commissioning', 'Handover', 'Post-Handover'],
  'City/Branch':         ['Delhi', 'Bengaluru', 'Pune', 'Chennai', 'Mumbai'],
  'Priority':            ['Critical', 'High', 'Medium', 'Low'],
  'RAG':                 ['Green', 'Amber', 'Red'],
  'Payment Status':      ['Not Due', 'Due', 'Partially Paid', 'Paid', 'Overdue'],
  'PO Status':           ['Draft', 'Pending Approval', 'Approved', 'Partially Delivered', 'Delivered', 'Closed', 'Cancelled'],
  'Approval Status':     ['Pending', 'Approved', 'Rejected', 'On Hold'],
  'Trade/Work Category': ['Civil', 'Partition', 'False Ceiling', 'Flooring', 'Painting', 'Carpentry/Furniture', 'Electrical', 'HVAC', 'Plumbing', 'Fire Fighting', 'Fire Alarm', 'Data/Networking', 'Glazing', 'Signage', 'Soft Furnishing', 'Cleaning', 'Other'],
  'UOM':                 ['Sqft', 'Sqm', 'Rmt', 'Nos', 'Set', 'Lsum', 'Kg', 'Cum', 'Day', 'Month'],
  'GST Rate':            ['0', '5', '12', '18', '28'],
  'Cost Head':           ['Material', 'Labour', 'Equipment', 'Subcontract', 'Design Fee', 'Statutory/Compliance', 'Site Overhead', 'Logistics', 'Other'],
  'Bill Type':           ['Advance', 'RA', 'Final', 'Retention'],
  'Discipline':          ['Architecture', 'Interior', 'Structural', 'MEP', 'HVAC', 'Electrical', 'Plumbing', 'Fire', 'ID/Furniture', 'Facade/Glazing', 'Signage'],
  'Weather':             ['Sunny', 'Cloudy', 'Rainy', 'Humid', 'Windy'],
  'Credit Rating':       ['A+', 'A', 'B+', 'B', 'C'],
  'Empanelment Status':  ['Empanelled', 'Preferred', 'Under Review', 'Blacklisted'],
  'Doc Type':            ['WO', 'BOQ Approval', 'Drawings', 'Insurance', 'Fire NOC', 'EHS Plan', 'Labour License', 'Completion Certificate', 'Warranty', 'As-built', 'O&M Manual', 'Handover Certificate'],
  'Risk Type':           ['Risk', 'Issue'],
  'Risk Category':       ['Commercial', 'Design', 'Client', 'Statutory', 'Execution', 'Resource', 'Quality', 'Safety', 'Vendor'],
  'Risk Status':         ['Open', 'In Progress', 'Mitigated', 'Closed'],
  'Snag Status':         ['Open', 'In Progress', 'Closed', 'Verified'],
  'Doc Status':          ['Pending', 'In Progress', 'Received', 'Not Applicable'],
  'User Role':           ['Admin', 'Project Director', 'PM', 'Coordinator', 'Viewer']
};

/* ============================================ 2. SHEET SCHEMA ================================
 * One config per sheet drives headers, styling, validation, forms, tables and the generic CRUD
 * engine. Adding a field = adding a line here. t: id|text|textarea|number|date|select|yn|email
 *   req:required  ro:readonly in UI  calc:server-computed  uniq:duplicate check
 *   sens:cost/margin (hidden from non Admin/Director)  hidden:never leaves the server
 *   cat:Settings dropdown category  min/max:number bounds  w:column width
 * ============================================================================================= */
function C(k, t, o) { var c = {k: k, t: t}; if (o) { for (var x in o) c[x] = o[x]; } return c; }

var SCHEMA = {

  Settings: {
    idKey: 'Key', idPrefix: 'SET', adminOnly: true, tab: APP.colors.master, ds: {col: 'Category', dir: 'asc'},
    cols: [
      C('Key', 'id', {ro: 1}),
      C('Category', 'select', {cat: '#SettingCats#', req: 1, w: 160}),
      C('Value', 'text', {req: 1, w: 220}),
      C('Info 1', 'text', {w: 200}),
      C('Info 2', 'text', {w: 140}),
      C('Sort', 'number', {w: 80}),
      C('Active', 'yn', {w: 90})
    ]
  },

  Projects: {
    idKey: 'Project ID', idPrefix: 'FOI', softDelete: true, projectKey: 'Project ID', tab: APP.colors.master,
    ds: {col: 'Project Code', dir: 'asc'},
    cols: [
      C('Project ID', 'id', {ro: 1, w: 110}),
      C('Project Code', 'text', {req: 1, uniq: 1, w: 110}),
      C('Project Name', 'text', {req: 1, w: 260}),
      C('Client ID', 'fkclient', {w: 100}),
      C('Client Name', 'calc', {w: 190}),
      C('Contact Person', 'text', {w: 140}),
      C('Phone', 'text', {w: 120}),
      C('Email', 'email', {w: 180}),
      C('Project Type', 'select', {cat: 'Project Type', w: 130}),
      C('Contract Type', 'select', {cat: 'Contract Type', w: 130}),
      C('City', 'select', {cat: 'City/Branch', w: 110}),
      C('Site Address', 'text', {w: 220}),
      C('Built-up Area (sq ft)', 'number', {w: 120}),
      C('Scope', 'textarea', {w: 260}),
      C('WO/LOA No', 'text', {w: 130}),
      C('WO Date', 'date', {w: 105}),
      C('Contract Value (excl GST)', 'number', {req: 1, w: 150}),
      C('GST %', 'select', {cat: 'GST Rate', num: 1, w: 80}),
      C('GST Amount', 'calc', {w: 130}),
      C('Total Contract Value', 'calc', {w: 150}),
      C('Approved Budget', 'number', {sens: 1, w: 140}),
      C('Est. Margin %', 'number', {sens: 1, w: 110}),
      C('Project Director', 'select', {cat: '#Users#', w: 140}),
      C('Project Manager', 'select', {cat: '#Users#', w: 140}),
      C('Coordinator', 'select', {cat: '#Users#', w: 130}),
      C('Site Engineer', 'select', {cat: '#Users#', w: 130}),
      C('Designer', 'select', {cat: '#Users#', w: 130}),
      C('Client PMC', 'text', {w: 130}),
      C('Start Date', 'date', {w: 105}),
      C('Planned End', 'date', {w: 105}),
      C('Revised End', 'date', {w: 105}),
      C('Actual End', 'date', {w: 105}),
      C('Duration (days)', 'calc', {w: 100}),
      C('Days Remaining', 'calc', {w: 100}),
      C('Delay Days', 'calc', {w: 90}),
      C('Status', 'select', {cat: 'Project Status', req: 1, w: 130}),
      C('Phase', 'select', {cat: 'Phase', w: 150}),
      C('Planned %', 'number', {w: 90}),
      C('Actual %', 'number', {w: 90}),
      C('Variance %', 'calc', {w: 90}),
      C('RAG', 'calc', {w: 90}),
      C('Billed', 'calc', {w: 130}),
      C('Received', 'calc', {w: 130}),
      C('Outstanding', 'calc', {w: 130}),
      C('Retention Held', 'calc', {sens: 1, w: 120}),
      C('Cost Incurred', 'calc', {sens: 1, w: 130}),
      C('Cost-to-Complete', 'calc', {sens: 1, w: 130}),
      C('Projected Margin %', 'calc', {sens: 1, w: 120}),
      C('Next Milestone', 'text', {w: 190}),
      C('Milestone Date', 'date', {w: 105}),
      C('Key Risk', 'text', {w: 190}),
      C('Priority', 'select', {cat: 'Priority', w: 100}),
      C('Last Updated', 'calc', {w: 140}),
      C('Updated By', 'calc', {w: 150}),
      C('Remarks', 'textarea', {w: 220}),
      C('Drive Folder Link', 'text', {w: 200}),
      C('Deleted', 'flag', {hidden: 1, w: 80})
    ]
  },

  Clients: {
    idKey: 'Client ID', idPrefix: 'CLI', projectKey: '', tab: APP.colors.master, ds: {col: 'Name', dir: 'asc'},
    cols: [
      C('Client ID', 'id', {ro: 1, w: 100}),
      C('Name', 'text', {req: 1, uniq: 1, w: 230}),
      C('Industry', 'text', {w: 140}),
      C('GST No', 'text', {w: 160}),
      C('PAN', 'text', {w: 120}),
      C('Billing Address', 'textarea', {w: 240}),
      C('Contact Persons', 'textarea', {w: 240}),
      C('Payment Terms (days)', 'number', {w: 120}),
      C('Credit Rating', 'select', {cat: 'Credit Rating', w: 100}),
      C('Active Projects', 'calc', {w: 100})
    ]
  },

  Vendors: {
    idKey: 'Vendor ID', idPrefix: 'VEN', tab: APP.colors.master, ds: {col: 'Vendor Name', dir: 'asc'},
    cols: [
      C('Vendor ID', 'id', {ro: 1, w: 100}),
      C('Vendor Name', 'text', {req: 1, uniq: 1, w: 220}),
      C('Trade', 'select', {cat: 'Trade/Work Category', w: 150}),
      C('Contact Person', 'text', {w: 140}),
      C('Phone', 'text', {w: 120}),
      C('Email', 'email', {w: 180}),
      C('City', 'select', {cat: 'City/Branch', w: 110}),
      C('GST No', 'text', {w: 160}),
      C('PAN', 'text', {w: 120}),
      C('Bank Details (masked)', 'text', {w: 170}),
      C('Payment Terms (days)', 'number', {w: 120}),
      C('Rating (1-5)', 'number', {min: 1, max: 5, w: 90}),
      C('Empanelment Status', 'select', {cat: 'Empanelment Status', w: 140}),
      C('Active WOs', 'calc', {w: 90}),
      C('Total Business Value', 'calc', {w: 150})
    ]
  },

  'BOQ': {
    idKey: 'Line ID', idPrefix: 'BOQ', projectKey: 'Project ID', tab: APP.colors.exec,
    ds: {col: 'Line ID', dir: 'asc'},
    cols: [
      C('Line ID', 'id', {ro: 1, w: 120}),
      C('Project ID', 'fkproj', {req: 1, w: 110}),
      C('Section/Trade', 'select', {cat: 'Trade/Work Category', req: 1, w: 150}),
      C('Item Code', 'text', {w: 110}),
      C('Description', 'textarea', {w: 280}),
      C('UOM', 'select', {cat: 'UOM', w: 80}),
      C('Qty', 'number', {req: 1, w: 100}),
      C('Client Rate', 'number', {w: 110}),
      C('Client Amount', 'calc', {w: 130}),
      C('Budget Rate', 'number', {w: 110}),
      C('Budget Amount', 'calc', {w: 130}),
      C('Qty Executed', 'number', {w: 100}),
      C('% Complete', 'calc', {w: 90}),
      C('Executed Value', 'calc', {w: 130}),
      C('Remarks', 'textarea', {w: 180})
    ]
  },

  'Procurement': {
    idKey: 'PO/WO No', idPrefix: 'PO', projectKey: 'Project ID', tab: APP.colors.exec,
    ds: {col: 'Date', dir: 'desc'},
    cols: [
      C('PO/WO No', 'id', {ro: 1, w: 120}),
      C('Date', 'date', {req: 1, w: 105}),
      C('Project ID', 'fkproj', {req: 1, w: 110}),
      C('Vendor', 'fkvendor', {w: 190}),
      C('Trade', 'select', {cat: 'Trade/Work Category', w: 140}),
      C('Description', 'textarea', {w: 260}),
      C('Qty', 'number', {w: 90}),
      C('UOM', 'select', {cat: 'UOM', w: 80}),
      C('Rate', 'number', {w: 110}),
      C('Basic Amount', 'calc', {w: 130}),
      C('GST %', 'select', {cat: 'GST Rate', num: 1, w: 80}),
      C('GST Amount', 'calc', {w: 120}),
      C('Total Amount', 'calc', {w: 130}),
      C('Status', 'select', {cat: 'PO Status', w: 140}),
      C('Required By', 'date', {w: 105}),
      C('Delivery Date', 'date', {w: 105}),
      C('Delay Days', 'calc', {w: 90}),
      C('Paid', 'number', {w: 110}),
      C('Balance Amount', 'calc', {w: 120}),
      C('Approved By', 'select', {cat: '#Users#', w: 130})
    ]
  },

  'Cost Tracker': {
    idKey: 'Entry ID', idPrefix: 'CST', projectKey: 'Project ID', adminOnlyCost: true, tab: APP.colors.exec,
    ds: {col: 'Date', dir: 'desc'},
    cols: [
      C('Entry ID', 'id', {ro: 1, w: 120}),
      C('Date', 'date', {req: 1, w: 105}),
      C('Project ID', 'fkproj', {req: 1, w: 110}),
      C('Cost Head', 'select', {cat: 'Cost Head', w: 150}),
      C('Vendor', 'text', {w: 170}),
      C('Invoice No', 'text', {w: 120}),
      C('Amount', 'number', {req: 1, sens: 1, w: 120}),
      C('GST %', 'select', {cat: 'GST Rate', num: 1, sens: 1, w: 80}),
      C('GST Amount', 'calc', {sens: 1, w: 110}),
      C('Total Amount', 'calc', {sens: 1, w: 120}),
      C('Payment Status', 'select', {cat: 'Payment Status', w: 130}),
      C('Remarks', 'textarea', {w: 180})
    ]
  },

  'Billing': {
    idKey: 'Bill No', idPrefix: 'RA', projectKey: 'Project ID', tab: APP.colors.exec,
    ds: {col: 'Bill Date', dir: 'desc'},
    cols: [
      C('Bill No', 'id', {ro: 1, w: 110}),
      C('Project ID', 'fkproj', {req: 1, w: 110}),
      C('Bill Type', 'select', {cat: 'Bill Type', w: 100}),
      C('Bill Date', 'date', {req: 1, w: 105}),
      C('Gross Amount', 'number', {req: 1, w: 130}),
      C('Retention %', 'number', {w: 90}),
      C('Retention Amount', 'calc', {w: 130}),
      C('GST %', 'select', {cat: 'GST Rate', num: 1, w: 80}),
      C('GST Amount', 'calc', {w: 120}),
      C('Net Receivable', 'calc', {w: 130}),
      C('Submitted Date', 'date', {w: 105}),
      C('Due Date', 'calc', {w: 105}),
      C('Received Amount', 'number', {w: 130}),
      C('Received Date', 'date', {w: 105}),
      C('Outstanding', 'calc', {w: 120}),
      C('Ageing Days', 'calc', {w: 90}),
      C('Status', 'select', {cat: 'Payment Status', w: 130})
    ]
  },

  'Drawings & Approvals': {
    idKey: 'Drawing No', idPrefix: 'DRW', projectKey: 'Project ID', tab: APP.colors.exec,
    ds: {col: 'Submitted Date', dir: 'desc'},
    cols: [
      C('Drawing No', 'id', {ro: 1, w: 110}),
      C('Ref No', 'text', {w: 120}),
      C('Project ID', 'fkproj', {req: 1, w: 110}),
      C('Title', 'text', {req: 1, w: 220}),
      C('Discipline', 'select', {cat: 'Discipline', w: 120}),
      C('Revision', 'text', {w: 80}),
      C('Issued Date', 'date', {w: 105}),
      C('Submitted Date', 'date', {w: 105}),
      C('Approval Status', 'select', {cat: 'Approval Status', w: 120}),
      C('Approved Date', 'date', {w: 105}),
      C('Pending Days', 'calc', {w: 90}),
      C('Remarks', 'textarea', {w: 180})
    ]
  },

  'DPR': {
    idKey: 'Entry ID', idPrefix: 'DPR', projectKey: 'Project ID', tab: APP.colors.exec,
    ds: {col: 'Date', dir: 'desc'},
    cols: [
      C('Entry ID', 'id', {ro: 1, w: 120}),
      C('Date', 'date', {req: 1, w: 105}),
      C('Project ID', 'fkproj', {req: 1, w: 110}),
      C('Site Engineer', 'select', {cat: '#Users#', w: 140}),
      C('Weather', 'select', {cat: 'Weather', w: 90}),
      C('Skilled Manpower', 'number', {w: 110}),
      C('Unskilled Manpower', 'number', {w: 120}),
      C('Work Done', 'textarea', {w: 280}),
      C('Planned for Tomorrow', 'textarea', {w: 240}),
      C('Material Received', 'textarea', {w: 200}),
      C('Issues', 'textarea', {w: 200}),
      C('Photo Link', 'text', {w: 170}),
      C('Safety Incident (Y/N)', 'yn', {w: 110})
    ]
  },

  'Variation Orders': {
    idKey: 'VO No', idPrefix: 'VO', projectKey: 'Project ID', tab: APP.colors.exec,
    ds: {col: 'Date', dir: 'desc'},
    cols: [
      C('VO No', 'id', {ro: 1, w: 110}),
      C('Project ID', 'fkproj', {req: 1, w: 110}),
      C('Description', 'textarea', {req: 1, w: 280}),
      C('Requested By', 'text', {w: 140}),
      C('Date', 'date', {w: 105}),
      C('Amount', 'number', {req: 1, w: 120}),
      C('Approval Status', 'select', {cat: 'Approval Status', w: 120}),
      C('Approved Date', 'date', {w: 105}),
      C('Timeline Impact (days)', 'number', {w: 120})
    ]
  },

  'Risks & Issues': {
    idKey: 'ID', idPrefix: 'RSK', projectKey: 'Project ID', tab: APP.colors.exec,
    ds: {col: 'ID', dir: 'asc'},
    cols: [
      C('ID', 'id', {ro: 1, w: 110}),
      C('Project ID', 'fkproj', {req: 1, w: 110}),
      C('Type', 'select', {cat: 'Risk Type', w: 90}),
      C('Description', 'textarea', {req: 1, w: 280}),
      C('Category', 'select', {cat: 'Risk Category', w: 120}),
      C('Probability (1-5)', 'number', {min: 1, max: 5, w: 100}),
      C('Impact (1-5)', 'number', {min: 1, max: 5, w: 90}),
      C('Score', 'calc', {w: 80}),
      C('Owner', 'select', {cat: '#Users#', w: 130}),
      C('Mitigation', 'textarea', {w: 240}),
      C('Target Date', 'date', {w: 105}),
      C('Status', 'select', {cat: 'Risk Status', w: 110})
    ]
  },

  'Snags & Handover': {
    idKey: 'Snag ID', idPrefix: 'SNG', projectKey: 'Project ID', tab: APP.colors.exec,
    ds: {col: 'Raised Date', dir: 'desc'},
    cols: [
      C('Snag ID', 'id', {ro: 1, w: 110}),
      C('Project ID', 'fkproj', {req: 1, w: 110}),
      C('Location', 'text', {w: 150}),
      C('Description', 'textarea', {req: 1, w: 280}),
      C('Trade', 'select', {cat: 'Trade/Work Category', w: 140}),
      C('Raised Date', 'date', {w: 105}),
      C('Assigned To', 'select', {cat: '#Users#', w: 130}),
      C('Target Date', 'date', {w: 105}),
      C('Closed Date', 'date', {w: 105}),
      C('Status', 'select', {cat: 'Snag Status', w: 110}),
      C('Photo Link', 'text', {w: 170})
    ]
  },

  'Documents Checklist': {
    idKey: 'Doc ID', idPrefix: 'DOC', projectKey: 'Project ID', tab: APP.colors.exec,
    ds: {col: 'Doc Type', dir: 'asc'},
    cols: [
      C('Doc ID', 'id', {ro: 1, w: 110}),
      C('Project ID', 'fkproj', {req: 1, w: 110}),
      C('Doc Type', 'select', {cat: 'Doc Type', w: 170}),
      C('Required', 'yn', {w: 90}),
      C('Received', 'yn', {w: 90}),
      C('Date', 'date', {w: 105}),
      C('Link', 'text', {w: 200}),
      C('Status', 'select', {cat: 'Doc Status', w: 120})
    ]
  },

  'Audit Log': {
    idKey: 'Timestamp', readOnly: true, tab: APP.colors.log, ds: {col: 'Timestamp', dir: 'desc'},
    cols: [
      C('Timestamp', 'text', {ro: 1, w: 150}),
      C('User', 'text', {ro: 1, w: 160}),
      C('Sheet', 'text', {ro: 1, w: 150}),
      C('Row Key', 'text', {ro: 1, w: 120}),
      C('Field', 'text', {ro: 1, w: 150}),
      C('Old', 'text', {ro: 1, w: 160}),
      C('New', 'text', {ro: 1, w: 160}),
      C('Action', 'text', {ro: 1, w: 100})
    ]
  },

  'System Log': {
    idKey: 'Timestamp', readOnly: true, hiddenSheet: true, tab: APP.colors.log,
    ds: {col: 'Timestamp', dir: 'desc'},
    cols: [
      C('Timestamp', 'text', {ro: 1, w: 150}),
      C('Level', 'text', {ro: 1, w: 80}),
      C('Function', 'text', {ro: 1, w: 160}),
      C('Message', 'text', {ro: 1, w: 300}),
      C('Details', 'text', {ro: 1, w: 300})
    ]
  }
};

var SHEET_ORDER = ['Settings', 'Projects', 'Clients', 'Vendors', 'BOQ', 'Procurement', 'Cost Tracker',
                   'Billing', 'Drawings & Approvals', 'DPR', 'Variation Orders', 'Risks & Issues',
                   'Snags & Handover', 'Documents Checklist', 'Audit Log', 'System Log'];

/* per-execution memo (spreadsheet object, sheet data, settings, fin sums) */
var _m = {ss: null, data: {}, calcCtx: null, ctx: null, scope: null, fin: null};

/* ============================================ 3. MENU & SETUP ================================ */

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('FOI Project Control')
      .addItem('Setup System', 'menuSetup')
      .addItem('Open Web App', 'menuOpenWebApp')
      .addSeparator()
      .addItem('Load Sample Data', 'menuLoadSample')
      .addItem('Clear Sample Data', 'menuClearSample')
      .addSeparator()
      .addItem('Backup Now', 'menuBackup')
      .addItem('Create Triggers', 'menuTriggers')
      .addItem('Add Me as Admin', 'menuAddMeAdmin')
      .addToUi();
  } catch (e) { /* ui not available in some contexts */ }
}

function menuSetup()      { uiMsg_('Setup System', setupSystem()); }
function menuLoadSample() { uiMsg_('Load Sample Data', loadSampleData()); }
function menuClearSample(){ uiMsg_('Clear Sample Data', clearSampleData()); }
function menuBackup()     { uiMsg_('Backup Now', dailyBackup(true)); }
function menuTriggers()   { uiMsg_('Create Triggers', createTriggers()); }
function menuAddMeAdmin() { uiMsg_('Add Me as Admin', seedOwnerUser_(true)); }

function menuOpenWebApp() {
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  uiMsg_('Open Web App', url ? 'Web App URL:\n' + url + '\n\nDeploy: Extensions ▸ Apps Script ▸ Deploy ▸ New deployment ▸ Web app (Execute as: Me).'
                             : 'No deployment found yet.\nOpen Extensions ▸ Apps Script ▸ Deploy ▸ New deployment ▸ Web app\n(Execute as: Me, Access: Anyone within domain or Anyone with link).');
}

function uiMsg_(title, msg) {
  try { SpreadsheetApp.getUi().alert(title, String(msg || 'Done.'), SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { logSystem_('INFO', title, msg, ''); }
}

/** Creates and styles every sheet, seeds dropdown masters and the owner as Admin. Idempotent. */
function setupSystem() {
  try {
    SHEET_ORDER.forEach(function (name) { ensureSheet_(name); });
    var seeded = seedSettings_();
    invalidate_();
    invalidateSettingsCache_();
    return 'Setup complete. ' + SHEET_ORDER.length + ' registers ready' + (seeded ? ', Settings seeded with default dropdown masters.' : '.');
  } catch (e) {
    logSystem_('ERROR', 'setupSystem', e && e.message, e && e.stack);
    throw new Error(friendly_(e));
  }
}

function ensureSheet_(name) {
  var ss = ss_(), sc = SCHEMA[name];
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sc.hiddenSheet && !sh.isSheetHidden()) sh.hideSheet();
  styleSheet_(name, sh);
  return sh;
}

/** Navy header, white bold text, frozen row, filter, formats, widths, validations, protection. */
function styleSheet_(name, sh) {
  var sc = SCHEMA[name];
  var headers = sc.cols.map(function (c) { return c.k; });
  var n = headers.length;
  var head = sh.getRange(1, 1, 1, n).setValues([headers]);
  head.setBackground(APP.colors.header).setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(10).setVerticalAlignment('middle').setWrap(true);
  sh.setFrozenRows(1);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(1, 1, 1, n).createFilter();
  sh.setTabColor(sc.tab || APP.colors.master);
  sc.cols.forEach(function (c, i) {
    var col = i + 1, width = c.w || (c.t === 'textarea' ? 260 : c.t === 'text' ? 170 : c.t === 'number' ? 130 : c.t === 'date' ? 110 : c.t === 'select' ? 140 : 110);
    sh.setColumnWidth(col, width);
    var rng = sh.getRange(2, col, sh.getMaxRows() - 1, 1);
    if (c.t === 'date') rng.setNumberFormat('yyyy-mm-dd');
    if (c.t === 'number') rng.setNumberFormat('#,##0.###');
    if (c.t === 'select' && c.cat && c.cat.charAt(0) !== '#' && APP.seed[c.cat]) {
      rng.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(APP.seed[c.cat], true).setAllowInvalid(true)
        .setHelpText('Value should ideally come from Settings ▸ "' + c.cat + '" (web app is the source of truth).').build());
    }
    if (c.t === 'yn') rng.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'], true).setAllowInvalid(true).build());
  });
  var guard = sh.getRange(1, 1, 1, n).protect();
  guard.setDescription('PMR header guard — do not rename columns');
  try { guard.setWarningOnly(true); } catch (e) {}
}

/** Seeds all dropdown masters + the current owner as Admin (only if Settings is empty). */
function seedSettings_() {
  var sh = sh_('Settings');
  if (sh.getLastRow() > 1) return false;
  var rows = [], i = 0;
  Object.keys(APP.seed).forEach(function (cat) {
    APP.seed[cat].forEach(function (v, j) {
      i++;
      rows.push(['SET-' + pad_(i, 4), cat, v, '', '', j + 1, 'Yes']);
    });
  });
  seedOwnerUser_(false, rows);
  sh.getRange(2, 1, rows.length, 7).setValues(rows);
  return true;
}

/** Adds the effective (owner) user as Admin. mode=true → called from menu on a live system. */
function seedOwnerUser_(fromMenu, rows) {
  var email = '', name = 'Owner';
  try { email = Session.getEffectiveUser().getEmail() || ''; } catch (e) {}
  if (email) name = email.split('@')[0].replace(/[._-]+/g, ' ').replace(/\b\w/g, function (m) { return m.toUpperCase(); });
  if (fromMenu) {
    var existing = usersList_();
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i].email || '').toLowerCase() === String(email).toLowerCase()) {
        return 'You are already configured as: ' + existing[i].name + ' (' + existing[i].role + ')';
      }
    }
    var sh = sh_('Settings');
    var key = 'SET-' + pad_(sh.getLastRow() + 1000, 4);
    sh.getRange(sh.getLastRow() + 1, 1, 1, 7).setValues([[key, 'Users', name, email, 'Admin', 0, 'Yes']]);
    invalidate_(); invalidateSettingsCache_(); _m.ctx = null;
    return 'Added ' + name + ' (' + email + ') as Admin. Refresh the web app.';
  }
  rows.push(['SET-USR-0001', 'Users', name, email, 'Admin', 0, 'Yes']);
  return rows;
}

/* ============================================ 4. WEB APP ENTRY =============================== */

function doGet() {
  try {
    return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle(APP.name)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) {
    logSystem_('ERROR', 'doGet', e && e.message, e && e.stack);
    throw new Error(friendly_(e));
  }
}

/* ============================================ 5. CORE UTILS ================================== */

function TZ_() { return Session.getScriptTimeZone() || 'Asia/Kolkata'; }
function ss_() { if (!_m.ss) _m.ss = SpreadsheetApp.getActive(); return _m.ss; }
function sh_(name) { var s = ss_().getSheetByName(name); if (!s) throw new Error('Sheet "' + name + '" is missing. Run FOI Project Control ▸ Setup System.'); return s; }

function readData_(name) {
  if (_m.data[name]) return _m.data[name];
  var sh = sh_(name), sc = SCHEMA[name];
  var lr = sh.getLastRow(), lc = Math.max(sh.getLastColumn(), sc.cols.length);
  var headers = lr > 0 ? sh.getRange(1, 1, 1, lc).getValues()[0] : sc.cols.map(function (c) { return c.k; });
  var hm = {};
  headers.forEach(function (h, i) { hm[String(h).trim()] = i; });
  // self-heal: append any schema columns missing on the sheet (schema evolution)
  var missing = sc.cols.filter(function (c) { return hm[c.k] === undefined; });
  if (missing.length) {
    var start = lc + 1;
    sh.getRange(1, start, 1, missing.length).setValues([missing.map(function (c) { return c.k; })]);
    missing.forEach(function (c, j) { hm[c.k] = start - 1 + j; });
    styleSheet_(name, sh);
    lc = sh.getLastColumn();
  }
  var rows = lr > 1 ? sh.getRange(2, 1, lr - 1, lc).getValues() : [];
  _m.data[name] = {sheet: sh, headers: headers, hm: hm, rows: rows, sc: sc};
  return _m.data[name];
}

function invalidate_(name) {
  if (name) { delete _m.data[name]; if (name === 'Billing' || name === 'Cost Tracker') _m.fin = null; if (name === 'Clients' || name === 'Settings' || name === 'Projects') _m.calcCtx = null; }
  else { _m.data = {}; _m.fin = null; _m.calcCtx = null; }
}

function rowToObj_(d, r) {
  var o = {};
  d.sc.cols.forEach(function (c) { var i = d.hm[c.k]; o[c.k] = (i === undefined || i >= r.length) ? '' : normVal_(r[i]); });
  return o;
}

function objToRow_(name, rec) {
  var sc = SCHEMA[name];
  return sc.cols.map(function (c) {
    var v = rec[c.k];
    if (v === undefined || v === null) return '';
    if (c.t === 'number') return v === '' ? '' : num_(v);
    return v;
  });
}

/** Server → client normalisation: Dates become 'yyyy-MM-dd' strings (no TZ drift). */
function normVal_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TZ_(), 'yyyy-MM-dd');
  return v;
}
function normOut_(v) { return v === null || v === undefined ? '' : String(v); }
function num_(v) { if (v === '' || v === null || v === undefined) return 0; var n = parseFloat(String(v).replace(/[,₹\s]/g, '')); return isNaN(n) ? 0 : n; }
function round1_(n) { return Math.round(n * 10) / 10; }
function pad_(n, l) { var s = String(Math.max(0, Math.floor(n))); while (s.length < l) s = '0' + s; return s; }
function todayStr_() { return Utilities.formatDate(new Date(), TZ_(), 'yyyy-MM-dd'); }
function nowStr_() { return Utilities.formatDate(new Date(), TZ_(), 'yyyy-MM-dd HH:mm'); }
function dOffset_(days) { return Utilities.formatDate(new Date(Date.now() + days * 86400000), TZ_(), 'yyyy-MM-dd'); }
function diffDays_(a, b) { return Math.round((Date.parse(b + 'T00:00:00Z') - Date.parse(a + 'T00:00:00Z')) / 86400000); }
function withLock_(fn) { var lock = LockService.getScriptLock(); lock.waitLock(30000); try { return fn(); } finally { try { lock.releaseLock(); } catch (e) {} } }
function friendly_(e) {
  var msg = (e && e.message) ? String(e.message) : String(e);
  if (msg.indexOf('Access denied') === 0) return msg;
  if (msg.indexOf('Duplicate') === 0 || msg.indexOf('Record ') === 0 || msg.indexOf('Sheet "') === 0 || msg.indexOf('Validation') === 0) return msg;
  return msg.length > 240 ? msg.slice(0, 240) + '…' : msg;
}

function logSystem_(level, fn, msg, details) {
  try {
    var sh = ss_().getSheetByName('System Log');
    if (!sh) return;
    sh.appendRow([nowStr_(), level, fn, String(msg || '').slice(0, 480), String(details || '').slice(0, 480)]);
  } catch (e) { console.error('logSystem failed: ' + e); }
}

function audit_(action, sheetName, rowKey, diffs, ctx) {
  try {
    var sh = sh_('Audit Log');
    var rows = [];
    if (!diffs || !diffs.length) diffs = [{f: '—', o: '', n: ''}];
    diffs.slice(0, 15).forEach(function (d) {
      rows.push([nowStr_(), (ctx && ctx.email) || 'system', sheetName, String(rowKey || ''), d.f,
                 String(d.o).slice(0, 180), String(d.n).slice(0, 180), action]);
    });
    sh.getRange(Math.max(sh.getLastRow() + 1, 2), 1, rows.length, 8).setValues(rows);
  } catch (e) { logSystem_('WARN', 'audit_', e && e.message, ''); }
}

/* ============================================ 6. SETTINGS CACHE ============================== */

function getSettings_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('PMR_SETTINGS_V1');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var d = readData_('Settings');
  var cats = {}, users = [];
  d.rows.forEach(function (r) {
    var o = rowToObj_(d, r);
    if (String(o['Active'] || 'Yes') !== 'Yes' && String(o['Active'] || '') !== '') return;
    var cat = String(o['Category'] || '');
    if (!cat) return;
    if (cat === 'Users') {
      users.push({name: String(o['Value'] || ''), email: String(o['Info 1'] || ''), role: String(o['Info 2'] || 'Viewer')});
    } else {
      if (!cats[cat]) cats[cat] = [];
      cats[cat].push(String(o['Value'] || ''));
    }
  });
  Object.keys(cats).forEach(function (k) { cats[k] = dedupe_(cats[k]); });
  users = users.filter(function (u) { return u.name; });
  var payload = {cats: cats, users: users};
  try { cache.put('PMR_SETTINGS_V1', JSON.stringify(payload), APP.settingsCacheSecs); } catch (e) {}
  return payload;
}
function dedupe_(arr) { var seen = {}, out = []; arr.forEach(function (v) { v = String(v || '').trim(); if (v && !seen[v]) { seen[v] = 1; out.push(v); } }); return out; }
function invalidateSettingsCache_() { try { CacheService.getScriptCache().remove('PMR_SETTINGS_V1'); } catch (e) {} _m.calcCtx = null; }
function settingsList_(cat) {
  if (cat === '#Users#') return usersList_().map(function (u) { return u.name; });
  if (cat === '#SettingCats#') return settingCats_();
  var s = getSettings_();
  return s.cats[cat] || [];
}
function usersList_() { return getSettings_().users; }
function settingCats_() {
  var cats = Object.keys(APP.seed);
  var d = readData_('Settings');
  d.rows.forEach(function (r) { var c = String(r[d.hm['Category']] || '').trim(); if (c && cats.indexOf(c) < 0) cats.push(c); });
  return cats;
}
function liveStatuses_() { var l = settingsList_('Live Status'); return l.length ? l : ['Awarded', 'Mobilization', 'Design Stage', 'Execution', 'Finishing', 'Snagging']; }

/* ============================================ 7. CONTEXT & ROLES =============================
 * Roles are enforced on the SERVER for every read and write — the client hides controls only
 * as a convenience. Admin: everything · Project Director: everything · PM: own projects only ·
 * Coordinator: add/edit, no delete · Viewer: read-only. Cost/margin data: Admin + Director.
 * ============================================================================================= */
function userCtx_() {
  if (_m.ctx) return _m.ctx;
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  if (!email) { try { email = Session.getEffectiveUser().getEmail() || ''; } catch (e2) {} }
  email = String(email).toLowerCase();
  var users = usersList_();
  var u = null;
  for (var i = 0; i < users.length; i++) { if (String(users[i].email || '').toLowerCase() === email) { u = users[i]; break; } }
  var role = u ? String(u.role || 'Viewer') : 'Viewer';
  _m.ctx = {
    email: email,
    name: u ? u.name : (email ? email.split('@')[0] : 'Guest'),
    role: role,
    isAdmin: role === 'Admin',
    seeCost: role === 'Admin' || role === 'Project Director',
    canWrite: role !== 'Viewer',
    canDelete: role === 'Admin' || role === 'Project Director' || role === 'PM'
  };
  return _m.ctx;
}

function canWrite_(ctx) { return ctx.canWrite; }
function canDelete_(ctx, sheetName, rec) {
  if (ctx.role === 'Admin' || ctx.role === 'Project Director') return true;
  if (ctx.role !== 'PM') return false; // Coordinator/Viewer never delete
  var allowed = scopeProjectsFor_(ctx);
  if (!allowed) return true;
  if (sheetName === 'Projects') return allowed.indexOf(String(rec['Project ID'] || '')) >= 0;
  var pk = SCHEMA[sheetName] && SCHEMA[sheetName].projectKey;
  return !!pk && allowed.indexOf(String(rec[pk] || '')) >= 0;
}

/** null = all projects; PM gets the set of project IDs they manage. */
function scopeProjectsFor_(ctx) {
  if (ctx.role !== 'PM') return null;
  if (_m.scope) return _m.scope;
  var d = readData_('Projects'), out = [];
  var iName = d.hm['Project Manager'], iId = d.hm['Project ID'], iDel = d.hm['Deleted'];
  d.rows.forEach(function (r) {
    if (iDel >= 0 && String(r[iDel]) === 'Yes') return;
    if (String(r[iName] || '').trim() === ctx.name.trim()) out.push(String(r[iId] || ''));
  });
  _m.scope = out;
  return out;
}

function checkScope_(ctx, sheetName, rec, existing) {
  if (ctx.role !== 'PM') return;
  var allowed = scopeProjectsFor_(ctx);
  if (!allowed) return;
  if (sheetName === 'Projects') {
    var key = existing ? existing['Project ID'] : rec['Project ID'];
    var mgr = existing ? existing['Project Manager'] : rec['Project Manager'];
    if (allowed.indexOf(String(key || '')) < 0 && String(mgr || '').trim() !== ctx.name.trim())
      throw new Error('Access denied: PMs can only manage their own projects.');
    if (!existing && String(rec['Project Manager'] || '').trim() !== ctx.name.trim())
      throw new Error('Access denied: as PM you can only create projects where you are the Project Manager.');
    return;
  }
  var pk = SCHEMA[sheetName].projectKey;
  if (!pk) return;
  var pid = String(rec[pk] || '');
  if (allowed.indexOf(pid) < 0) throw new Error('Access denied: project ' + (pid || '(blank)') + ' is not assigned to you.');
}

/** Strips hidden/sensitive fields before data leaves the server; masks bank details. */
function sanitizeOut_(ctx, sheetName, obj) {
  var sc = SCHEMA[sheetName], out = {};
  sc.cols.forEach(function (c) {
    if (c.hidden) return;
    if (c.sens && !ctx.seeCost) return;
    var v = obj[c.k];
    out[c.k] = (c.k === 'Bank Details (masked)' && !ctx.isAdmin && v) ? String(v).replace(/.(?=.{4})/g, '*') : v;
  });
  return out;
}

/* ============================================ 8. BOOTSTRAP =================================== */

/** ONE call that boots the SPA: user/role, dropdown masters, lookups, form schema and KPIs. */
function getBootstrapData() {
  try {
    var ctx = userCtx_();
    var s = getSettings_();
    var allowed = scopeProjectsFor_(ctx);
    // lookups for pickers/filters
    var pd = readData_('Projects');
    var projects = [];
    pd.rows.forEach(function (r) {
      var o = rowToObj_(pd, r);
      if (String(o['Deleted'] || '') === 'Yes') return;
      if (allowed && allowed.indexOf(String(o['Project ID'] || '')) < 0) return;
      projects.push({id: String(o['Project ID'] || ''), code: String(o['Project Code'] || ''),
                     name: String(o['Project Name'] || ''), city: String(o['City'] || ''),
                     status: String(o['Status'] || ''), pm: String(o['Project Manager'] || '')});
    });
    var cd = readData_('Clients'), clients = [];
    cd.rows.forEach(function (r) { var o = rowToObj_(cd, r); if (o['Client ID']) clients.push({id: String(o['Client ID']), name: String(o['Name'] || '')}); });
    var vd = readData_('Vendors'), vendors = [];
    vd.rows.forEach(function (r) { var o = rowToObj_(vd, r); if (o['Vendor ID']) vendors.push({id: String(o['Vendor ID']), name: String(o['Vendor Name'] || '')}); });
    // schema meta for generic tables/forms (server is the single source of truth)
    var schema = {};
    Object.keys(SCHEMA).forEach(function (name) {
      schema[name] = SCHEMA[name].cols.map(function (c) {
        return {k: c.k, t: c.t, req: !!c.req, ro: !!c.ro || !!c.calc, cat: c.cat || '', sens: !!c.sens,
                min: (c.min === undefined ? null : c.min), max: (c.max === undefined ? null : c.max), w: c.w || 0};
      });
    });
    var kpis = {};
    try { kpis = buildDashboard_({}).kpis; } catch (e) { logSystem_('WARN', 'bootstrap kpis', e && e.message, ''); }
    return {
      ok: true,
      user: {email: ctx.email, name: ctx.name, role: ctx.role, isAdmin: ctx.isAdmin,
             seeCost: ctx.seeCost, canWrite: ctx.canWrite, canDelete: ctx.canDelete},
      settings: {cats: s.cats, users: s.users},
      settingCats: settingCats_(),
      lookups: {projects: projects, clients: clients, vendors: vendors},
      schema: schema,
      sheetOrder: SHEET_ORDER,
      kpis: kpis,
      meta: {appName: APP.name, today: todayStr_(), pageSize: APP.pageSize, docFolders: APP.docFolderNames}
    };
  } catch (e) {
    logSystem_('ERROR', 'getBootstrapData', e && e.message, e && e.stack);
    throw new Error(friendly_(e));
  }
}

/* ============================================ 9. GENERIC CRUD ENGINE ========================= */

/** Server-side search / filter / sort / pagination. p = {sheet, page, search, filters, sort, all} */
function listRecords(p) {
  try {
    p = p || {};
    var ctx = userCtx_();
    var name = p.sheet;
    var sc = SCHEMA[name];
    if (!sc) throw new Error('Unknown register: ' + name);
    if (name === 'Cost Tracker' && !ctx.seeCost) throw new Error('Access denied: cost data is restricted to Admin and Project Director roles.');
    var d = readData_(name);
    var objs = d.rows.map(function (r) { return rowToObj_(d, r); });
    if (sc.softDelete) objs = objs.filter(function (o) { return String(o['Deleted'] || '') !== 'Yes'; });
    var allowed = scopeProjectsFor_(ctx);
    if (allowed && sc.projectKey) objs = objs.filter(function (o) { return allowed.indexOf(String(o[sc.projectKey] || '')) >= 0; });
    if (p.filters) {
      Object.keys(p.filters).forEach(function (k) {
        var v = p.filters[k];
        if (v === '' || v === null || v === undefined) return;
        var arr = Object.prototype.toString.call(v) === '[object Array]' ? v.map(function (x) { return String(x).toLowerCase(); }) : null;
        objs = objs.filter(function (o) {
          var ov = normOut_(o[k]).toLowerCase();
          return arr ? arr.indexOf(ov) >= 0 : ov === String(v).toLowerCase();
        });
      });
    }
    if (p.search) {
      var q = String(p.search).toLowerCase();
      objs = objs.filter(function (o) {
        for (var i = 0; i < sc.cols.length; i++) {
          var c = sc.cols[i];
          if (c.hidden) continue;
          if (normOut_(o[c.k]).toLowerCase().indexOf(q) >= 0) return true;
        }
        return false;
      });
    }
    // Projects/Clients/Vendors: refresh computed rollups in memory so lists are always correct
    if (name === 'Projects' || name === 'Clients' || name === 'Vendors') {
      var cx = calcCtx_();
      objs = objs.map(function (o) { return Object.assign({}, o, computeRow_(name, o, cx, true, ctx)); });
    }
    var sort = p.sort || sc.ds || {col: sc.idKey, dir: 'asc'};
    if (sort && sort.col) {
      var dir = sort.dir === 'desc' ? -1 : 1;
      objs.sort(function (a, b) {
        var av = normOut_(a[sort.col]), bv = normOut_(b[sort.col]);
        var an = parseFloat(String(av).replace(/,/g, '')), bn = parseFloat(String(bv).replace(/,/g, ''));
        if (!isNaN(an) && !isNaN(bn) && /^[\d,.\-\s₹%]+$/.test(av) && av !== '') { if (an < bn) return -1 * dir; if (an > bn) return 1 * dir; return 0; }
        if (av < bv) return -1 * dir; if (av > bv) return 1 * dir; return 0;
      });
    }
    var total = objs.length;
    var pageSize = Math.min(Math.max(1, p.pageSize || APP.pageSize), 500);
    var pages = Math.max(1, Math.ceil(total / pageSize));
    var page = Math.min(Math.max(1, p.page || 1), pages);
    if (p.all) return {rows: objs.map(function (o) { return sanitizeOut_(ctx, name, o); }), total: total};
    var rows = objs.slice((page - 1) * pageSize, page * pageSize).map(function (o) { return sanitizeOut_(ctx, name, o); });
    return {rows: rows, total: total, page: page, pages: pages, pageSize: pageSize, role: ctx.role};
  } catch (e) {
    logSystem_('ERROR', 'listRecords ' + (p && p.sheet), e && e.message, e && e.stack);
    throw new Error(friendly_(e));
  }
}

/** Create/update with validation, duplicate checks, LockService, audit and rollup recompute. */
function saveRecord(p) {
  try {
    var ctx = userCtx_();
    var name = p.sheet, sc = SCHEMA[name];
    if (!sc) throw new Error('Unknown register: ' + name);
    if (sc.readOnly) throw new Error('This register is system-managed and cannot be edited from the app.');
    if (name === 'Cost Tracker' && !ctx.seeCost) throw new Error('Access denied: cost data is restricted to Admin and Project Director roles.');
    if (name === 'Settings' && !ctx.isAdmin) throw new Error('Access denied: only Admin users can manage Settings.');
    if (!canWrite_(ctx)) throw new Error('Access denied: your role (' + ctx.role + ') is read-only.');
    var rec = cleanRecord_(name, p.record || {}, ctx);
    var result = withLock_(function () {
      var d = readData_(name);
      var existing = null, idx = -1;
      var idk = sc.idKey;
      if (rec[idk]) {
        for (var i = 0; i < d.rows.length; i++) {
          if (String(d.rows[i][d.hm[idk]]) === String(rec[idk])) { idx = i; existing = rowToObj_(d, d.rows[i]); break; }
        }
        if (idx < 0) throw new Error('Record ' + rec[idk] + ' was not found. It may have been deleted — refresh and retry.');
      }
      checkScope_(ctx, name, rec, existing);
      var merged, created = false;
      if (idx < 0) { created = true; rec[idk] = rec[idk] || nextId_(name); merged = Object.assign({}, rec); }
      else merged = Object.assign({}, existing, rec);
      // validate the FULL merged record (partial update payloads inherit valid stored values)
      validateRecord_(name, merged, ctx);
      sc.cols.forEach(function (c) { // duplicate prevention on unique columns
        if (!c.uniq) return;
        var v = String(merged[c.k] || '').toLowerCase();
        if (!v) return;
        for (var i = 0; i < d.rows.length; i++) {
          if (i === idx) continue;
          if (String(d.rows[i][d.hm[c.k]] || '').toLowerCase() === v)
            throw new Error('Duplicate "' + c.k + '": "' + merged[c.k] + '" already exists.');
        }
      });
      merged = Object.assign({}, merged, computeRow_(name, merged, calcCtx_(), false, ctx));
      var rowArr = objToRow_(name, merged);
      var at = created ? Math.max(d.sheet.getLastRow() + 1, 2) : idx + 2;
      d.sheet.getRange(at, 1, 1, sc.cols.length).setValues([rowArr]);
      invalidate_(name);
      if (name === 'Settings' || name === 'Clients' || name === 'Projects') { _m.scope = null; }
      var diffs = [], base = created ? {} : existing;
      sc.cols.forEach(function (c) {
        if (c.hidden) return;
        var b = normOut_(base[c.k] === undefined ? '' : base[c.k]), n = normOut_(merged[c.k] === undefined ? '' : merged[c.k]);
        if (String(b) !== String(n)) diffs.push({f: c.k, o: b, n: n});
      });
      audit_(created ? 'CREATE' : 'UPDATE', name, merged[sc.idKey], diffs, ctx);
      if (created && name === 'Projects') autoDocChecklist_(merged, ctx);
      return {record: sanitizeOut_(ctx, name, merged), created: created};
    });
    var pk = SCHEMA[name].projectKey;
    if (name === 'Settings') invalidateSettingsCache_();
    if (pk && name !== 'Projects') { try { recomputeProject_(result.record[pk]); } catch (e) { logSystem_('WARN', 'recompute', e && e.message, ''); } }
    return result;
  } catch (e) {
    logSystem_('ERROR', 'saveRecord ' + (p && p.sheet), e && e.message, e && e.stack);
    throw new Error(friendly_(e));
  }
}

/** Soft delete for Projects (flag), hard delete elsewhere — always audited. */
function deleteRecord(p) {
  try {
    var ctx = userCtx_();
    var name = p.sheet, sc = SCHEMA[name];
    if (!sc) throw new Error('Unknown register: ' + name);
    if (sc.readOnly) throw new Error('This register is system-managed and cannot be deleted from the app.');
    if (name === 'Cost Tracker' && !ctx.seeCost) throw new Error('Access denied: cost data is restricted.');
    if (name === 'Settings' && !ctx.isAdmin) throw new Error('Access denied: only Admin users can manage Settings.');
    if (!ctx.canWrite) throw new Error('Access denied: your role (' + ctx.role + ') is read-only.');
    var id = String(p.id || '');
    if (!id) throw new Error('Validation: missing record id.');
    var out = withLock_(function () {
      var d = readData_(name);
      var idx = -1, obj = null;
      for (var i = 0; i < d.rows.length; i++) { if (String(d.rows[i][d.hm[sc.idKey]]) === id) { idx = i; obj = rowToObj_(d, d.rows[i]); break; } }
      if (idx < 0) throw new Error('Record ' + id + ' was not found. It may already be deleted.');
      if (!canDelete_(ctx, name, obj)) throw new Error('Access denied: you cannot delete this record.');
      var pid = sc.projectKey && name !== 'Projects' ? String(obj[sc.projectKey] || '') : '';
      if (sc.softDelete) {
        d.sheet.getRange(idx + 2, d.hm['Deleted'] + 1).setValue('Yes');
      } else {
        d.sheet.deleteRow(idx + 2);
      }
      invalidate_(name);
      audit_('DELETE' + (sc.softDelete ? '(soft)' : ''), name, id, [{f: sc.idKey, o: id, n: ''}], ctx);
      return pid;
    });
    if (out) { try { recomputeProject_(out); } catch (e) {} }
    if (name === 'Projects') { _m.scope = null; }
    return {ok: true, id: id, soft: !!sc.softDelete};
  } catch (e) {
    logSystem_('ERROR', 'deleteRecord ' + (p && p.sheet), e && e.message, e && e.stack);
    throw new Error(friendly_(e));
  }
}

/** Strips unknown/computed/sensitive fields from the client payload, coerces types. */
function cleanRecord_(name, rec, ctx) {
  var sc = SCHEMA[name], out = {};
  sc.cols.forEach(function (c) {
    if (c.hidden || c.calc) return;             // computed server-side, never trusted from client
    if (c.t === 'id') { out[c.k] = rec[c.k] !== undefined ? String(rec[c.k] || '') : undefined; return; }
    if (c.sens && !ctx.seeCost) return;
    var v = rec[c.k];
    if (v === undefined) return;
    v = normVal_(v);
    if (c.t === 'number') {
      out[c.k] = (v === '' || v === null) ? '' : num_(v);
      if (c.cat === 'GST Rate') out[c.k] = out[c.k] === '' ? '' : num_(out[c.k]);
    } else if (c.t === 'yn') {
      var s = String(v).trim();
      out[c.k] = (s.toLowerCase() === 'no') ? 'No' : (s.toLowerCase() === 'yes' ? 'Yes' : '');
    } else {
      out[c.k] = String(v).trim();
    }
  });
  return out;
}

/** Server-side validation: required, numeric ≥ bounds, date format, dropdown membership, FKs. */
function validateRecord_(name, rec, ctx) {
  var sc = SCHEMA[name];
  sc.cols.forEach(function (c) {
    if (c.hidden || c.calc) return;
    var v = rec[c.k];
    if (c.req && (v === undefined || v === '' || v === null))
      throw new Error('Validation: "' + c.k + '" is required.');
    if (v === '' || v === undefined || v === null) return;
    if (c.t === 'number') {
      var n = num_(v);
      if (isNaN(n)) throw new Error('Validation: "' + c.k + '" must be a number.');
      var min = c.min !== undefined && c.min !== null ? c.min : 0;
      if (n < min) throw new Error('Validation: "' + c.k + '" must be ≥ ' + min + '.');
      if (c.max !== undefined && c.max !== null && n > c.max) throw new Error('Validation: "' + c.k + '" must be ≤ ' + c.max + '.');
    }
    if (c.t === 'date' && !/^\d{4}-\d{2}-\d{2}$/.test(String(v)))
      throw new Error('Validation: "' + c.k + '" must be a valid date (YYYY-MM-DD).');
    if (c.t === 'yn' && ['Yes', 'No'].indexOf(String(v)) < 0)
      throw new Error('Validation: "' + c.k + '" must be Yes or No.');
    if (c.t === 'email' && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v)))
      throw new Error('Validation: "' + c.k + '" must be a valid email address.');
    if (c.t === 'select') {
      var opts = settingsList_(c.cat);
      if (opts.length && opts.indexOf(String(v)) < 0)
        throw new Error('Validation: "' + String(v) + '" is not in the "' + c.cat + '" list. Add it under Settings ▸ Dropdown Lists.');
    }
    if (c.t === 'fkproj') {
      var ids = projectIds_();
      if (ids.indexOf(String(v)) < 0) throw new Error('Validation: Project "' + v + '" was not found.');
    }
    if (c.t === 'fkclient') {
      var cds = readData_('Clients'), found = false;
      for (var i = 0; i < cds.rows.length; i++) { if (String(cds.rows[i][cds.hm['Client ID']]) === String(v)) { found = true; break; } }
      if (!found) throw new Error('Validation: Client "' + v + '" was not found.');
    }
    if (c.t === 'fkvendor') {
      var names = settingsVendorNames_();
      if (names.indexOf(String(v)) < 0) throw new Error('Validation: Vendor "' + v + '" was not found in the Vendors register.');
    }
  });
}

function projectIds_() {
  var d = readData_('Projects'), out = [], iId = d.hm['Project ID'], iDel = d.hm['Deleted'];
  d.rows.forEach(function (r) { if (iDel < 0 || String(r[iDel]) !== 'Yes') out.push(String(r[iId] || '')); });
  return out;
}
function settingsVendorNames_() {
  var d = readData_('Vendors'), out = [], iN = d.hm['Vendor Name'];
  d.rows.forEach(function (r) { var v = String(r[iN] || '').trim(); if (v) out.push(v); });
  return out;
}

/** Auto ID with LockService: FOI-YYYY-001 · PO- · RA- · VO- · DPR- · SNG- · CST- · RSK- … */
function nextId_(name) {
  var sc = SCHEMA[name];
  var d = readData_(name);
  var i = d.hm[sc.idKey];
  if (i === undefined) return sc.idPrefix + '-0001';
  var yr = Utilities.formatDate(new Date(), TZ_(), 'yyyy');
  var pre = sc.idPrefix + '-' + yr + '-';
  var max = 0;
  d.rows.forEach(function (r) {
    var v = String(r[i] || '');
    if (v.indexOf(pre) === 0) { var n = parseInt(v.slice(pre.length), 10); if (!isNaN(n) && n > max) max = n; }
  });
  return pre + pad_(max + 1, 3);
}

/* ============================================ 10. DERIVED FIELDS & ROLLUPS =================== */

/** Cross-sheet rollup cache (read once per execution): billed/received/retention per project
 *  from Billing, cost from Cost Tracker. Recomputed automatically when those sheets change. */
function getFinSums_() {
  if (_m.fin) return _m.fin;
  var fin = {};
  var bd = readData_('Billing');
  var iP = bd.hm['Project ID'], iN = bd.hm['Net Receivable'], iR = bd.hm['Received Amount'], iRet = bd.hm['Retention Amount'];
  bd.rows.forEach(function (r) {
    var p = String(r[iP] || ''); if (!p) return;
    var f = fin[p] || (fin[p] = {billed: 0, received: 0, retention: 0, cost: 0});
    f.billed += num_(iN >= 0 ? r[iN] : 0);
    f.received += num_(iR >= 0 ? r[iR] : 0);
    f.retention += num_(iRet >= 0 ? r[iRet] : 0);
  });
  try {
    var cd = readData_('Cost Tracker');
    var cP = cd.hm['Project ID'], cT = cd.hm['Total Amount'];
    cd.rows.forEach(function (r) {
      var p = String(r[cP] || ''); if (!p) return;
      var f = fin[p] || (fin[p] = {billed: 0, received: 0, retention: 0, cost: 0});
      f.cost += num_(cT >= 0 ? r[cT] : 0);
    });
  } catch (e) { /* viewer contexts never reach here; non-privileged saves don't need cost */ }
  _m.fin = fin;
  return fin;
}

function clientsMap_() {
  var d = readData_('Clients'), out = {};
  var iId = d.hm['Client ID'], iN = d.hm['Name'], iT = d.hm['Payment Terms (days)'];
  d.rows.forEach(function (r) {
    out[String(r[iId] || '')] = {name: String(r[iN] || ''), terms: num_(iT >= 0 ? r[iT] : 30) || 30};
  });
  return out;
}

function calcCtx_() {
  if (_m.calcCtx) return _m.calcCtx;
  _m.calcCtx = {today: todayStr_(), fin: getFinSums_(), clients: clientsMap_(),
                live: liveStatuses_(), email: userCtx_().email};
  return _m.calcCtx;
}

/**
 * Computes every derived column per sheet. Pure function of the record (+ cached rollups),
 * used for both persisted rows (memOnly=false) and in-memory list refreshes (memOnly=true).
 * Projects RAG rule: Red if delay>15d or variance<-15% · Amber if delay>0 or variance<-5% · else Green.
 */
function computeRow_(name, rec, cx, memOnly, ctx) {
  var o = {};
  if (name === 'Projects') {
    var gst = num_(rec['GST %']), cv = num_(rec['Contract Value (excl GST)']);
    o['GST Amount'] = cv ? round1_(cv * gst / 100) : '';
    o['Total Contract Value'] = cv ? round1_(cv + num_(o['GST Amount'])) : '';
    var start = String(rec['Start Date'] || ''), pend = String(rec['Revised End'] || rec['Planned End'] || '');
    var aend = String(rec['Actual End'] || ''), status = String(rec['Status'] || '');
    var terminal = APP.terminalStatuses.indexOf(status) >= 0;
    o['Duration (days)'] = (start && pend) ? diffDays_(start, pend) : '';
    o['Days Remaining'] = (!terminal && pend) ? diffDays_(cx.today, pend) : '';
    var delay = 0;
    if (terminal && aend && pend) delay = Math.max(0, diffDays_(pend, aend));
    else if (!terminal && pend) delay = Math.max(0, -num_(o['Days Remaining']));
    o['Delay Days'] = delay;
    var plan = num_(rec['Planned %']), act = num_(rec['Actual %']);
    o['Variance %'] = round1_(act - plan);
    o['RAG'] = (delay > 15 || num_(o['Variance %']) < -15) ? 'Red' : ((delay > 0 || num_(o['Variance %']) < -5) ? 'Amber' : 'Green');
    var f = cx.fin[String(rec['Project ID'] || '')] || {billed: 0, received: 0, retention: 0, cost: 0};
    o['Billed'] = round1_(f.billed);
    o['Received'] = round1_(f.received);
    o['Outstanding'] = round1_(Math.max(0, f.billed - f.received));
    o['Retention Held'] = round1_(f.retention);
    o['Cost Incurred'] = round1_(f.cost);
    var budget = num_(rec['Approved Budget']);
    o['Cost-to-Complete'] = budget ? round1_(Math.max(0, budget - f.cost)) : '';
    o['Projected Margin %'] = budget ? round1_((budget - f.cost) / budget * 100) : '';
    o['Client Name'] = '';
    var cli = cx.clients[String(rec['Client ID'] || '')];
    if (cli) o['Client Name'] = cli.name;
    if (!memOnly) { o['Last Updated'] = nowStr_(); o['Updated By'] = (ctx && ctx.email) || cx.email || 'system'; }
    return o;
  }
  if (name === 'Clients') {
    var n = 0, iId = String(rec['Client ID'] || '');
    if (iId) {
      var pd = readData_('Projects'), iP = pd.hm['Client ID'], iS = pd.hm['Status'], iDel = pd.hm['Deleted'];
      pd.rows.forEach(function (r) {
        if (iDel >= 0 && String(r[iDel]) === 'Yes') return;
        if (String(r[iP] || '') === iId && APP.terminalStatuses.indexOf(String(r[iS] || '')) < 0) n++;
      });
    }
    o['Active Projects'] = n;
    return o;
  }
  if (name === 'Vendors') {
    var wos = 0, biz = 0, vname = String(rec['Vendor Name'] || '');
    var dd = readData_('Procurement');
    var vI = dd.hm['Vendor'], vS = dd.hm['Status'], vT = dd.hm['Total Amount'];
    dd.rows.forEach(function (r) {
      if (String(r[vI] || '') === vname) {
        if (['Closed', 'Cancelled'].indexOf(String(r[vS] || '')) < 0) wos++;
        biz += num_(vT >= 0 ? r[vT] : 0);
      }
    });
    o['Active WOs'] = wos;
    o['Total Business Value'] = round1_(biz);
    return o;
  }
  if (name === 'BOQ') {
    var qty = num_(rec['Qty']), cr = num_(rec['Client Rate']), br = num_(rec['Budget Rate']), qe = num_(rec['Qty Executed']);
    o['Client Amount'] = round1_(qty * cr);
    o['Budget Amount'] = round1_(qty * br);
    var pct = qty > 0 ? Math.min(100, qe / qty * 100) : 0;
    o['% Complete'] = qty > 0 ? round1_(pct) : '';
    o['Executed Value'] = round1_(num_(o['Budget Amount']) * pct / 100);
    return o;
  }
  if (name === 'Procurement') {
    var q = num_(rec['Qty']), rate = num_(rec['Rate']), g = num_(rec['GST %']);
    var basic = round1_(q * rate), gstA = round1_(basic * g / 100);
    o['Basic Amount'] = basic;
    o['GST Amount'] = gstA;
    o['Total Amount'] = round1_(basic + gstA);
    var req = String(rec['Required By'] || ''), del = String(rec['Delivery Date'] || '');
    o['Delay Days'] = (req && del) ? Math.max(0, diffDays_(req, del)) : 0;
    o['Balance Amount'] = round1_(Math.max(0, num_(o['Total Amount']) - num_(rec['Paid'])));
    return o;
  }
  if (name === 'Cost Tracker') {
    var amt = num_(rec['Amount']), cg = num_(rec['GST %']);
    o['GST Amount'] = round1_(amt * cg / 100);
    o['Total Amount'] = round1_(amt + num_(o['GST Amount']));
    return o;
  }
  if (name === 'Billing') {
    var gross = num_(rec['Gross Amount']), rp = num_(rec['Retention %']), bg = num_(rec['GST %']);
    var ret = round1_(gross * rp / 100);
    var gstB = round1_((gross - ret) * bg / 100);
    o['Retention Amount'] = ret;
    o['GST Amount'] = gstB;
    o['Net Receivable'] = round1_(gross - ret + gstB);
    var sub = String(rec['Submitted Date'] || String(rec['Bill Date'] || ''));
    var terms = 30;
    var cliB = cx.clients[String(rec['Client ID'] || rec['__clientId'] || '')];
    var projCli = null;
    if (!cliB) {
      var prd = readData_('Projects');
      for (var i = 0; i < prd.rows.length; i++) {
        var po = rowToObj_(prd, prd.rows[i]);
        if (String(po['Project ID'] || '') === String(rec['Project ID'] || '')) { projCli = cx.clients[String(po['Client ID'] || '')]; break; }
      }
    }
    if (projCli) terms = projCli.terms;
    o['Due Date'] = sub ? Utilities.formatDate(new Date(Date.parse(sub + 'T00:00:00Z') + terms * 86400000), TZ_(), 'yyyy-MM-dd') : '';
    var ramt = num_(rec['Received Amount']);
    o['Outstanding'] = round1_(Math.max(0, num_(o['Net Receivable']) - ramt));
    o['Ageing Days'] = (sub && num_(o['Outstanding']) > 0) ? Math.max(0, diffDays_(sub, cx.today)) : 0;
    return o;
  }
  if (name === 'Drawings & Approvals') {
    var subd = String(rec['Submitted Date'] || '');
    o['Pending Days'] = (subd && String(rec['Approval Status'] || '') === 'Pending') ? Math.max(0, diffDays_(subd, cx.today)) : 0;
    return o;
  }
  if (name === 'Risks & Issues') {
    o['Score'] = num_(rec['Probability (1-5)']) * num_(rec['Impact (1-5)']);
    return o;
  }
  return o;
}

/** Rewrites the rollup columns of one project row (called after transaction changes). */
function recomputeProject_(projectId) {
  if (!projectId) return;
  withLock_(function () {
    var d = readData_('Projects'), sc = SCHEMA['Projects'];
    for (var i = 0; i < d.rows.length; i++) {
      if (String(d.rows[i][d.hm['Project ID']]) !== String(projectId)) continue;
      var obj = rowToObj_(d, d.rows[i]);
      var merged = Object.assign({}, obj, computeRow_('Projects', obj, calcCtx_(), false, userCtx_()));
      d.sheet.getRange(i + 2, 1, 1, sc.cols.length).setValues([objToRow_('Projects', merged)]);
      invalidate_('Projects');
      return;
    }
  });
}

/** A new project automatically gets its full Documents Checklist pre-created. */
function autoDocChecklist_(proj, ctx) {
  try {
    var types = settingsList_('Doc Type');
    if (!types.length) return;
    var d = readData_('Documents Checklist'), sc = SCHEMA['Documents Checklist'];
    var rows = [];
    types.forEach(function (t) {
      rows.push([nextId_('Documents Checklist'), proj['Project ID'], t, 'Yes', 'No', '', '', 'Pending']);
      invalidate_('Documents Checklist');
    });
    d.sheet.getRange(Math.max(d.sheet.getLastRow() + 1, 2), 1, rows.length, sc.cols.length).setValues(rows);
    invalidate_('Documents Checklist');
    audit_('CREATE', 'Documents Checklist', proj['Project ID'], [{f: 'Checklist', o: '', n: types.length + ' doc types seeded'}], ctx);
  } catch (e) { logSystem_('WARN', 'autoDocChecklist_', e && e.message, ''); }
}

/* ============================================ 11. PROJECT 360 ================================ */

function getProjectDetail(projectId) {
  try {
    var ctx = userCtx_();
    var d = readData_('Projects');
    var proj = null;
    for (var i = 0; i < d.rows.length; i++) {
      if (String(d.rows[i][d.hm['Project ID']]) === String(projectId)) { proj = rowToObj_(d, d.rows[i]); break; }
    }
    if (!proj) throw new Error('Project ' + projectId + ' was not found.');
    var allowed = scopeProjectsFor_(ctx);
    if (allowed && allowed.indexOf(String(projectId)) < 0) throw new Error('Access denied: this project is not assigned to you.');
    proj = Object.assign({}, proj, computeRow_('Projects', proj, calcCtx_(), true, ctx));
    proj = sanitizeOut_(ctx, 'Projects', proj);
    var pid = String(projectId);
    function sumIn_(sheet, key, valKey, filterFn) {
      var dd = readData_(sheet), tot = 0, n = 0, iK = dd.hm[key], iV = dd.hm[valKey];
      dd.rows.forEach(function (r) {
        if (String(r[iK] || '') !== pid) return;
        if (filterFn && !filterFn(rowToObj_(dd, r))) return;
        n++; tot += num_(iV >= 0 ? r[iV] : 0);
      });
      return {count: n, total: round1_(tot)};
    }
    var boq = sumIn_('BOQ', 'Project ID', 'Budget Amount');
    var boqEx = sumIn_('BOQ', 'Project ID', 'Executed Value');
    var po = sumIn_('Procurement', 'Project ID', 'Total Amount');
    var bills = sumIn_('Billing', 'Project ID', 'Net Receivable');
    var stats = {
      boq: {lines: boq.count, budget: boq.total, executed: boqEx.total, pct: boq.total ? round1_(boqEx.total / boq.total * 100) : 0},
      po: po,
      billing: {count: bills.count, gross: bills.total},
      snagsOpen: sumIn_('Snags & Handover', 'Project ID', 'Snag ID', function (o) { return ['Open', 'In Progress'].indexOf(String(o['Status'] || '')) >= 0; }).count,
      risksOpen: sumIn_('Risks & Issues', 'Project ID', 'ID', function (o) { return String(o['Status'] || '') !== 'Closed'; }).count,
      drawingsPending: sumIn_('Drawings & Approvals', 'Project ID', 'Drawing No', function (o) { return String(o['Approval Status'] || '') === 'Pending'; }).count,
      voPending: sumIn_('Variation Orders', 'Project ID', 'Amount', function (o) { return String(o['Approval Status'] || '') === 'Pending'; })
    };
    var fin = ctx.seeCost ? (getFinSums_()[pid] || {billed: 0, received: 0, retention: 0, cost: 0}) : null;
    return {project: proj, stats: stats, fin: fin};
  } catch (e) {
    logSystem_('ERROR', 'getProjectDetail', e && e.message, e && e.stack);
    throw new Error(friendly_(e));
  }
}

/* ============================================ 12. DASHBOARD & NOTIFICATIONS ================== */

/** Full dashboard payload: KPIs, chart aggregates and watch tables. f = {city, pm, status, type} */
function buildDashboard_(f) {
  var ctx = userCtx_();
  f = f || {};
  var cx = calcCtx_();
  var d = readData_('Projects');
  var allowed = scopeProjectsFor_(ctx);
  var projs = [];
  d.rows.forEach(function (r) {
    var o = rowToObj_(d, r);
    if (String(o['Deleted'] || '') === 'Yes') return;
    if (allowed && allowed.indexOf(String(o['Project ID'] || '')) < 0) return;
    projs.push(Object.assign({}, o, computeRow_('Projects', o, cx, true, ctx)));
  });
  if (f.city) projs = projs.filter(function (o) { return String(o['City'] || '') === f.city; });
  if (f.pm) projs = projs.filter(function (o) { return String(o['Project Manager'] || '') === f.pm; });
  if (f.status) projs = projs.filter(function (o) { return String(o['Status'] || '') === f.status; });
  if (f.type) projs = projs.filter(function (o) { return String(o['Project Type'] || '') === f.type; });

  var liveSet = {};
  cx.live.forEach(function (s) { liveSet[s] = 1; });
  var isLive = function (o) { return !!liveSet[String(o['Status'] || '')]; };

  var totalVal = 0, billed = 0, received = 0, outstanding = 0, live = 0, delayed = 0, red = 0, margins = [], liveArr = [];
  projs.forEach(function (o) {
    totalVal += num_(o['Total Contract Value']);
    billed += num_(o['Billed']);
    received += num_(o['Received']);
    outstanding += num_(o['Outstanding']);
    if (isLive(o)) { live++; liveArr.push(o); if (num_(o['Delay Days']) > 0) delayed++; if (String(o['RAG']) === 'Red') red++; }
    if (num_(o['Approved Budget']) > 0) margins.push(num_(o['Projected Margin %']));
  });
  var kpis = {
    totalProjects: projs.length,
    liveProjects: live,
    totalValue: round1_(totalVal),
    billed: round1_(billed),
    received: round1_(received),
    outstanding: round1_(outstanding),
    avgMargin: (ctx.seeCost && margins.length) ? round1_(margins.reduce(function (a, b) { return a + b; }, 0) / margins.length) : null,
    delayed: delayed,
    redFlags: red,
    collectionPct: billed > 0 ? round1_(received / billed * 100) : 0
  };

  // ── chart aggregates ──────────────────────────────────────────────────────────────────────
  var byStatus = {}, byCity = {};
  projs.forEach(function (o) {
    var s = String(o['Status'] || 'Unknown');
    byStatus[s] = (byStatus[s] || 0) + 1;
    var c = String(o['City'] || 'Unknown');
    byCity[c] = (byCity[c] || 0) + num_(o['Total Contract Value']);
  });
  var statusChart = Object.keys(byStatus).map(function (k) { return {label: k, value: byStatus[k]}; }).sort(function (a, b) { return b.value - a.value; });
  var cityChart = Object.keys(byCity).map(function (k) { return {label: k, value: round1_(byCity[k])}; }).sort(function (a, b) { return b.value - a.value; });

  // monthly billing vs collection (last 12 months)
  var months = [], mk = {};
  var now = new Date();
  for (var m = 11; m >= 0; m--) {
    var dt = new Date(now.getFullYear(), now.getMonth() - m, 1);
    var key = Utilities.formatDate(dt, TZ_(), 'yyyy-MM');
    mk[key] = {billed: 0, received: 0, label: Utilities.formatDate(dt, TZ_(), 'MMM yy')};
    months.push(key);
  }
  var bd = readData_('Billing'), allowedS = scopeProjectsFor_(ctx);
  var iP = bd.hm['Project ID'], iBD = bd.hm['Bill Date'], iN = bd.hm['Net Receivable'], iRD = bd.hm['Received Date'], iR = bd.hm['Received Amount'];
  var billKeys = {};
  bd.rows.forEach(function (r) {
    var o = rowToObj_(bd, r);
    if (allowedS && allowedS.indexOf(String(o['Project ID'] || '')) < 0) return;
    var b = String(o['Bill Date'] || '').slice(0, 7);
    if (mk[b] !== undefined) mk[b].billed += num_(o['Net Receivable']);
    var rc = String(o['Received Date'] || '').slice(0, 7);
    if (mk[rc] !== undefined) mk[rc].received += num_(o['Received Amount']);
  });
  var monthly = {labels: months.map(function (k) { return mk[k].label; }),
                 billed: months.map(function (k) { return round1_(mk[k].billed); }),
                 received: months.map(function (k) { return round1_(mk[k].received); })};

  // top outstanding clients
  var cliOut = {};
  projs.forEach(function (o) {
    var c = String(o['Client Name'] || String(o['Client ID'] || '—'));
    cliOut[c] = (cliOut[c] || 0) + num_(o['Outstanding']);
  });
  var topClients = Object.keys(cliOut).map(function (k) { return {label: k, value: round1_(cliOut[k])}; })
    .filter(function (x) { return x.value > 0; })
    .sort(function (a, b) { return b.value - a.value; }).slice(0, 10);

  // planned vs actual (top 10 live by value)
  var prog = liveArr.slice().sort(function (a, b) { return num_(b['Total Contract Value']) - num_(a['Total Contract Value']); }).slice(0, 10)
    .map(function (o) { return {name: String(o['Project Code'] || o['Project ID']), full: String(o['Project Name'] || ''), planned: num_(o['Planned %']), actual: num_(o['Actual %'])}; });

  // cost vs budget (top 8 live by budget) — privileged only
  var costBudget = null;
  if (ctx.seeCost) {
    costBudget = liveArr.filter(function (o) { return num_(o['Approved Budget']) > 0; })
      .sort(function (a, b) { return num_(b['Approved Budget']) - num_(a['Approved Budget']); }).slice(0, 8)
      .map(function (o) { return {name: String(o['Project Code'] || o['Project ID']), budget: num_(o['Approved Budget']), cost: num_(o['Cost Incurred'])}; });
  }

  // ── watch tables ──────────────────────────────────────────────────────────────────────────
  var delayedRows = liveArr.filter(function (o) { return num_(o['Delay Days']) > 0 || String(o['RAG']) === 'Red'; })
    .sort(function (a, b) { return num_(b['Delay Days']) - num_(a['Delay Days']); }).slice(0, 15)
    .map(function (o) {
      return {id: o['Project ID'], code: o['Project Code'], name: o['Project Name'], pm: o['Project Manager'],
              city: o['City'], rag: o['RAG'], delay: num_(o['Delay Days']), variance: num_(o['Variance %']),
              end: String(o['Revised End'] || o['Planned End'] || ''), status: o['Status']};
    });
  var milestones = projs.filter(function (o) {
    var md = String(o['Milestone Date'] || '');
    if (!md || APP.terminalStatuses.indexOf(String(o['Status'] || '')) >= 0) return false;
    var dd = diffDays_(cx.today, md);
    return dd >= -7 && dd <= 15;
  }).sort(function (a, b) { return diffDays_(cx.today, String(a['Milestone Date'])) - diffDays_(cx.today, String(b['Milestone Date'])); }).slice(0, 15)
    .map(function (o) {
      return {id: o['Project ID'], code: o['Project Code'], name: o['Project Name'], milestone: o['Next Milestone'],
              date: o['Milestone Date'], days: diffDays_(cx.today, String(o['Milestone Date'])), pm: o['Project Manager']};
    });
  var ageingRows = [], buckets = {b30: 0, b60: 0, b90: 0, b90p: 0, total: 0};
  bd.rows.forEach(function (r) {
    var o = rowToObj_(bd, r);
    if (allowedS && allowedS.indexOf(String(o['Project ID'] || '')) < 0) return;
    var out = num_(o['Outstanding']);
    if (out <= 0) return;
    var age = num_(o['Ageing Days']);
    if (age <= 30) buckets.b30++; else if (age <= 60) buckets.b60++; else if (age <= 90) buckets.b90++; else buckets.b90p++;
    buckets.total += out;
    ageingRows.push({bill: o['Bill No'], projectId: o['Project ID'], type: o['Bill Type'], net: num_(o['Net Receivable']),
                     outstanding: out, ageing: age, due: o['Due Date'], status: o['Status']});
  });
  ageingRows.sort(function (a, b) { return b.ageing - a.ageing; });
  var ageingTop = ageingRows.slice(0, 15);

  // dropdown sources for dashboard filters
  var pms = {}, types = {}, cities = {};
  projs.forEach(function (o) { if (o['Project Manager']) pms[o['Project Manager']] = 1; if (o['Project Type']) types[o['Project Type']] = 1; if (o['City']) cities[o['City']] = 1; });

  return {
    kpis: kpis,
    charts: {status: statusChart, city: cityChart, monthly: monthly, topClients: topClients, progress: prog, costBudget: costBudget},
    tables: {delayed: delayedRows, milestones: milestones, ageing: ageingTop},
    buckets: buckets,
    filterSources: {pms: Object.keys(pms).sort(), types: Object.keys(types).sort(), cities: Object.keys(cities).sort()},
    scopeNote: ctx.role === 'PM' ? 'Showing your assigned projects only' : ''
  };
}

function getDashboardData(f) {
  try { return buildDashboard_(f); }
  catch (e) { logSystem_('ERROR', 'getDashboardData', e && e.message, e && e.stack); throw new Error(friendly_(e)); }
}

/** Bell + 8 AM digest source: delayed projects, next milestones, 30/60/90 ageing, approvals. */
function getNotifications() {
  try {
    var ctx = userCtx_();
    var dash = buildDashboard_({});
    var items = [], n = 0;
    dash.tables.delayed.forEach(function (r) {
      items.push({type: 'delay', icon: 'alert', label: (r.code || r.id) + ' — ' + r.name, sub: 'Delayed ' + r.delay + ' days · RAG ' + r.rag, route: '#/projects/' + r.id});
    });
    dash.tables.milestones.forEach(function (r) {
      items.push({type: 'milestone', icon: 'flag', label: (r.code || r.id) + ' — ' + (r.milestone || 'Milestone'), sub: r.date + ' (' + (r.days >= 0 ? 'in ' + r.days + 'd' : Math.abs(r.days) + 'd overdue') + ')', route: '#/projects/' + r.id});
    });
    dash.tables.ageing.forEach(function (r) {
      if (r.ageing >= 30) items.push({type: 'bill', icon: 'receipt', label: r.bill + ' · ' + r.projectId, sub: 'Outstanding ' + fmtINRServer_(r.outstanding) + ' · ' + r.ageing + ' days', route: '#/m/billing'});
    });
    [['Variation Orders', 'VO No', 'vo'], ['Drawings & Approvals', 'Drawing No', 'drawing']].forEach(function (cfg) {
      if (!userCtx_().canWrite && cfg[2] === 'vo') return;
      var d = readData_(cfg[0]);
      var iS = d.hm['Approval Status'], iT = d.hm[cfg[1]], iP = d.hm['Project ID'];
      var c = 0;
      d.rows.forEach(function (r) {
        if (String(r[iS] || '') !== 'Pending') return;
        if (userCtx_().role === 'PM') { var al = scopeProjectsFor_(userCtx_()); if (al && al.indexOf(String(r[iP] || '')) < 0) return; }
        c++; if (c <= 8) items.push({type: 'approval', icon: 'check', label: String(r[iT] || '') + ' · ' + String(r[iP] || ''), sub: cfg[0] + ' pending approval', route: '#/m/' + (cfg[2] === 'vo' ? 'vo' : 'drawings')});
      });
    });
    return {
      counts: {delayed: dash.kpis.delayed, red: dash.kpis.redFlags, milestones: dash.tables.milestones.length,
               bills30: dash.buckets.b60 + dash.buckets.b90 + dash.buckets.b90p, approvals: items.filter(function (i) { return i.type === 'approval'; }).length},
      total: items.length, items: items.slice(0, 40)
    };
  } catch (e) { logSystem_('ERROR', 'getNotifications', e && e.message, e && e.stack); throw new Error(friendly_(e)); }
}

function fmtINRServer_(n) {
  n = num_(n);
  var a = Math.abs(n);
  if (a >= 1e7) return '₹' + (n / 1e7).toFixed(2) + ' Cr';
  if (a >= 1e5) return '₹' + (n / 1e5).toFixed(2) + ' L';
  return '₹' + Math.round(n).toLocaleString('en-IN');
}

/* ============================================ 13. DRIVE, PDF, EXPORT, EMAIL ================== */

/** Builds the standard 09-folder Drive structure for a project and saves the link. */
function createProjectFolder(projectId) {
  try {
    var ctx = userCtx_();
    if (!ctx.canWrite) throw new Error('Access denied: your role (' + ctx.role + ') is read-only.');
    var d = readData_('Projects');
    var proj = null, rowIdx = -1;
    for (var i = 0; i < d.rows.length; i++) {
      if (String(d.rows[i][d.hm['Project ID']]) === String(projectId)) { proj = rowToObj_(d, d.rows[i]); rowIdx = i; break; }
    }
    if (!proj) throw new Error('Project ' + projectId + ' was not found.');
    if (String(proj['Drive Folder Link'] || '')) return {link: proj['Drive Folder Link'], existed: true};
    return withLock_(function () {
      var parent = folder_(APP.projectsFolder);
      var sub = folder_((proj['Project Code'] || projectId) + ' — ' + (proj['Project Name'] || ''), parent);
      var links = [];
      APP.docFolderNames.forEach(function (f) { links.push(folder_(f, sub).getUrl()); });
      var url = sub.getUrl();
      d.sheet.getRange(rowIdx + 2, d.hm['Drive Folder Link'] + 1).setValue(url);
      invalidate_('Projects');
      audit_('UPDATE', 'Projects', projectId, [{f: 'Drive Folder Link', o: '', n: url}], ctx);
      return {link: url, existed: false, folders: links};
    });
  } catch (e) { logSystem_('ERROR', 'createProjectFolder', e && e.message, e && e.stack); throw new Error(friendly_(e)); }
}

function folder_(name, parent) {
  var it = (parent || DriveApp).getFoldersByName(name);
  return it.hasNext() ? it.next() : (parent ? parent.createFolder(name) : DriveApp.createFolder(name));
}

/** Branded one-page project summary PDF → saved into the project's Drive folder. Returns link. */
function generateProjectSummaryPDF(projectId) {
  try {
    var ctx = userCtx_();
    var det = getProjectDetail(projectId);
    var proj = det.project;
    var html = projectPdfHtml_(proj, det.stats, det.fin, ctx);
    var blob = Utilities.newBlob(html, 'text/html', 'summary.html').getAs('application/pdf')
      .setName('ProjectSummary_' + (proj['Project Code'] || projectId) + '_' + todayStr_() + '.pdf');
    var folder = null;
    var link = String(proj['Drive Folder Link'] || '');
    if (link) { try { folder = DriveApp.getFolderById(link.match(/[-\w]{25,}/)[0]); } catch (e) {} }
    if (!folder) folder = folder_(APP.projectsFolder);
    var file = folder.createFile(blob);
    return {link: file.getUrl(), name: file.getName()};
  } catch (e) { logSystem_('ERROR', 'generateProjectSummaryPDF', e && e.message, e && e.stack); throw new Error(friendly_(e)); }
}

function projectPdfHtml_(p, st, fin, ctx) {
  function r(n) { return (n === '' || n === undefined || n === null) ? '—' : '₹' + Number(n).toLocaleString('en-IN'); }
  function row(a, b) { return '<tr><td style="padding:5px 10px;color:#475569;width:38%">' + a + '</td><td style="padding:5px 10px;font-weight:600">' + b + '</td></tr>'; }
  var finRows = '';
  if (fin) {
    finRows = '<h2 style="font-size:13px;color:#4F46E5;margin:14px 0 4px">Financials</h2>'
      + '<table style="width:100%;border-collapse:collapse;font-size:11px">'
      + row('Approved Budget', r(p['Approved Budget'])) + row('Cost Incurred', r(fin.cost))
      + row('Cost-to-Complete', r(p['Cost-to-Complete'])) + row('Projected Margin %', (p['Projected Margin %'] || '—') + ' %')
      + row('Billed / Received', r(p['Billed']) + ' / ' + r(p['Received'])) + row('Outstanding', r(p['Outstanding']))
      + row('Retention Held', r(p['Retention Held'])) + '</table>';
  }
  return '<div style="font-family:Arial,sans-serif;color:#0F172A;padding:26px">'
    + '<div style="background:linear-gradient(90deg,#4F46E5,#7C3AED);color:#fff;padding:16px 20px;border-radius:10px">'
    + '<div style="font-size:11px;letter-spacing:2px;opacity:.85">FOI PROJECT MASTER REGISTER — ' + APP.short.toUpperCase() + '</div>'
    + '<div style="font-size:19px;font-weight:bold;margin-top:3px">' + esc_(p['Project Name']) + '</div>'
    + '<div style="font-size:11px;opacity:.9;margin-top:3px">' + esc_(p['Project Code']) + ' · ' + esc_(p['Project ID']) + ' · ' + esc_(p['City']) + ' · ' + esc_(p['Client Name']) + '</div></div>'
    + '<table style="width:100%;border-collapse:collapse;font-size:11px;margin-top:12px">'
    + row('Status / Phase', esc_(p['Status']) + ' / ' + esc_(p['Phase']))
    + row('RAG · Priority', esc_(p['RAG']) + ' · ' + esc_(p['Priority']))
    + row('Progress (Planned → Actual)', num_(p['Planned %']) + '% → <b>' + num_(p['Actual %']) + '%</b> (Var ' + num_(p['Variance %']) + '%)')
    + row('Schedule', esc_(p['Start Date']) + ' → ' + esc_(p['Revised End'] || p['Planned End']) + ' · Delay ' + num_(p['Delay Days']) + 'd · Remaining ' + esc_(p['Days Remaining']) + 'd')
    + row('Contract Value (incl GST)', r(p['Total Contract Value']))
    + row('Next Milestone', esc_(p['Next Milestone']) + ' ' + esc_(p['Milestone Date']))
    + row('Key Risk', esc_(p['Key Risk'] || '—'))
    + row('Team (PD / PM / Coordinator)', esc_(p['Project Director']) + ' / ' + esc_(p['Project Manager']) + ' / ' + esc_(p['Coordinator']))
    + '</table>'
    + '<h2 style="font-size:13px;color:#4F46E5;margin:14px 0 4px">Execution Snapshot</h2>'
    + '<table style="width:100%;border-collapse:collapse;font-size:11px">'
    + row('BOQ lines / Budget / Executed', st.boq.lines + ' · ' + r(st.boq.budget) + ' · ' + st.boq.pct + '%')
    + row('Procurement (PO/WO)', st.po.count + ' orders · ' + r(st.po.total))
    + row('Snags open / Risks open / Drawings pending', st.snagsOpen + ' / ' + st.risksOpen + ' / ' + st.drawingsPending)
    + '</table>' + finRows
    + '<div style="margin-top:16px;font-size:9px;color:#94A3B8">Generated ' + nowStr_() + ' by ' + esc_(ctx.email || 'system') + ' · ' + esc_(APP.name) + '</div></div>';
}

function esc_(s) { return String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

/** CSV of any current filtered view (all rows, server-side filtered). */
function exportCSV(p) {
  try {
    p = p || {}; p.all = true;
    var res = listRecords(p);
    var sc = SCHEMA[p.sheet];
    var cols = sc.cols.filter(function (c) { return !c.hidden && !(c.sens && !userCtx_().seeCost); }).map(function (c) { return c.k; });
    function q(v) { var s = v === undefined || v === null ? '' : String(v); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }
    var lines = [cols.map(q).join(',')];
    res.rows.forEach(function (o) { lines.push(cols.map(function (k) { return q(o[k]); }).join(',')); });
    return {name: p.sheet.replace(/[^\w]+/g, '_') + '_' + todayStr_() + '.csv', csv: lines.join('\n'), total: res.total};
  } catch (e) { logSystem_('ERROR', 'exportCSV', e && e.message, e && e.stack); throw new Error(friendly_(e)); }
}

function misRecipients_() {
  var to = [];
  usersList_().forEach(function (u) { if ((u.role === 'Admin' || u.role === 'Project Director') && u.email) to.push(u.email); });
  if (!to.length) { try { var o = Session.getEffectiveUser().getEmail(); if (o) to.push(o); } catch (e) {} }
  return to.join(',');
}

/** Weekly MIS (Mon 7 AM): KPI summary + delayed/red table as HTML body and PDF attachment. */
function emailWeeklyMIS() {
  try {
    var dash = buildDashboard_({});
    var html = misHtml_(dash, 'Weekly MIS');
    var pdf = Utilities.newBlob(html, 'text/html', 'mis.html').getAs('application/pdf').setName('FOI_Weekly_MIS_' + todayStr_() + '.pdf');
    var to = misRecipients_();
    if (to) MailApp.sendEmail({to: to, subject: 'FOI Weekly MIS — ' + todayStr_(), htmlBody: html, attachments: [pdf]});
    logSystem_('INFO', 'emailWeeklyMIS', 'sent to ' + to, '');
    return 'Weekly MIS emailed to ' + to;
  } catch (e) { logSystem_('ERROR', 'emailWeeklyMIS', e && e.message, e && e.stack); throw new Error(friendly_(e)); }
}

/** Daily digest (8 AM): delayed projects, due milestones, 30/60/90 outstanding, pending approvals. */
function dailyDigest() {
  try {
    var n = getNotifications();
    var k = buildDashboard_({}).kpis;
    var html = '<div style="font-family:Arial,sans-serif;color:#0F172A;max-width:640px">'
      + '<div style="background:linear-gradient(90deg,#4F46E5,#7C3AED);color:#fff;padding:14px 18px;border-radius:10px">'
      + '<b>FOI Daily Digest — ' + todayStr_() + '</b></div>'
      + '<p>Live: <b>' + k.liveProjects + '</b> · Delayed: <b>' + k.delayed + '</b> · Red: <b>' + k.redFlags
      + '</b> · Outstanding: <b>' + fmtINRServer_(k.outstanding) + '</b></p>';
    if (n.items.length) {
      html += '<ul style="font-size:13px;line-height:1.7">';
      n.items.slice(0, 25).forEach(function (i) { html += '<li><b>' + esc_(i.label) + '</b><br><span style="color:#64748B">' + esc_(i.sub) + '</span></li>'; });
      html += '</ul>';
    } else html += '<p>No alerts today.</p>';
    html += '</div>';
    var to = misRecipients_();
    if (to) MailApp.sendEmail({to: to, subject: 'FOI Daily Digest — ' + todayStr_(), htmlBody: html});
    logSystem_('INFO', 'dailyDigest', 'sent to ' + to + ' (' + n.total + ' items)', '');
    return 'Digest emailed to ' + to;
  } catch (e) { logSystem_('ERROR', 'dailyDigest', e && e.message, e && e.stack); throw new Error(friendly_(e)); }
}

function misHtml_(dash, title) {
  var k = dash.kpis;
  function r(n) { return fmtINRServer_(n); }
  var rows = dash.tables.delayed.map(function (x) {
    return '<tr><td style="padding:4px 8px">' + esc_(x.code) + '</td><td style="padding:4px 8px">' + esc_(x.name) + '</td><td style="padding:4px 8px">' + x.delay + 'd</td><td style="padding:4px 8px">' + esc_(x.rag) + '</td><td style="padding:4px 8px">' + esc_(x.pm) + '</td></tr>';
  }).join('');
  return '<div style="font-family:Arial,sans-serif;color:#0F172A;padding:20px;max-width:760px">'
    + '<div style="background:linear-gradient(90deg,#4F46E5,#7C3AED);color:#fff;padding:16px 20px;border-radius:10px;font-size:18px"><b>' + esc_(APP.name) + ' — ' + title + '</b><div style="font-size:11px;opacity:.9">' + todayStr_() + '</div></div>'
    + '<table style="width:100%;border-collapse:collapse;font-size:12px;margin-top:12px">'
    + '<tr><td style="padding:6px">Total Projects</td><td style="padding:6px"><b>' + k.totalProjects + '</b></td>'
    + '<td style="padding:6px">Live</td><td style="padding:6px"><b>' + k.liveProjects + '</b></td>'
    + '<td style="padding:6px">Delayed</td><td style="padding:6px"><b>' + k.delayed + '</b></td></tr>'
    + '<tr><td style="padding:6px">Contract Value</td><td style="padding:6px"><b>' + r(k.totalValue) + '</b></td>'
    + '<td style="padding:6px">Billed</td><td style="padding:6px"><b>' + r(k.billed) + '</b></td>'
    + '<td style="padding:6px">Received</td><td style="padding:6px"><b>' + r(k.received) + '</b></td></tr>'
    + '<tr><td style="padding:6px">Outstanding</td><td style="padding:6px"><b>' + r(k.outstanding) + '</b></td>'
    + '<td style="padding:6px">Avg Margin</td><td style="padding:6px"><b>' + (k.avgMargin === null ? '—' : k.avgMargin + '%') + '</b></td>'
    + '<td style="padding:6px">Red Flags</td><td style="padding:6px"><b>' + k.redFlags + '</b></td></tr></table>'
    + '<h3 style="font-size:13px;color:#4F46E5">Delayed / Red Projects</h3>'
    + '<table style="width:100%;border-collapse:collapse;font-size:11px;border:1px solid #E5E9F2">'
    + '<tr style="background:#F4F6FB"><th style="text-align:left;padding:5px 8px">Code</th><th style="text-align:left;padding:5px 8px">Project</th><th style="padding:5px 8px">Delay</th><th style="padding:5px 8px">RAG</th><th style="text-align:left;padding:5px 8px">PM</th></tr>'
    + (rows || '<tr><td colspan="5" style="padding:8px">None 🎉</td></tr>') + '</table></div>';
}

/* ============================================ 14. TRIGGERS & BACKUP ========================== */

function createTriggers() {
  deleteTriggers();
  ScriptApp.newTrigger('dailyDigest').timeBased().everyDays(1).atHour(8).create();
  ScriptApp.newTrigger('dailyBackup').timeBased().everyDays(1).atHour(2).create();
  ScriptApp.newTrigger('emailWeeklyMIS').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  return 'Triggers created: Daily Digest 08:00 · Weekly MIS Mon 07:00 · Daily Backup 02:00.';
}

function deleteTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['dailyDigest', 'emailWeeklyMIS', 'dailyBackup'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  return 'Existing PMR triggers removed.';
}

/** Copies this spreadsheet into the backup folder; keeps the newest APP.backupKeep copies. */
function dailyBackup(manual) {
  try {
    var f = folder_(APP.backupFolder);
    var name = ss_().getName() + ' BACKUP ' + Utilities.formatDate(new Date(), TZ_(), 'yyyy-MM-dd');
    var existing = f.getFilesByName(name);
    if (existing.hasNext()) {
      if (!manual) return 'Backup for today already exists.';
      existing.next().setTrashed(true);
    }
    DriveApp.getFileById(ss_().getId()).makeCopy(name, f);
    var copies = [], it2 = f.getFiles();
    while (it2.hasNext()) { var file = it2.next(); if (file.getName().indexOf(' BACKUP ') >= 0) copies.push(file); }
    copies.sort(function (a, b) { return a.getName() < b.getName() ? 1 : -1; });
    for (var i = APP.backupKeep; i < copies.length; i++) copies[i].setTrashed(true);
    var msg = 'Backup created: "' + name + '" (keeping last ' + APP.backupKeep + ').';
    logSystem_('INFO', 'dailyBackup', msg, '');
    return msg;
  } catch (e) { logSystem_('ERROR', 'dailyBackup', e && e.message, e && e.stack); throw new Error(friendly_(e)); }
}

/* ============================================ 15. SAMPLE DATA ================================ */

var SAMPLE_KEY = 'PMR_SAMPLE_IDS';

/** 3 fully linked demo projects: 200-seater office (Bengaluru), showroom (Delhi), cafeteria (Pune). */
function loadSampleData() {
  try {
    setupSystem();
    var props = PropertiesService.getUserProperties();
    if (props.getProperty(SAMPLE_KEY)) throw new Error('Sample data is already loaded. Run "Clear Sample Data" first.');
    var ids = {};
    function track(sheet, id) { (ids[sheet] = ids[sheet] || []).push(id); }
    function add(sheet, rec) { var r = saveRecord({sheet: sheet, record: rec}); track(sheet, r.record[SCHEMA[sheet].idKey]); return r.record; }

    // users (only added if missing; tracked sample emails are removable via Clear Sample Data)
    ['Rahul Verma|rahul.verma@sample.foi.in|PM', 'Priya Nair|priya.nair@sample.foi.in|PM', 'Anita Desai|anita.desai@sample.foi.in|Coordinator',
     'Suresh Kumar|suresh.kumar@sample.foi.in|Coordinator', 'Meera Joshi|meera.joshi@sample.foi.in|Coordinator', 'Arjun Mehta|arjun.mehta@sample.foi.in|Project Director']
      .forEach(function (u) {
        var p = u.split('|');
        var ex = usersList_().some(function (x) { return String(x.email).toLowerCase() === p[1]; });
        if (!ex) add('Settings', {Category: 'Users', Value: p[0], 'Info 1': p[1], 'Info 2': p[2], Active: 'Yes'});
      });
    // clients
    var c1 = add('Clients', {Name: 'Acme Technologies Pvt Ltd', Industry: 'IT/ITES', 'GST No': '29AAACA1234F1Z5', PAN: 'AAACA1234F', 'Billing Address': 'Prestige Tech Park, Kadubeesanahalli, Bengaluru 560103', 'Contact Persons': 'Kiran Rao (Finance) +91 98450 11223\nDevika S (Admin)', 'Payment Terms (days)': 45, 'Credit Rating': 'A+'});
    var c2 = add('Clients', {Name: 'Nova Retail India Ltd', Industry: 'Retail', 'GST No': '07AABCN5678M1Z8', PAN: 'AABCN5678M', 'Billing Address': 'DLF Mall, Saket, New Delhi 110017', 'Contact Persons': 'Rakesh Malhotra +91 98110 44556', 'Payment Terms (days)': 30, 'Credit Rating': 'A'});
    var c3 = add('Clients', {Name: 'GreenLeaf Foods Pvt Ltd', Industry: 'F&B', 'GST No': '27AAGCG9012P1Z2', PAN: 'AAGCG9012P', 'Billing Address': 'Hinjewadi Phase 2, Pune 411057', 'Contact Persons': 'Sunita Kulkarni +91 98220 77889', 'Payment Terms (days)': 30, 'Credit Rating': 'B+'});
    // vendors
    var v1 = add('Vendors', {'Vendor Name': 'Workspace Interiors Pvt Ltd', Trade: 'Carpentry/Furniture', 'Contact Person': 'Mahesh Iyer', Phone: '+91 98450 33445', Email: 'sales@workspaceinteriors.in', City: 'Bengaluru', 'GST No': '29AAECW1111A1Z6', 'Bank Details (masked)': 'HDFC 50100XXXX8821', 'Payment Terms (days)': 30, 'Rating (1-5)': 4, 'Empanelment Status': 'Preferred'});
    var v2 = add('Vendors', {'Vendor Name': 'CoolAir HVAC Solutions', Trade: 'HVAC', 'Contact Person': 'Farhan Ali', Phone: '+91 98110 22334', Email: 'projects@coolairhvac.in', City: 'Delhi', 'GST No': '07AAECC2222B1Z4', 'Bank Details (masked)': 'ICICI 0034XXXX1207', 'Payment Terms (days)': 45, 'Rating (1-5)': 3, 'Empanelment Status': 'Empanelled'});
    var v3 = add('Vendors', {'Vendor Name': 'Shakti Electricals', Trade: 'Electrical', 'Contact Person': 'Ravi Sharma', Phone: '+91 98220 55667', Email: 'ravi@shaktielectricals.co.in', City: 'Pune', 'GST No': '27AAECS3333C1Z9', 'Bank Details (masked)': 'SBI 1122XXXX9034', 'Payment Terms (days)': 30, 'Rating (1-5)': 4, 'Empanelment Status': 'Empanelled'});
    var v4 = add('Vendors', {'Vendor Name': 'GlowLux Lighting', Trade: 'Electrical', 'Contact Person': 'Nikhil Chopra', Phone: '+91 98104 66778', Email: 'orders@glowlux.in', City: 'Delhi', 'GST No': '07AAFCG4444D1Z1', 'Bank Details (masked)': 'AXIS 2211XXXX4455', 'Payment Terms (days)': 21, 'Rating (1-5)': 5, 'Empanelment Status': 'Preferred'});
    // projects
    var p1 = add('Projects', {'Project Code': 'AT-BLR-001', 'Project Name': 'Acme Technologies — 200-Seater Office Fit-Out', 'Client ID': c1['Client ID'], 'Contact Person': 'Kiran Rao', Phone: '+91 98450 11223', Email: 'kiran.rao@acmetech.example', 'Project Type': 'Office', 'Contract Type': 'Turnkey', City: 'Bengaluru', 'Site Address': 'Prestige Tech Park, Tower B, Kadubeesanahalli, Bengaluru', 'Built-up Area (sq ft)': 42000, Scope: 'Complete interior fit-out: 200 workstations, 6 cabins, 4 boardrooms, cafeteria, full MEP.', 'WO/LOA No': 'ATL/WO/2026/118', 'WO Date': dOffset_(-170), 'Contract Value (excl GST)': 65000000, 'GST %': '18', 'Approved Budget': 57000000, 'Est. Margin %': 12.3, 'Project Director': 'Arjun Mehta', 'Project Manager': 'Rahul Verma', Coordinator: 'Anita Desai', 'Site Engineer': 'Suresh Kumar', Designer: 'Meera Joshi', 'Client PMC': 'CBRE (A. Sharma)', 'Start Date': dOffset_(-160), 'Planned End': dOffset_(70), Status: 'Execution', Phase: 'Services', 'Planned %': 68, 'Actual %': 66, 'Next Milestone': 'Level 4 services closure', 'Milestone Date': dOffset_(9), 'Key Risk': 'Long-lead lighting fixtures from Italy', Priority: 'Critical', Remarks: 'Weekly PMR review Fridays 4 PM.'});
    var p2 = add('Projects', {'Project Code': 'NRS-DEL-001', 'Project Name': 'Nova Retail — Flagship Showroom, DLF Mall', 'Client ID': c2['Client ID'], 'Contact Person': 'Rakesh Malhotra', Phone: '+91 98110 44556', Email: 'rakesh.m@novaretail.example', 'Project Type': 'Showroom', 'Contract Type': 'Design & Build', City: 'Delhi', 'Site Address': 'DLF Mall, 2nd Floor, Saket, New Delhi', 'Built-up Area (sq ft)': 12500, Scope: 'Flagship retail showroom: facade glazing, premium flooring, accent lighting, VM fixtures.', 'WO/LOA No': 'NRI/WO/2026/042', 'WO Date': dOffset_(-100), 'Contract Value (excl GST)': 28000000, 'GST %': '18', 'Approved Budget': 25200000, 'Est. Margin %': 10, 'Project Director': 'Arjun Mehta', 'Project Manager': 'Priya Nair', Coordinator: 'Suresh Kumar', Designer: 'Meera Joshi', 'Client PMC': 'Internal (S. Bhatia)', 'Start Date': dOffset_(-92), 'Planned End': dOffset_(-6), 'Revised End': dOffset_(25), Status: 'Execution', Phase: 'Finishes', 'Planned %': 55, 'Actual %': 48, 'Next Milestone': 'Facade glazing completion', 'Milestone Date': dOffset_(6), 'Key Risk': 'Mall night-work permit restrictions', Priority: 'High'});
    var p3 = add('Projects', {'Project Code': 'GLF-PN-001', 'Project Name': 'GreenLeaf Foods — Corporate Cafeteria, Hinjewadi', 'Client ID': c3['Client ID'], 'Contact Person': 'Sunita Kulkarni', Phone: '+91 98220 77889', Email: 'sunita.k@greenleaf.example', 'Project Type': 'Cafeteria/Food Court', 'Contract Type': 'Lump Sum', City: 'Pune', 'Site Address': 'Hinjewadi Phase 2, Pune', 'Built-up Area (sq ft)': 8000, Scope: '300-cover corporate cafeteria: kitchen equipment, HVAC, seating, finishes.', 'WO/LOA No': 'GLF/WO/2025/317', 'WO Date': dOffset_(-220), 'Contract Value (excl GST)': 16500000, 'GST %': '18', 'Approved Budget': 14850000, 'Est. Margin %': 9.8, 'Project Director': 'Arjun Mehta', 'Project Manager': 'Rahul Verma', Coordinator: 'Anita Desai', 'Site Engineer': 'Suresh Kumar', Designer: 'Meera Joshi', 'Start Date': dOffset_(-210), 'Planned End': dOffset_(-20), Status: 'Snagging', Phase: 'Testing & Commissioning', 'Planned %': 92, 'Actual %': 96, 'Next Milestone': 'Fire NOC inspection', 'Milestone Date': dOffset_(4), 'Key Risk': 'Kitchen commissioning vendor availability', Priority: 'High'});

    // BOQ lines (client rate + budget rate drive contract & budget totals)
    [[p1, 'Partition', 'GYPSUM-100', 'Gypsum board partition 100mm', 'Sqft', 18000, 245, 215],
     [p1, 'False Ceiling', 'CLG-GRID', 'Grid ceiling + mineral fiber tiles', 'Sqft', 22000, 190, 165],
     [p1, 'Flooring', 'VT-PVC', 'Premium vinyl tile flooring', 'Sqft', 26000, 210, 180],
     [p1, 'Electrical', 'EL-POINT', 'Low current + power points package', 'Nos', 3200, 850, 720],
     [p1, 'HVAC', 'HVAC-VRF', 'VRF air conditioning system', 'Sqft', 42000, 95, 82],
     [p1, 'Carpentry/Furniture', 'FURN-WS', 'Modular workstations + chairs (200 pax)', 'Set', 200, 32000, 27500],
     [p2, 'Glazing', 'GLZ-FACADE', 'Structural glazing facade', 'Sqft', 2400, 1450, 1280],
     [p2, 'Flooring', 'MRB-FLR', 'Marble-look porcelain flooring', 'Sqft', 9000, 320, 268],
     [p2, 'Electrical', 'ACC-LGT', 'Accent + track lighting package', 'Nos', 450, 2600, 2250],
     [p3, 'Civil', 'CIV-KIT', 'Kitchen civil block work', 'Lsum', 1, 1850000, 1620000],
     [p3, 'HVAC', 'HVAC-DUCT', 'Kitchen exhaust + fresh air ducting', 'Rmt', 480, 1650, 1400],
     [p3, 'Soft Furnishing', 'SEAT-300', 'Seating for 300 covers', 'Nos', 300, 4200, 3600]
    ].forEach(function (b) {
      add('BOQ', {'Project ID': b[0]['Project ID'], 'Section/Trade': b[1], 'Item Code': b[2], Description: b[3], UOM: b[4], Qty: b[5], 'Client Rate': b[6], 'Budget Rate': b[7], 'Qty Executed': Math.round(b[5] * (b[0] === p1 ? 0.66 : (b[0] === p2 ? 0.5 : 0.95)))});
    });
    // procurement (PO/WO)
    [[p1, v1, 'Carpentry/Furniture', 'Modular workstations — phase 1 (120 sets)', 120, 'Set', 27500, '18', 'Approved', -20, -2, 2200000],
     [p1, v4, 'Electrical', 'LED lighting fixtures — Italy import', 1, 'Lsum', 3250000, '18', 'Pending Approval', 10, '', 0],
     [p2, v2, 'HVAC', 'VRV outdoor + indoor units', 1, 'Lsum', 2100000, '18', 'Delivered', -35, -12, 2100000],
     [p2, v4, 'Electrical', 'Track + accent lighting', 1, 'Lsum', 980000, '18', 'Partially Delivered', -15, 3, 400000],
     [p3, v3, 'Electrical', 'Kitchen HT/LT panel + wiring', 1, 'Lsum', 1250000, '18', 'Closed', -80, -55, 1250000]
    ].forEach(function (o) {
      add('Procurement', {Date: dOffset_(o[9] - 30), 'Project ID': o[0]['Project ID'], Vendor: o[1]['Vendor Name'], Trade: o[2], Description: o[3], Qty: o[4], UOM: o[5], Rate: o[6], 'GST %': o[7], Status: o[8], 'Required By': dOffset_(o[9]), 'Delivery Date': o[10] ? dOffset_(o[10]) : '', Paid: o[11], 'Approved By': 'Arjun Mehta'});
    });
    // costs (Admin/Director only)
    [[p1, 'Material', 'Workspace Interiors Pvt Ltd', 'WSI/2411', 18970000, '18', 'Partially Paid', -40],
     [p1, 'Labour', '—', 'LAB-OCT-1', 4250000, '0', 'Paid', -12],
     [p1, 'Subcontract', 'CoolAir HVAC Solutions', 'CAH/8891', 3820000, '18', 'Paid', -55],
     [p2, 'Material', 'GlowLux Lighting', 'GLX/1102', 980000, '18', 'Due', -20],
     [p2, 'Subcontract', 'Metro Glazing Works', 'MGW/221', 1720000, '18', 'Partially Paid', -30],
     [p3, 'Equipment', 'Kitchen King Equipments', 'KKE/771', 2450000, '18', 'Paid', -90],
     [p3, 'Labour', '—', 'LAB-SEP-3', 1850000, '0', 'Paid', -25]
    ].forEach(function (c) {
      add('Cost Tracker', {Date: dOffset_(c[7]), 'Project ID': c[0]['Project ID'], 'Cost Head': c[1], Vendor: c[2], 'Invoice No': c[3], Amount: c[4], 'GST %': c[5], 'Payment Status': c[6]});
    });
    // billing (RA bills with retention, GST, receipts)
    [[p1, 'RA', -95, 9500000, 5, 8550000, -80], [p1, 'RA', -40, 12500000, 5, 0, 0],
     [p2, 'Advance', -85, 2800000, 0, 2800000, -70], [p2, 'RA', -78, 4200000, 5, 0, 0],
     [p3, 'RA', -120, 6600000, 5, 6270000, -100], [p3, 'RA', -35, 4300000, 5, 1500000, -8]
    ].forEach(function (b) {
      add('Billing', {'Project ID': b[0]['Project ID'], 'Bill Type': b[1], 'Bill Date': dOffset_(b[2]), 'Gross Amount': b[3], 'Retention %': b[4], 'GST %': '18', 'Submitted Date': dOffset_(b[2] + 2), 'Received Amount': b[5], 'Received Date': b[6] ? dOffset_(b[6]) : '', Status: b[5] > 0 ? (b[5] >= Math.round(b[3] * 0.95) ? 'Paid' : 'Partially Paid') : 'Overdue'});
    });
    // drawings & approvals
    [[p1, 'ID plans', 'Interior', 'R2', -120, 'Approved', -95], [p1, 'L4 services GFC', 'MEP', 'R1', -60, 'Pending', 0],
     [p2, 'Facade glazing details', 'Facade/Glazing', 'R3', -70, 'Approved', -50], [p2, 'VM fixture layout', 'Interior', 'R1', -25, 'Pending', 0],
     [p3, 'Kitchen equipment layout', 'MEP', 'R2', -150, 'Approved', -130]]
      .forEach(function (dw, idx) {
        add('Drawings & Approvals', {'Ref No': dw[0]['Project Code'].slice(0, 3).toUpperCase() + '-' + pad_(idx + 1, 3), 'Project ID': dw[0]['Project ID'], Title: dw[1], Discipline: dw[2], Revision: dw[3], 'Issued Date': dOffset_(dw[4] - 10), 'Submitted Date': dOffset_(dw[4]), 'Approval Status': dw[5], 'Approved Date': dw[6] ? dOffset_(dw[6]) : ''});
      });
    // DPRs
    [[p1, -2, 'Sunny', 38, 22, 'L4 partition close-out, ceiling grid 60% done', 'Flooring start zone B', 'Gypsum boards 400 nos', 'Cement lift waiting'],
     [p2, -1, 'Cloudy', 16, 9, 'Facade glazing panel 12-18 installed', 'Panels 19-24, flooring zone 2', 'Aluminium composite panels', 'Mall work permit delay 2 hrs'],
     [p3, -1, 'Humid', 22, 30, 'Kitchen equipment commissioning day 2', 'Hood balancing, snag rectification', '—', '2 snags re-opened by client QA']]
      .forEach(function (dp) {
        add('DPR', {Date: dOffset_(dp[1]), 'Project ID': dp[0]['Project ID'], 'Site Engineer': dp[0] === p3 ? 'Suresh Kumar' : 'Anita Desai', Weather: dp[2], 'Skilled Manpower': dp[3], 'Unskilled Manpower': dp[4], 'Work Done': dp[5], 'Planned for Tomorrow': dp[6], 'Material Received': dp[7], Issues: dp[8], 'Safety Incident (Y/N)': 'No'});
      });
    // variation orders
    add('Variation Orders', {'Project ID': p1['Project ID'], Description: 'Additional acoustic baffles in boardrooms (client request)', 'Requested By': 'Kiran Rao (Acme)', Date: dOffset_(-30), Amount: 1250000, 'Approval Status': 'Approved', 'Approved Date': dOffset_(-18), 'Timeline Impact (days)': 0});
    add('Variation Orders', {'Project ID': p2['Project ID'], Description: 'Upgrade facade LED media strip to DMX addressable', 'Requested By': 'Rakesh Malhotra (Nova)', Date: dOffset_(-10), Amount: 680000, 'Approval Status': 'Pending', 'Approved Date': '', 'Timeline Impact (days)': 5});
    // risks & issues
    [[p1, 'Risk', 'Long-lead lighting fixtures from Italy may delay level 4 closure', 'Vendor', 3, 4, 'Arjun Mehta', 'Air-freight partial shipment; local substitute for zone B', 12, 'In Progress'],
     [p1, 'Issue', 'Client PMC sign-off cycle averaging 9 days vs 4 planned', 'Client', 4, 3, 'Rahul Verma', 'Escalated in weekly PMR; standing 48h SLA agreed', 5, 'Open'],
     [p2, 'Risk', 'Mall permits restrict work after 22:00', 'Statutory', 4, 4, 'Priya Nair', 'Night-shift crew 2x; alternate sequence for noisy trades', 15, 'Mitigated'],
     [p3, 'Issue', 'Kitchen commissioning vendor availability during festive week', 'Resource', 3, 3, 'Rahul Verma', 'Pre-booking confirmed with penalty clause', 7, 'In Progress']]
      .forEach(function (rk) {
        add('Risks & Issues', {'Project ID': rk[0]['Project ID'], Type: rk[1], Description: rk[2], Category: rk[3], 'Probability (1-5)': rk[4], 'Impact (1-5)': rk[5], Owner: rk[6], Mitigation: rk[7], 'Target Date': dOffset_(rk[8]), Status: rk[9]});
      });
    // snags
    [[p3, 'Kitchen line 2', 'Hood velocity below spec at corner', 'HVAC', -6, 'Suresh Kumar', 3, 'In Progress', 0],
     [p3, 'Dining hall', '2 floor tiles chipped near buffet counter', 'Flooring', -4, 'Anita Desai', -1, 'Closed', -2],
     [p3, 'Washroom M', 'Sensor faucet intermittent', 'Plumbing', -2, 'Suresh Kumar', 5, 'Open', 0],
     [p2, 'Entrance', 'Glass door closer needs adjustment', 'Glazing', -1, 'Suresh Kumar', 4, 'Open', 0]]
      .forEach(function (sn) {
        add('Snags & Handover', {'Project ID': sn[0]['Project ID'], Location: sn[1], Description: sn[2], Trade: sn[3], 'Raised Date': dOffset_(sn[4]), 'Assigned To': sn[5], 'Target Date': dOffset_(sn[6]), 'Closed Date': sn[8] ? dOffset_(sn[8]) : '', Status: sn[7]});
      });

    props.setProperty(SAMPLE_KEY, JSON.stringify(ids));
    invalidate_();
    _m.ctx = null; _m.scope = null;
    return 'Sample data loaded: 3 linked projects (Bengaluru office · Delhi showroom · Pune cafeteria) with BOQ, POs, costs, bills, drawings, DPRs, VOs, risks and snags.';
  } catch (e) { logSystem_('ERROR', 'loadSampleData', e && e.message, e && e.stack); throw new Error(friendly_(e)); }
}

/** Removes everything created by loadSampleData() (tracked via UserProperties). */
function clearSampleData() {
  try {
    var props = PropertiesService.getUserProperties();
    var raw = props.getProperty(SAMPLE_KEY);
    if (!raw) return 'No sample data found — nothing to clear.';
    var ids = JSON.parse(raw);
    var removed = withLock_(function () {
      var removed = 0;
      ['BOQ', 'Procurement', 'Cost Tracker', 'Billing', 'Drawings & Approvals', 'DPR', 'Variation Orders',
       'Risks & Issues', 'Snags & Handover', 'Documents Checklist', 'Projects', 'Vendors', 'Clients'].forEach(function (sheet) {
        var list = ids[sheet] || [];
        if (!list.length) return;
        var d = readData_(sheet), sc = SCHEMA[sheet];
        var iK = d.hm[sc.idKey];
        var doomed = [];
        d.rows.forEach(function (r, i) { if (list.indexOf(String(r[iK] || '')) >= 0) doomed.push(i + 2); });
        for (var i = doomed.length - 1; i >= 0; i--) { d.sheet.deleteRow(doomed[i]); removed++; }
        invalidate_(sheet);
      });
      var dS = readData_('Settings');
      var iC = dS.hm['Category'], iE = dS.hm['Info 1'];
      var doomed = [];
      dS.rows.forEach(function (r, i) {
        if (String(r[iC]) === 'Users' && String(r[iE] || '').indexOf('@sample.foi.in') >= 0) doomed.push(i + 2);
      });
      for (var j = doomed.length - 1; j >= 0; j--) { dS.sheet.deleteRow(doomed[j]); removed++; }
      invalidate_();
      invalidateSettingsCache_();
      return removed;
    });
    props.deleteProperty(SAMPLE_KEY);
    _m.scope = null; _m.ctx = null;
    return 'Sample data cleared (' + removed + ' rows removed).';
  } catch (e) { logSystem_('ERROR', 'clearSampleData', e && e.message, e && e.stack); throw new Error(friendly_(e)); }
}
