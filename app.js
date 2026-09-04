/* ============================================================
   Work Payment Log
   A direct port of Rews_Work_Payment_Log.xlsx.

   Formulas mirrored from the workbook:
     hours = IF(OR(start="",finish=""),0,ROUND(MAX(0,(finish-start)*24-break),4))
     time  = IF(hours=0,"-",INT(hours)&"h "&TEXT(ROUND(frac*60,0),"00")&"m")
     pay   = ROUND(hours*rate,2)
     weekly / period totals = SUMIFS over the timesheet by date range
     days worked            = unique completed work dates with hours > 0
   ============================================================ */
(function () {
  "use strict";

  var STORE_KEY = "wpl.v1";
  var THEME_KEY = "wpl.theme";
  var SYNC_KEY = "wpl.sync";
  var DEVICE_KEY = "wpl.device";
  var CLOUD_SYNC_SIGNATURE_KEY = "wpl.cloud.syncedSignature";
  var APP_COLORS = {
    orange: { label: "Orange", light: ["#e0530a","#f79009","#ffeede","#fdf7f2","#f8ede3","#ffffff","#fdf6f0","#f0e2d6","#f7ece3","#231710","#6b5748","#9a8271"], dark: ["#ff9a4d","#ffb56b","#3a2616","#161009","#120d08","#231a12","#2c2118","#3d2e21","#2f231a","#fdf3ea","#d2b9a3","#a08a74"] },
    blue:   { label: "Blue",   light: ["#2563eb","#60a5fa","#eaf2ff","#f5f8ff","#eaf1ff","#ffffff","#f1f6ff","#d9e5f7","#e7effb","#17233a","#52647f","#8190a6"], dark: ["#60a5fa","#93c5fd","#172b4d","#09121f","#070d16","#121f30","#17283c","#293e58","#1f3045","#f3f8ff","#bed0e8","#8298b4"] },
    green:  { label: "Green",  light: ["#078a63","#34b98a","#e5f7f0","#f4fbf8","#e6f5ef","#ffffff","#eef9f5","#d5ebe2","#e5f3ed","#142c25","#4e6b61","#7c988e"], dark: ["#34d399","#6ee7b7","#15372d","#091612","#07110e","#12231d","#172d25","#29473b","#1f382f","#eefbf5","#b8d9cc","#7fa295"] },
    purple: { label: "Purple", light: ["#7c3aed","#a78bfa","#f2edff","#faf7ff","#f1ebff","#ffffff","#f7f3ff","#e6dcfa","#f0e9fb","#271b3a","#665579","#9485a3"], dark: ["#a78bfa","#c4b5fd","#2d2146","#140e1d","#0f0a16","#21172d","#2a1d38","#423157","#342642","#faf5ff","#d5c4e6","#9f8db3"] },
    pink:   { label: "Pink",   light: ["#c0266d","#f472b6","#fff0f7","#fff7fb","#fdebf4","#ffffff","#fff2f8","#f5d9e7","#fae8f1","#361725","#7d5265","#a57f90"], dark: ["#f472b6","#f9a8d4","#421c31","#1b0d14","#14090f","#2a1620","#351b28","#542a3e","#412132","#fff3f8","#e4becf","#ad8597"] },
    teal:   { label: "Teal",   light: ["#0f766e","#2bbbad","#e7f7f5","#f4fbfa","#e6f5f3","#ffffff","#eff9f8","#d3ebe8","#e4f3f1","#12302d","#4b6e6a","#799895"], dark: ["#2dd4bf","#5eead4","#123a36","#071615","#06100f","#102523","#142f2c","#244945","#1b3935","#effdfa","#b7dcd7","#7ca39e"] }
  };
  function fullWeek() { return Number(state.settings.weekHours) || 40; }

  /* ---------- state ---------- */
  /* Declared before load() runs: load() sets it, and a `var … = false` sitting
     below would re-run its initialiser afterwards and wipe the flag. */
  var migrationApplied = false;    // set by load(), persisted once at boot

  var state = load();
  var lastSettingsSnapshot = topLevelSettingsSnapshot(state.settings);
  var device = loadDevice();
  var syncCfg = loadSyncCfg();
  var applyingRemote = false;      // guards against sync → save → sync loops

  /* ---------- persistence ---------- */
  function load() {
    try {
      var raw = localStorage.getItem(STORE_KEY);
      if (raw) {
        var s = JSON.parse(raw);
        if (s && s.settings && Array.isArray(s.entries)) {
          s.settings = Object.assign({}, WPL_SEED.settings, s.settings);
          s.settings.leaveRates = Object.assign({}, WPL_SEED.leaveRates,
            s.settings.leaveRates || {});

          /* Anyone already using the app has answered these questions in
             practice — do not put them through first-run setup. */
          if (!s.onboardMigrated) {
            s.settings.onboarded = true;
            s.onboardMigrated = true;
            migrationApplied = true;
          }
          if (typeof s.settingsUpdatedAt !== "number") s.settingsUpdatedAt = 0;
          if (!s.settingsFieldUpdatedAt || typeof s.settingsFieldUpdatedAt !== "object") {
            s.settingsFieldUpdatedAt = {};
            Object.keys(s.settings).forEach(function (key) { s.settingsFieldUpdatedAt[key] = s.settingsUpdatedAt || 0; });
            migrationApplied = true;
          }
          if (!Array.isArray(s.history)) s.history = [];

          /* v1.1.4: the unpaid break is 15 minutes, not 30. Anyone still carrying
             the old half-hour default is moved across once, and the timestamp is
             bumped so the correction wins on the other device too rather than
             being overwritten by its stale copy. Set it back by hand and it
             stays — this never runs twice. */
          if (!s.breakMigrated) {
            if (Number(s.settings.breakHours) === 0.5) {
              s.settings.breakHours = 0.25;
              s.settingsUpdatedAt = Date.now();
            }
            s.breakMigrated = true;
            migrationApplied = true;
          }

          /* v1.1.5: the shift ends at 15:45. Same one-time move for anyone still
             on the old 16:00 default. Only the prefill for new days changes —
             days already logged keep whatever times they were saved with. */
          /* v1.3.0: the app used to arrive pre-filled months ahead. Anything
             still untouched (updatedAt 0) and dated after today is removed, so
             you log your own days from here. Anything you edited is kept. */
          if (!s.futureSeedCleared) {
            var todayIso = toISO(today());
            s.entries.forEach(function (e) {
              if (!e.deleted && !e.updatedAt && e.date > todayIso) {
                e.deleted = true;
                e.updatedAt = Date.now();
              }
            });
            s.futureSeedCleared = true;
            migrationApplied = true;
          }

          if (!s.finishMigrated) {
            if (s.settings.defaultFinish === "16:00") {
              s.settings.defaultFinish = "15:45";
              s.settingsUpdatedAt = Date.now();
            }
            s.finishMigrated = true;
            migrationApplied = true;
          }
          s.entries.forEach(function (e, index) {
            if (!e.id) e.id = legacyEntryId(e, index);
            if (typeof e.updatedAt !== "number") e.updatedAt = 0;
            if (typeof e.deleted !== "boolean") e.deleted = false;
            if (typeof e.leave !== "string") e.leave = "";
          });
          ensureProductData(s);
          return s;
        }
      }
    } catch (err) { /* fall through to seed */ }
    return seedState();
  }

  function seedState() {
    var seeded = {
      settings: Object.assign({}, WPL_SEED.settings,
        { leaveRates: Object.assign({}, WPL_SEED.leaveRates) }),
      settingsUpdatedAt: 0,
      settingsFieldUpdatedAt: {},
      /* Only days that have actually happened — the rest are yours to log. */
      entries: WPL_SEED.entries.filter(function (e) {
        return e.date <= toISO(today());
      }).map(function (e) {
        return {
          id: uid(), date: e.date, start: e.start || "", finish: e.finish || "",
          notes: e.notes || "", leave: e.leave || "", employerId: "default", updatedAt: 0, deleted: false
        };
      }),
      history: [],
      breakMigrated: true,
      finishMigrated: true,
      futureSeedCleared: true
    };
    ensureProductData(seeded);
    return seeded;
  }

  /* Additive product migration. Old backups and synced stores continue to work:
     every existing day belongs to one generated default employer and the old
     rate/shift fields remain the source for that employer. */
  function ensureProductData(s) {
    var st = s.settings || (s.settings = {});
    if (!Array.isArray(st.employers) || !st.employers.length) {
      st.employers = [{ id: "default", name: "My job", color: "#e85d0f",
        rate: Number(st.rate || 0), start: st.defaultStart || "08:00",
        finish: st.defaultFinish || "15:45", breakHours: Number(st.breakHours || 0) }];
      migrationApplied = true;
    }
    if (!st.defaultEmployerId || !st.employers.some(function (x) { return x.id === st.defaultEmployerId; })) {
      st.defaultEmployerId = st.employers[0].id;
    }
    if (!Array.isArray(st.shiftTemplates)) {
      st.shiftTemplates = [{ id: "usual", name: "Usual shift",
        start: st.defaultStart || "08:00", finish: st.defaultFinish || "15:45" }];
      migrationApplied = true;
    }
    st.notifications = Object.assign({ logHours: false, payday: false, missingFinish: false,
      incomplete: false, backup: false, sync: false, payslip:false, spending:false, updates: true }, st.notifications || {});
    if (!validStoredTime(st.quietStart || "")) st.quietStart="22:00";
    if (!validStoredTime(st.quietEnd || "")) st.quietEnd="07:00";
    if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(st.logReminderTime || "")) st.logReminderTime = "18:00";
    if (!Array.isArray(st.logReminderDays) || !st.logReminderDays.length) {
      st.logReminderDays = [1, 2, 3, 4, 5];
    } else {
      st.logReminderDays = st.logReminderDays.map(Number).filter(function (d, i, all) {
        return d >= 0 && d <= 6 && all.indexOf(d) === i;
      });
      if (!st.logReminderDays.length) st.logReminderDays = [1, 2, 3, 4, 5];
    }
    if (!st.paidPeriods || typeof st.paidPeriods !== "object") st.paidPeriods = {};
    st.takeHome = Object.assign({ country: "MT", profile: "mt-single",
      social: "mt-prorata", otherIncome: 0, workMonths: 12, customTax: 0,
      customSocial: 0, customOther: 0, customFixed: 0, estimatedHours: 0,
      estimateEmployerId: st.defaultEmployerId }, st.takeHome || {});
    st.takeHome.workMonths = Math.max(1, Math.min(12, parseInt(st.takeHome.workMonths,10) || 12));
    st.takeHome.estimatedHours = Math.max(0, Number(st.takeHome.estimatedHours) || 0);
    if (!st.employers.some(function(e){return e.id===st.takeHome.estimateEmployerId;})) st.takeHome.estimateEmployerId=st.defaultEmployerId;
    st.takeHome.payslip = Object.assign({gross:0,tax:0,social:0,other:0,net:0},st.takeHome.payslip||{});
    var defaultSpendingCategories = [
      { id: "bills", name: "Bills", color: "#3b82f6", mode: "percent", value: 35 },
      { id: "food", name: "Food", color: "#f59e0b", mode: "percent", value: 20 },
      { id: "savings", name: "Savings", color: "#10b981", mode: "percent", value: 25 },
      { id: "spending", name: "Spending money", color: "#8b5cf6", mode: "percent", value: 20 }
    ];
    if (!st.spendingPlan || typeof st.spendingPlan !== "object") {
      st.spendingPlan = { incomeMode: "takehome", customIncome: 0,
        categories: defaultSpendingCategories };
      migrationApplied = true;
    }
    st.spendingPlan.incomeMode = st.spendingPlan.incomeMode === "custom" ? "custom" : "takehome";
    st.spendingPlan.customIncome = Math.max(0, Number(st.spendingPlan.customIncome) || 0);
    if (!Array.isArray(st.spendingPlan.transactions)) st.spendingPlan.transactions = [];
    if (!Array.isArray(st.spendingPlan.bills)) st.spendingPlan.bills=[];
    if (!Array.isArray(st.spendingPlan.goals)) st.spendingPlan.goals=[];
    /* An empty list is valid: somebody who has no regular costs may remove
       every starter category. Only old stores without a list get defaults. */
    if (!Array.isArray(st.spendingPlan.categories)) {
      st.spendingPlan.categories = defaultSpendingCategories;
      migrationApplied = true;
    }
    st.spendingPlan.categories = st.spendingPlan.categories.map(function (category, index) {
      category = category && typeof category === "object" ? category : {};
      var color = /^#[0-9a-f]{6}$/i.test(category.color || "") ? category.color : "#8b5cf6";
      return { id: String(category.id || ("category-" + index + "-" + uid())),
        name: String(category.name || "New category").slice(0, 40), color: color,
        mode: category.mode === "amount" ? "amount" : "percent",
        value: Math.max(0, Number(category.value) || 0), enabled: category.enabled !== false };
    });
    st.spendingPlan.transactions = st.spendingPlan.transactions.map(function (transaction, index) {
      transaction = transaction && typeof transaction === "object" ? transaction : {};
      return { id: String(transaction.id || ("transaction-" + index + "-" + uid())),
        date: validStoredDate(transaction.date) ? transaction.date : toISO(today()),
        categoryId: String(transaction.categoryId || ""),
        amount: Math.max(0, Number(transaction.amount) || 0),
        note: String(transaction.note || "").slice(0, 60),
        updatedAt: Number(transaction.updatedAt) || 0 };
    }).filter(function (transaction) { return transaction.amount > 0; });
    st.spendingPlan.bills=st.spendingPlan.bills.map(function(bill,index){bill=bill&&typeof bill==="object"?bill:{};return{id:String(bill.id||("bill-"+index+"-"+uid())),name:String(bill.name||"Bill").slice(0,40),amount:Math.max(0,Number(bill.amount)||0),nextDue:validStoredDate(bill.nextDue)?bill.nextDue:toISO(today()),categoryId:String(bill.categoryId||""),enabled:bill.enabled!==false};}).filter(function(bill){return bill.amount>0;});
    st.spendingPlan.goals=st.spendingPlan.goals.map(function(goal,index){goal=goal&&typeof goal==="object"?goal:{};return{id:String(goal.id||("goal-"+index+"-"+uid())),name:String(goal.name||"Savings goal").slice(0,40),target:Math.max(0,Number(goal.target)||0),saved:Math.max(0,Number(goal.saved)||0),targetDate:validStoredDate(goal.targetDate)?goal.targetDate:""};}).filter(function(goal){return goal.target>0;});
    /* v1.10.17: the old stipend choice hid social security even when a user
       had selected an employee contribution category. That combination means
       they are employed, so move it to the matching student-employee profile. */
    if (st.takeHome.country === "MT" && st.takeHome.profile === "mt-stipend" && st.takeHome.social !== "none") {
      var wasUnder18 = st.takeHome.social === "mt-under18" || st.takeHome.social === "mt-apprentice-under18";
      st.takeHome.profile = wasUnder18 ? "mt-student-under18" : "mt-student-18";
      st.takeHome.social = wasUnder18 ? "mt-apprentice-under18" : "mt-apprentice18";
      s.settingsUpdatedAt = Date.now();
      migrationApplied = true;
    }
    (s.entries || []).forEach(function (e) {
      if (!e.employerId) { e.employerId = st.defaultEmployerId; migrationApplied = true; }
      var employer = st.employers.filter(function (x) { return x.id === e.employerId; })[0] || st.employers[0];
      /* Freeze the terms used for an existing shift. Before this migration an
         employer rate or break change silently recalculated every old payday. */
      if (typeof e.rateSnapshot !== "number" || !isFinite(e.rateSnapshot)) {
        e.rateSnapshot = Math.max(0, Number(employer && employer.rate) || 0);
        migrationApplied = true;
      }
      if (typeof e.breakHours !== "number" || !isFinite(e.breakHours)) {
        e.breakHours = Math.max(0, Number(employer && employer.breakHours) || 0);
        migrationApplied = true;
      }
      if (Array.isArray(e.breaks)) {
        e.breaks = e.breaks.map(function (item, index) {
          item = item && typeof item === "object" ? item : {};
          return { id: String(item.id || ("break-" + index + "-" + uid())),
            duration: Math.max(0, Number(item.duration) || 0), paid: !!item.paid };
        }).filter(function (item) { return item.duration > 0; });
        e.breakHours = xround(e.breaks.reduce(function (sum, item) { return sum + (item.paid ? 0 : item.duration); }, 0), 4);
      }
      if (["planned", "completed", "approved", "paid"].indexOf(e.status) === -1) {
        e.status = e.start && e.finish ? "completed" : "planned";
        migrationApplied = true;
      }
      if (typeof e.payMultiplier !== "number" || !isFinite(e.payMultiplier) || e.payMultiplier < 0) { e.payMultiplier=1;migrationApplied=true; }
      e.payMultiplier=Math.min(10,e.payMultiplier);
      if (["standard","overtime","night","sunday","public","custom"].indexOf(e.rateType)===-1) e.rateType="standard";
    });
  }

  /* ---------- types of day ----------
     A leave day is stored like any other: the times decide the hours, so paid
     leave is simply a normal day's hours carrying a label, and unpaid leave has
     no times and therefore no hours. Nothing in the pay maths needed changing.
     `paid` here only decides what the editor prefills when you pick a type. */
  var LEAVE_TYPES = [
    { key: "",            label: "Working day",              tone: "work" },
    { key: "annual",      label: "Annual leave",             tone: "leave" },
    { key: "public",      label: "Public holiday",           tone: "holiday" },
    { key: "sick",        label: "Sick leave",               tone: "sick" },
    { key: "injury",      label: "Injury leave",             tone: "sick" },
    { key: "quarantine",  label: "Quarantine leave",         tone: "sick" },
    { key: "bereavement", label: "Bereavement leave",        tone: "leave" },
    { key: "marriage",    label: "Marriage leave",           tone: "leave" },
    { key: "birth",       label: "Birth leave",              tone: "leave" },
    { key: "maternity",   label: "Maternity leave",          tone: "leave" },
    { key: "adoption",    label: "Adoption leave",           tone: "leave" },
    { key: "parental",    label: "Parental leave",           tone: "unpaid" },
    { key: "carers",      label: "Carer's leave",            tone: "unpaid" },
    { key: "family",      label: "Urgent family leave",      tone: "unpaid" },
    { key: "jury",        label: "Jury service",             tone: "leave" },
    { key: "study",       label: "Study or exam leave",      tone: "leave" },
    { key: "toil",        label: "Time off in lieu",         tone: "leave" },
    { key: "unpaid",      label: "Unpaid leave",             tone: "unpaid" },
    { key: "other",       label: "Other leave",              tone: "unpaid" },
    { key: "off",         label: "Day off",                  tone: "off" }
  ];

  /* What a day of this leave pays, as a percentage of a normal day. */
  function leavePct(key) {
    if (!key) return 100;
    var rates = state.settings.leaveRates || {};
    var v = rates[key];
    return typeof v === "number" ? v : (typeof WPL_SEED.leaveRates[key] === "number"
      ? WPL_SEED.leaveRates[key] : 100);
  }

  function leaveInfo(key) {
    for (var i = 0; i < LEAVE_TYPES.length; i++) {
      if (LEAVE_TYPES[i].key === (key || "")) {
        var t = LEAVE_TYPES[i];
        return { key: t.key, label: t.label, tone: t.tone, pct: leavePct(t.key),
                 paid: leavePct(t.key) > 0 };
      }
    }
    return { key: "", label: "Working day", tone: "work", pct: 100, paid: true };
  }
  function isLeave(e) { return !!(e && e.leave); }

  /* ---------- version history ----------
     Append-only, capped, and synced like everything else, so the log on the
     phone and the log on the Mac end up telling the same story. */
  function snapshotOf(e) {
    if (!e) return null;
    return {
      id: e.id || "", date: e.date || "",
      start: e.start || "", finish: e.finish || "", notes: e.notes || "",
      leave: e.leave || "", employerId: e.employerId || state.settings.defaultEmployerId,
      rateSnapshot: Math.max(0, Number(e.rateSnapshot) || 0),
      breakHours: Math.max(0, Number(e.breakHours) || 0),
      breaks: Array.isArray(e.breaks) ? e.breaks.map(function(item){return {id:item.id,duration:Number(item.duration)||0,paid:!!item.paid};}) : undefined,
      status: ["planned", "completed", "approved", "paid"].indexOf(e.status) !== -1 ? e.status : "completed",
      payMultiplier:e.payMultiplier==null?1:Math.min(10,Math.max(0,Number(e.payMultiplier)||0)),rateType:e.rateType||"standard",
      deleted: !!e.deleted
    };
  }

  function sameSnapshot(a, b) {
    if (!a || !b) return false;
    return a.start === b.start && a.finish === b.finish && a.notes === b.notes &&
      (a.leave || "") === (b.leave || "") && (a.employerId || "") === (b.employerId || "") &&
      Number(a.rateSnapshot || 0) === Number(b.rateSnapshot || 0) &&
      Number(a.breakHours || 0) === Number(b.breakHours || 0) &&
      JSON.stringify(a.breaks || []) === JSON.stringify(b.breaks || []) &&
      (a.status || "completed") === (b.status || "completed") &&
      Number(a.payMultiplier==null?1:a.payMultiplier) === Number(b.payMultiplier==null?1:b.payMultiplier) && (a.rateType||"standard") === (b.rateType||"standard") &&
      !!a.deleted === !!b.deleted;
  }

  function recordHistory(kind, date, before, after, action) {
    if (kind === "day" && sameSnapshot(before, after)) return;   // nothing actually moved
    state.history.push({
      id: uid() + uid(),
      at: Date.now(),
      device: device.name,
      kind: kind,
      date: date || "",
      entryId: kind === "day" ? ((after && after.id) || (before && before.id) || "") : "",
      action: action || (before ? "edit" : "add"),
      before: before,
      after: after
    });
    if (state.history.length > WPLSync.HISTORY_LIMIT) {
      state.history = state.history.slice(-WPLSync.HISTORY_LIMIT);
    }
  }

  function loadDevice() {
    try {
      var d = JSON.parse(localStorage.getItem(DEVICE_KEY) || "null");
      if (d && d.id) return d;
    } catch (err) { /* new device */ }
    var d = { id: uid() + uid(), name: defaultDeviceName() };
    try { localStorage.setItem(DEVICE_KEY, JSON.stringify(d)); } catch (err) { /* ignore */ }
    return d;
  }

  function defaultDeviceName() {
    if (window.WPLBridge && window.WPLBridge.deviceName) {
      try { return window.WPLBridge.deviceName(); } catch (err) { /* ignore */ }
    }
    return window.WPLDesktop ? "Mac" : "Phone";
  }

  function loadSyncCfg() {
    try {
      var c = JSON.parse(localStorage.getItem(SYNC_KEY) || "null");
      if (c) {
        if (typeof c.enabled !== "boolean") c.enabled = !!(c.host || c.code);
        return Object.assign({ enabled: false, host: "", code: "", auto: true, lastAt: 0, lastPeer: "", dirtyCount: 0 }, c);
      }
    } catch (err) { /* defaults */ }
    return { enabled: false, host: "", code: "", auto: true, lastAt: 0, lastPeer: "", dirtyCount: 0 };
  }

  function saveSyncCfg() {
    try { localStorage.setItem(SYNC_KEY, JSON.stringify(syncCfg)); } catch (err) { /* ignore */ }
  }

  function save() {
    if (!applyingRemote) stampChangedSettings();
    try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); }
    catch (err) { toast("Could not save. Your device storage is full."); }
    if (!applyingRemote) {
      /* Count changes made here rather than comparing timestamps: the two
         devices' clocks are never exactly aligned, and a day that arrived
         from the Mac must not look like something still owed to it. */
      syncCfg.dirtyCount = (syncCfg.dirtyCount || 0) + 1;
      saveSyncCfg();
      markCloudNeedsSync();
      pushToHost();
      cloudSoon();
    }
    backupIfDue();
    syncWorkReminderSchedule();
    scheduleWidgetRefresh();
    lastSettingsSnapshot = topLevelSettingsSnapshot(state.settings);
  }

  function topLevelSettingsSnapshot(settings) {
    var out = {};
    Object.keys(settings || {}).forEach(function (key) {
      try { out[key] = JSON.stringify(settings[key]); }
      catch (err) { out[key] = String(settings[key]); }
    });
    return out;
  }

  function stampChangedSettings() {
    var current = topLevelSettingsSnapshot(state.settings), keys = {};
    Object.keys(lastSettingsSnapshot || {}).forEach(function(k){keys[k]=true;});
    Object.keys(current).forEach(function(k){keys[k]=true;});
    var changed = Object.keys(keys).filter(function(key){return current[key] !== lastSettingsSnapshot[key];});
    if (!changed.length) return;
    var stamp = Math.max(Date.now(), Number(state.settingsUpdatedAt) || 0);
    state.settingsUpdatedAt = stamp;
    if (!state.settingsFieldUpdatedAt || typeof state.settingsFieldUpdatedAt !== "object") state.settingsFieldUpdatedAt = {};
    changed.forEach(function(key){state.settingsFieldUpdatedAt[key]=stamp;});
  }

  /* Live entries — tombstones stay in the store so a delete can travel to the
     other device, but nothing else in the app should ever see them. */
  function entries() {
    return state.entries.filter(function (e) { return !e.deleted; });
  }

  function uid() {
    return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  }

  /* Old pre-sync backups may not contain ids. A deterministic id keeps the
     same legacy shift from becoming two shifts when two devices import it. */
  function legacyEntryId(entry, index) {
    var text = [entry && entry.date, index || 0].join("|");
    var hash = 2166136261;
    for (var i = 0; i < text.length; i++) {
      hash ^= text.charCodeAt(i);
      hash = Math.imul(hash, 16777619);
    }
    return "legacy-" + (hash >>> 0).toString(36);
  }

  /* ---------- date helpers (all local, no UTC drift) ---------- */
  var DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  var MONTHS = ["January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"];

  function parseDate(iso) {
    if (!iso) return null;
    var p = iso.split("-");
    return new Date(+p[0], +p[1] - 1, +p[2]);
  }
  function toISO(d) {
    return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
  }
  function pad2(n) { return (n < 10 ? "0" : "") + n; }
  function addDays(d, n) { var x = new Date(d.getTime()); x.setDate(x.getDate() + n); return x; }
  function today() { var n = new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); }
  function dayDiff(a, b) { return Math.round((b - a) / 86400000); }
  function weekStart(d) { return addDays(d, -((d.getDay() + 6) % 7)); }   // Monday
  function fmtDate(d) { return pad2(d.getDate()) + "/" + pad2(d.getMonth() + 1) + "/" + d.getFullYear(); }
  function fmtDateLong(d) { return d.getDate() + " " + MONTHS[d.getMonth()] + " " + d.getFullYear(); }
  function fmtDateShort(d) { return d.getDate() + " " + MONTHS[d.getMonth()].slice(0, 3); }
  function dayName(d) { return DAYS[d.getDay()]; }

  /* ---------- money / hours formatting ---------- */
  function currencySymbol() {
    var st = state.settings || {};
    if (st.currency === "custom") return st.currencySymbol || "";
    for (var i = 0; i < WPL_SEED.currencies.length; i++) {
      if (WPL_SEED.currencies[i].code === st.currency) return WPL_SEED.currencies[i].symbol;
    }
    return (st.currency || "EUR") + " ";
  }
  function money(n) {
    var neg = n < 0;
    var s = Math.abs(n).toFixed(2);
    var parts = s.split(".");
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return (neg ? "-" : "") + currencySymbol() + parts.join(".");
  }
  function dec(n) { return n.toFixed(2); }

  /* INT(h)&"h "&TEXT(ROUND(frac*60,0),"00")&"m", "-" when zero */
  function hm(hours) {
    if (!hours) return "-";
    var h = Math.floor(hours);
    var m = Math.round((hours - h) * 60);
    if (m === 60) { h += 1; m = 0; }
    return h + "h " + pad2(m) + "m";
  }
  function hmPlain(hours) { return hm(hours) === "-" ? "0h 00m" : hm(hours); }

  function minutesOf(t) {
    if (!t) return null;
    var p = t.split(":");
    return (+p[0]) * 60 + (+p[1]);
  }

  /* ---------- core calculations ---------- */
  /* Excel's ROUND is half away from zero; the epsilon keeps binary floats
     such as 63.824999999999996 landing on the same cent Excel shows. */
  function xround(x, places) {
    var f = Math.pow(10, places);
    var eps = (x >= 0 ? 1 : -1) * 1e-9;
    return Math.round((x + eps) * f) / f;
  }
  function employerFor(entry) {
    var id = entry && entry.employerId || state.settings.defaultEmployerId;
    var list = state.settings.employers || [];
    return list.filter(function (x) { return x.id === id; })[0] || list[0] || {
      id: "default", name: "My job", color: "#e85d0f", rate: state.settings.rate,
      start: state.settings.defaultStart, finish: state.settings.defaultFinish,
      breakHours: state.settings.breakHours
    };
  }
  var entryCalcCache = {};
  function calculateEntry(entry) {
    entry = entry || {};
    var key = [entry.start || "", entry.finish || "", entry.leave || "",
      entry.employerId || "", entry.rateSnapshot, entry.breakHours, JSON.stringify(entry.breaks || []),entry.payMultiplier,entry.rateType,
      entry.updatedAt || 0, state.settingsUpdatedAt || 0].join("|");
    if (entry.id && entryCalcCache[entry.id] && entryCalcCache[entry.id].key === key) {
      return entryCalcCache[entry.id];
    }
    var s = minutesOf(entry.start), f = minutesOf(entry.finish), h = 0;
    var shiftHours = 0, overnight = false;
    var employer = employerFor(entry);
    var breakItems = Array.isArray(entry.breaks) ? entry.breaks.map(function(item){return {duration:Math.max(0,Number(item.duration)||0),paid:!!item.paid};}).filter(function(item){return item.duration>0;}) : null;
    var breakHours = breakItems ? breakItems.reduce(function(sum,item){return sum+(item.paid?0:item.duration);},0)
      : typeof entry.breakHours === "number" ? Math.max(0, entry.breakHours) : Math.max(0, Number(employer.breakHours) || 0);
    var paidBreakHours = breakItems ? breakItems.reduce(function(sum,item){return sum+(item.paid?item.duration:0);},0) : 0;
    var totalBreakHours = breakHours + paidBreakHours;
    var hourlyRate = typeof entry.rateSnapshot === "number"
      ? Math.max(0, entry.rateSnapshot) : Math.max(0, Number(employer.rate) || 0);
    var payMultiplier=entry.payMultiplier==null?1:Math.min(10,Math.max(0,Number(entry.payMultiplier)||0));
    if (s !== null && f !== null) {
      if (f < s) { f += 24 * 60; overnight = true; }
      shiftHours = (f - s) / 60;
      h = xround(Math.max(0, shiftHours - breakHours), 4);
    }
    var pct = entry.leave ? leavePct(entry.leave) : 100;
    var result = { key: key, hours: h,
      shiftHours: xround(shiftHours, 4), breakHours: xround(breakHours,4), paidBreakHours:xround(paidBreakHours,4), totalBreakHours:xround(totalBreakHours,4),
      invalidBreak: shiftHours > 0 && totalBreakHours > shiftHours,
      overnight: overnight, hourlyRate: hourlyRate,
      payMultiplier:payMultiplier,pay: xround(h * hourlyRate * payMultiplier * (pct / 100), 2) };
    if (entry.id) entryCalcCache[entry.id] = result;
    return result;
  }
  function hoursOf(entry) {
    return calculateEntry(entry).hours;
  }
  function payOf(entry) {
    return calculateEntry(entry).pay;
  }
  function sorted() {
    return entries().sort(function (a, b) {
      return a.date < b.date ? -1 : a.date > b.date ? 1
        : (a.start || "").localeCompare(b.start || "") || String(a.id || "").localeCompare(String(b.id || ""));
    });
  }
  function inRange(fromISO, toISOd) {
    return entries().filter(function (e) { return e.date >= fromISO && e.date <= toISOd; });
  }
  function countsAsWorkedDay(entry, calculatedHours, todayISO) {
    /* A saved future/repeating shift has hours, but it has not been worked yet.
       Keep planned shifts in the timesheet and forecasts while excluding them
       from every figure labelled "Days worked". Old entries without a status
       are completed work, matching the migration used elsewhere in the app. */
    return !isLeave(entry) && calculatedHours > 0 &&
      entry.date <= todayISO && (entry.status || "completed") !== "planned";
  }
  function totalsOf(list) {
    var h = 0, p = 0, leavePay = 0, workedDates = {}, leaveDates = {};
    var todayISO = toISO(today());
    list.forEach(function (e) {
      var eh = hoursOf(e);
      h += eh; p += payOf(e);
      if (isLeave(e)) { leaveDates[e.date] = true; leavePay += payOf(e); }
      else if (countsAsWorkedDay(e, eh, todayISO)) workedDates[e.date] = true; // two shifts are still one worked day
    });
    return {
      hours: xround(h, 4), pay: xround(p, 2), days: Object.keys(workedDates).length,
      leaveDays: Object.keys(leaveDates).length, leavePay: xround(leavePay, 2)
    };
  }
  function grandTotals() { return totalsOf(entries()); }

  /* Weekly summary: Monday–Sunday weeks spanning every logged day. */
  function weeks() {
    var list = sorted();
    if (!list.length) return [];
    var cur = weekStart(parseDate(list[0].date));
    var last = weekStart(parseDate(list[list.length - 1].date));
    var out = [], byWeek = {};
    list.forEach(function (e) {
      var key = toISO(weekStart(parseDate(e.date)));
      (byWeek[key] || (byWeek[key] = [])).push(e);
    });
    while (cur <= last && out.length < 520) {
      var end = addDays(cur, 6);
      var t = totalsOf(byWeek[toISO(cur)] || []);
      out.push({ start: cur, end: end, hours: t.hours, pay: t.pay, days: t.days });
      cur = addDays(cur, 7);
    }
    return out;
  }

  /* Pay periods: first payday, then every `cycleDays`; the period is the
     `cycleDays` days ending on the payday itself. */
  function periods() {
    var first = parseDate(state.settings.firstPayday);
    var cycle = Math.max(1, parseInt(state.settings.cycleDays, 10) || 28);
    if (!first) return [];
    var list = sorted();
    var lastEntry = list.length ? parseDate(list[list.length - 1].date) : today();
    var limit = lastEntry > today() ? lastEntry : today();
    var out = [], payday = first, pointer = 0;
    while (out.length < 6 || (payday <= limit && out.length < 400)) {
      var start = addDays(payday, -cycle + 1);
      var startISO = toISO(start), paydayISO = toISO(payday), periodEntries = [];
      while (pointer < list.length && list[pointer].date < startISO) pointer++;
      while (pointer < list.length && list[pointer].date <= paydayISO) {
        periodEntries.push(list[pointer]); pointer++;
      }
      var t = totalsOf(periodEntries);
      out.push({
        payday: payday, start: start, end: payday,
        days: t.days, hours: t.hours, pay: t.pay
      });
      payday = addDays(payday, cycle);
    }
    return out;
  }

  function nextPeriod(existing) {
    var t = today(), ps = existing || periods();
    /* A payday is still the current payday until that day has finished. Apart
       from making the wording more natural, this keeps reminders, analytics
       and take-home pay on the same cycle on payday itself. */
    for (var i = 0; i < ps.length; i++) if (ps[i].payday >= t) return ps[i];
    return ps.length ? ps[ps.length - 1] : null;
  }

  /* ============================================================
     RENDERING
     ============================================================ */
  var $ = function (id) { return document.getElementById(id); };
  var VIEW_META = {
    overview: "Overview",
    timesheet: "Timesheet",
    analytics: "Analytics",
    takehome: "Take-home pay",
    spending: "Spending plan",
    weekly: "Weekly summary",
    periods: "Pay periods",
    settings: "Settings",
    history: "Version history"
  };
  var PHONE_VIEWS = ["overview", "timesheet", "analytics", "takehome", "spending", "settings"];
  var currentView = "overview";
  /* Entrance animations are for discovering a page, not for every tab click.
     Replaying them made the Overview pay gradient fade through the dark page
     whenever someone returned from Timesheet or another tab. */
  var visitedViews = { overview: true };

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }

  var phoneTitleTimer = null;
  function setPhoneTabPosition(position, tracking) {
    var tabbar = $("tabbar");
    if (!tabbar || position < 0) return;
    tabbar.classList.toggle("is-swipe-tracking", !!tracking);
    tabbar.style.setProperty("--tab-position", String(position));
  }

  function animatePhoneTitle(previous, next) {
    var from = PHONE_VIEWS.indexOf(previous), to = PHONE_VIEWS.indexOf(next);
    var title = $("viewTitle");
    title.classList.remove("phone-title-next", "phone-title-previous");
    clearTimeout(phoneTitleTimer);
    if (from < 0 || to < 0 || from === to || !window.matchMedia ||
        !window.matchMedia("(max-width: 820px)").matches || reduceMotion) return;
    void title.offsetWidth;
    title.classList.add(to > from ? "phone-title-next" : "phone-title-previous");
    phoneTitleTimer = setTimeout(function () {
      title.classList.remove("phone-title-next", "phone-title-previous");
    }, 210);
  }

  function showView(name) {
    if (!VIEW_META[name] || !$('view-' + name)) return;
    var previous = currentView;
    currentView = name;
    var returning = !!visitedViews[name];
    /* Prepare the destination while it is still hidden. Android WebView could
       otherwise paint one frame of its stale shell before tables, totals and
       charts were filled, which looked like a flicker at the end of a swipe. */
    Object.keys(VIEW_META).forEach(function (v) {
      var view = $("view-" + v);
      if (v === name) view.classList.toggle("view-revisit", returning);
    });
    render();
    Object.keys(VIEW_META).forEach(function (v) {
      $("view-" + v).classList.toggle("hidden", v !== name);
    });
    visitedViews[name] = true;
    [].forEach.call(document.querySelectorAll(".nav-item,.tab"), function (b) {
      b.classList.toggle("active", b.dataset.view === name);
    });
    $("viewTitle").textContent = VIEW_META[name];
    setPhoneTabPosition(PHONE_VIEWS.indexOf(name), false);
    animatePhoneTitle(previous, name);
    $("backBtn").classList.toggle("hidden", name !== "history");
    $("scroll").scrollTop = 0;
    $("scroll").scrollLeft = 0;
  }

  function render() {
    applyAppColor();
    if (currentView === "overview") renderOverview();
    else if (currentView === "timesheet") renderTimesheet();
    else if (currentView === "analytics") renderAnalytics();
    else if (currentView === "takehome") renderTakeHome();
    else if (currentView === "spending") renderSpending();
    else if (currentView === "weekly") renderWeekly();
    else if (currentView === "periods") renderPeriods();
    else if (currentView === "settings") renderSettings();
    else if (currentView === "history") renderHistory();
    $("brandSub").textContent = money(grandTotals().pay) + " logged";
  }

  /* Counts a figure up rather than snapping to it. Short, eased, and skipped
     entirely when the OS asks for reduced motion or the jump is trivial. */
  var reduceMotion = window.matchMedia &&
    window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  var countTokens = {};
  function countTo(el, value, fmt) {
    var from = parseFloat(el.getAttribute("data-v") || "0");
    el.setAttribute("data-v", value);

    /* A frozen animation must never leave a wrong figure on screen, and
       requestAnimationFrame stops dead in a backgrounded app. So: skip the
       animation when hidden, cancel superseded runs, and always land the exact
       value with a timer that survives being paused. */
    if (reduceMotion || document.hidden || Math.abs(value - from) < 0.005) {
      el.textContent = fmt(value);
      return;
    }

    var token = (countTokens[el.id] = (countTokens[el.id] || 0) + 1);
    var start = 0, dur = 620;

    function step(ts) {
      if (countTokens[el.id] !== token) return;      // a newer value took over
      if (!start) start = ts;
      var p = Math.min(1, (ts - start) / dur);
      var eased = 1 - Math.pow(1 - p, 3);
      el.textContent = fmt(from + (value - from) * eased);
      if (p < 1) requestAnimationFrame(step);
    }
    requestAnimationFrame(step);
    setTimeout(function () {
      if (countTokens[el.id] === token) el.textContent = fmt(value);
    }, dur + 90);
  }

  /* Eight weeks of hours, so the shape of the run is visible at a glance. */
  function renderChart() {
    var ws = weeks().slice(-8);
    var el = $("ovChart");
    if (!el) return;
    if (!ws.length) { el.innerHTML = '<p class="muted tiny">Nothing logged yet.</p>'; return; }

    var peak = Math.max.apply(null, ws.map(function (w) { return w.hours; }).concat([1]));
    var curStart = toISO(weekStart(today()));
    var total = 0;

    el.innerHTML = ws.map(function (w, i) {
      total += w.pay;
      var h = Math.max(3, Math.round((w.hours / peak) * 82));
      var isNow = toISO(w.start) === curStart;
      return '<div class="chart-col' + (isNow ? " is-now" : "") + (w.hours === 0 ? " is-empty" : "") + '">' +
        '<span class="chart-val">' + (w.hours ? dec(w.hours) : "") + '</span>' +
        '<div class="chart-bar" style="height:' + h + 'px;animation-delay:' + (i * 45) + 'ms"></div>' +
        '<span class="chart-lab">' + fmtDateShort(w.start).replace(" ", " ") + '</span>' +
        '</div>';
    }).join("");
    $("ovChartTotal").textContent = money(xround(total, 2));
  }

  /* A running count of leave taken, by type — the thing a timesheet alone
     never tells you when someone asks how much annual leave you have left. */
  function renderLeave() {
    var card = $("leaveCard");
    if (!card) return;
    var counts = {}, order = [], paid = 0;

    entries().forEach(function (e) {
      if (!isLeave(e)) return;
      if (!counts[e.leave]) { counts[e.leave] = { days: 0, pay: 0 }; order.push(e.leave); }
      counts[e.leave].days++;
      counts[e.leave].pay += payOf(e);
      paid += payOf(e);
    });

    card.classList.toggle("hidden", !order.length);
    if (!order.length) return;

    var total = order.reduce(function (n, k) { return n + counts[k].days; }, 0);
    $("leaveTotal").textContent = total + (total === 1 ? " day" : " days");
    order.sort(function (a, b) { return counts[b].days - counts[a].days; });
    $("leaveList").innerHTML = order.map(function (k) {
      var c = counts[k], info = leaveInfo(k);
      return '<div class="leave-row"><span>' + esc(info.label) + "</span>" +
        '<span><span class="n">' + c.days + (c.days === 1 ? " day" : " days") + "</span>" +
        '<span class="amt"> · ' + money(xround(c.pay, 2)) + "</span></span></div>";
    }).join("");
  }

  function runningShift() {
    return entries().filter(function(e){return !e.leave && !!e.start && !e.finish;})
      .sort(function(a,b){return (b.date+" "+b.start).localeCompare(a.date+" "+a.start);})[0] || null;
  }

  function runningHours(entry) {
    if (!entry || !entry.start) return 0;
    var date=parseDate(entry.date),start=minutesOf(entry.start);
    if (!date || start===null) return 0;
    return Math.max(0,(Date.now()-(date.getTime()+start*60000))/3600000);
  }

  function renderClockAction(running) {
    if (!$("qClock")) return;
    running=running||runningShift();
    $("qClockLabel").textContent=running?"Clock out":"Clock in";
    $("qClockTime").textContent=running?hmPlain(runningHours(running))+" so far":"Start a shift now";
    $("qClock").classList.toggle("is-running",!!running);
  }

  function toggleClock() {
    var running=runningShift(),now=new Date(),nowTime=pad2(now.getHours())+":"+pad2(now.getMinutes());
    if (running) {
      if (dayDiff(parseDate(running.date),today())>1) { openEditor(running.id);toast("Check this older shift before clocking out");return; }
      var before=snapshotOf(running);
      running.finish=nowTime;running.status="completed";running.updatedAt=Date.now();
      var calc=calculateEntry(running);
      if (calc.invalidBreak) running.breakHours=0;
      recordHistory("day",running.date,before,snapshotOf(running),"clock-out");
      save();render();toast("Clocked out · "+hm(hoursOf(running)));
      return;
    }
    var point=Date.now(),inside=entries().some(function(e){var interval=shiftInterval(e);return interval&&point>=interval.start&&point<interval.end;});
    if (inside) { toast("A saved shift is already running at this time");return; }
    var employer=employerFor({employerId:state.settings.defaultEmployerId});
    var rec={id:uid(),date:toISO(today()),start:nowTime,finish:"",notes:"",leave:"",
      employerId:employer.id,rateSnapshot:Math.max(0,Number(employer.rate)||0),
      breakHours:Math.max(0,Number(employer.breakHours)||0),status:"planned",payMultiplier:1,rateType:"standard",updatedAt:Date.now(),deleted:false};
    state.entries.push(rec);recordHistory("day",rec.date,null,snapshotOf(rec),"clock-in");
    save();render();toast("Clocked in at "+nowTime);
  }

  /* ---------- overview ---------- */
  function renderOverview() {
    var g = grandTotals();
    countTo($("ovTotalPay"), g.pay, money);
    renderChart();
    renderLeave();
    $("ovTotalTime").textContent = hmPlain(g.hours);
    $("ovTotalHours").textContent = dec(g.hours);
    $("ovDays").textContent = g.days;

    var np = nextPeriod();
    if (np) {
      var d = dayDiff(today(), np.payday);
      $("ovNextPayday").textContent = fmtDateLong(np.payday);
      $("ovNextPaydayDay").textContent = DAYS[np.payday.getDay()] + " · every " +
        state.settings.cycleDays + " days";
      $("ovCountdown").textContent = d > 1 ? "in " + d + " days" : d === 1 ? "tomorrow" : d === 0 ? "today" : "past";
      $("ovPayDue").textContent = money(np.pay);
      $("ovPeriodRange").textContent = fmtDateShort(np.start) + " – " + fmtDateShort(np.end);
      $("ovPeriodHours").textContent = hmPlain(np.hours) + " · " + np.days + " days";
    }

    var ws = weekStart(today()), we = addDays(ws, 6);
    var wt = totalsOf(inRange(toISO(ws), toISO(we)));
    $("ovWeekRange").textContent = fmtDateShort(ws) + " – " + fmtDateShort(we);
    $("ovWeekPay").textContent = money(wt.pay);
    $("ovWeekTime").textContent = hmPlain(wt.hours) + " over " + wt.days + (wt.days === 1 ? " day" : " days");
    var pct = Math.min(100, (wt.hours / fullWeek()) * 100);
    $("ovWeekBar").style.width = pct + "%";
    $("ovWeekBarLabel").textContent = dec(wt.hours) + " of " + dec(fullWeek()) + " h";

    var todayEntries = entries().filter(function (e) { return e.date === toISO(today()); });
    var todayTotal = totalsOf(todayEntries), running = runningShift();
    $("ovToday").textContent = !todayEntries.length ? "Not logged"
      : running ? "Started " + running.start
      : todayEntries.length > 1 ? todayEntries.length + " shifts · " + hm(todayTotal.hours) + " · " + money(todayTotal.pay)
      : isLeave(todayEntries[0]) ? leaveInfo(todayEntries[0].leave).label + (hoursOf(todayEntries[0]) > 0 ? " · " + money(payOf(todayEntries[0])) : "")
      : hoursOf(todayEntries[0]) > 0 ? todayEntries[0].start + "–" + todayEntries[0].finish + " · " + money(payOf(todayEntries[0])) : "Day off";
    renderClockAction(running);
    var takeHome=currentTakeHomeModel(),spending=spendingModel();
    $("ovTakeHome").textContent=money(takeHome.netMonth||0);
    $("ovSafeSpend").textContent=money(spending.safeToSpend||0);
    $("ovIncomplete").textContent=entries().filter(function(entry){return !!entry.start!==!!entry.finish;}).length;
    var cloudOverview=(window.WPLCloud&&WPLCloud.cfg)?WPLCloud.cfg():{};
    $("ovCloudStatus").textContent=cloudOverview.lastAt?"Saved "+fmtAgo(cloudOverview.lastAt):(window.WPLCloud&&WPLCloud.signedIn()?"Waiting":"Off");
    $("ovBackupStatus").textContent=backupCfg.lastAt?fmtAgo(backupCfg.lastAt):"Not yet";

    /* "Recent" means days that have actually happened. The timesheet can run
       ahead of today — the spreadsheet was filled in months in advance — and
       listing next month's shifts under "Recent days" is just wrong. */
    var iso = toISO(today());
    var all = sorted();
    var past = all.filter(function (e) { return e.date <= iso; });
    var showing = past.length ? past.slice(-6).reverse() : all.slice(0, 6);

    $("ovRecentTitle").textContent = past.length ? "Recent days" : "Upcoming days";
    $("ovRecent").innerHTML = showing.length
      ? showing.map(function (e) { return rowHTML(e, true); }).join("")
      : '<p class="muted tiny">No days logged yet.</p>';
    bindRows($("ovRecent"));
  }

  /* ---------- timesheet ---------- */
  var tsSelected = {};
  var tsMode = "list";
  var calendarAnchor = null;
  function rowHTML(e, compact) {
    var d = parseDate(e.date), h = hoursOf(e), p = payOf(e);
    var off = h === 0;
    var isToday = e.date === toISO(today());
    var info = leaveInfo(e.leave);
    var status = ["planned", "completed", "approved", "paid"].indexOf(e.status) !== -1 ? e.status : "completed";

    var main, sub;
    if (isLeave(e)) {
      main = info.label + (h > 0 || !info.paid ? "" : " · unpaid");
      sub = (h > 0 ? e.start + " – " + e.finish + " · " : "") + (e.notes || fmtDate(d));
    } else {
      main = off ? (e.notes || "Day off") : e.start + " – " + e.finish;
      sub = off ? fmtDate(d) : (e.notes ? e.notes : fmtDate(d));
    }

    return '<div class="row-wrap' + (compact ? " compact" : "") + (isToday ? " has-today" : "") + (tsSelected[e.id] ? " is-selected" : "") + '">' +
      (compact ? "" : '<button class="row-check" data-select="' + e.id + '" aria-label="Select ' + esc(fmtDate(d)) + '">' +
        (tsSelected[e.id] ? "✓" : "") + '</button>') +
      '<button class="row' + (off ? " is-off" : "") + (isToday ? " is-today" : "") +
      '" data-id="' + e.id + '">' +
      '<span class="daybox"><b>' + d.getDate() + '</b><i>' + dayName(d) + '</i></span>' +
      '<span class="row-main"><span class="row-times">' + esc(main) +
      (isLeave(e) ? ' <span class="leave-tag tone-' + info.tone + '">leave</span>' : "") +
      (status !== "completed" ? ' <span class="shift-status status-' + status + '">' + status + '</span>' : "") +
      '</span><span class="row-sub">' + esc(sub) + '</span></span>' +
      '<span class="row-end"><span class="row-pay">' + money(p) + '</span>' +
      '<span class="row-hrs">' + hm(h) + '</span></span>' +
      (compact ? "" : '<span class="chev">›</span>') +
      '</button>' + (compact ? "" : '<button class="row-expand" data-expand="' + e.id + '" aria-label="Show shift details">⌄</button><div class="row-detail" data-detail="' + e.id + '"><span><b>Employer</b>' + esc(employerFor(e).name) + '</span><span><b>Date</b>' + esc(fmtDateLong(d)) + '</span><span><b>Type</b>' + esc(info.label) + '</span><span><b>Shift</b>' + esc(e.start && e.finish ? e.start + " – " + e.finish : "No complete shift") + '</span><span><b>Hours</b>' + hm(h) + '</span><span><b>Gross pay</b>' + money(p) + '</span>' + (e.notes ? '<span class="detail-note"><b>Notes</b>' + esc(e.notes) + '</span>' : "") + '</div>') + '</div>';
  }

  function bindRows(scope) {
    [].forEach.call(scope.querySelectorAll(".row"), function (r) {
      r.onclick = function () { openEditor(r.dataset.id); };
    });
    [].forEach.call(scope.querySelectorAll(".row-check"), function (b) {
      b.onclick = function (ev) {
        ev.stopPropagation();
        if (tsSelected[b.dataset.select]) delete tsSelected[b.dataset.select];
        else tsSelected[b.dataset.select] = true;
        renderTimesheet();
      };
    });
    [].forEach.call(scope.querySelectorAll(".row-expand"),function(b){b.onclick=function(ev){ev.stopPropagation();var d=scope.querySelector('[data-detail="'+b.dataset.expand+'"]');if(d){d.classList.toggle("open");b.classList.toggle("open");b.setAttribute("aria-label",d.classList.contains("open")?"Hide shift details":"Show shift details");}};});
  }

  function filteredTimesheet() {
    var q = ($("tsSearch").value || "").trim().toLowerCase();
    var filter = $("tsFilter").value, from = $("tsFrom").value, to = $("tsTo").value;
    var employer = $("tsEmployer").value || "all";
    var list = sorted().filter(function (e) {
      var h = hoursOf(e), incomplete = (!!e.start !== !!e.finish), pct = leavePct(e.leave);
      if (from && e.date < from || to && e.date > to) return false;
      if (employer !== "all" && e.employerId !== employer) return false;
      if (filter === "worked" && (h === 0 || isLeave(e))) return false;
      if (filter === "off" && e.leave !== "off") return false;
      if (filter === "unpaid" && (!isLeave(e) || pct > 0 || e.leave === "off")) return false;
      if (filter === "leave" && !isLeave(e)) return false;
      if (filter === "incomplete" && !incomplete) return false;
      if (["planned", "approved", "paid"].indexOf(filter) !== -1 && (e.status || "completed") !== filter) return false;
      if (!q) return true;
      var d = parseDate(e.date), emp = employerFor(e);
      var hay = (fmtDate(d) + " " + dayName(d) + " " + MONTHS[d.getMonth()] + " " +
        (e.notes || "") + " " + e.start + " " + e.finish + " " + emp.name + " " +
        leaveInfo(e.leave).label).toLowerCase();
      return hay.indexOf(q) !== -1;
    });
    var sort = $("tsSort").value;
    list.sort(function (a, b) {
      if (sort === "date-asc") return a.date.localeCompare(b.date);
      if (sort === "hours-desc") return hoursOf(b) - hoursOf(a) || b.date.localeCompare(a.date);
      if (sort === "pay-desc") return payOf(b) - payOf(a) || b.date.localeCompare(a.date);
      return b.date.localeCompare(a.date);
    });
    return list;
  }

  function renderTimesheet() {
    fillEmployerSelects();
    var list = filteredTimesheet();

    var t = totalsOf(list);
    $("tsDays").textContent = t.days;
    $("tsHours").textContent = dec(t.hours);
    $("tsTime").textContent = hmPlain(t.hours);
    $("tsPay").textContent = money(t.pay);

    $("tsEmpty").classList.toggle("hidden", list.length > 0);
    $("tsBulkBar").classList.toggle("hidden", !Object.keys(tsSelected).length);
    $("tsSelectedCount").textContent = Object.keys(tsSelected).length + " selected";
    $("tsList").classList.toggle("hidden", tsMode !== "list");
    $("tsCalendar").classList.toggle("hidden", tsMode !== "calendar");

    /* group by month, newest month first, days ascending inside */
    var groups = {}, order = [];
    list.slice().sort(function (a,b) { return a.date.localeCompare(b.date); }).forEach(function (e) {
      var k = e.date.slice(0, 7);
      if (!groups[k]) { groups[k] = []; order.push(k); }
      groups[k].push(e);
    });
    order.sort().reverse();

    $("tsList").innerHTML = order.map(function (k) {
      var g = groups[k], gt = totalsOf(g);
      var d = parseDate(k + "-01");
      return '<div class="month"><div class="month-head"><h3>' +
        MONTHS[d.getMonth()] + " " + d.getFullYear() + '</h3><span>' +
        hmPlain(gt.hours) + " · " + money(gt.pay) + '</span></div>' +
        '<div class="rows">' + g.map(function (e) { return rowHTML(e); }).join("") + "</div></div>";
    }).join("");
    bindRows($("tsList"));
    if (tsMode === "calendar") renderCalendar(list);
  }

  function renderCalendar(list) {
    var anchor = calendarAnchor || ($("tsTo").value ? parseDate($("tsTo").value) : today());
    var first = new Date(anchor.getFullYear(), anchor.getMonth(), 1);
    calendarAnchor = first;
    var last = new Date(anchor.getFullYear(), anchor.getMonth() + 1, 0);
    var byDate = {};
    list.forEach(function (e) { (byDate[e.date] || (byDate[e.date] = [])).push(e); });
    var html = '<div class="calendar-head"><div class="calendar-picker">' +
      '<button class="cal-nav" data-cal-move="-1" aria-label="Previous month">‹</button>' +
      '<select class="cal-month-pick" aria-label="Calendar month">' + MONTHS.map(function (m, i) {
        return '<option value="' + i + '"' + (i === first.getMonth() ? " selected" : "") + '>' + m + '</option>';
      }).join("") + '</select><input class="cal-year-pick" type="number" min="1900" max="2200" inputmode="numeric" value="' + first.getFullYear() + '" aria-label="Calendar year">' +
      '<button class="cal-nav" data-cal-move="1" aria-label="Next month">›</button>' +
      '<button class="cal-today" data-cal-today>Today</button></div><span>Tap a day to log or edit it</span></div><div class="calendar-grid">' +
      ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"].map(function (x) { return '<b class="cal-label">' + x + '</b>'; }).join("");
    var pad = (first.getDay() + 6) % 7;
    for (var p = 0; p < pad; p++) html += '<span class="cal-blank"></span>';
    for (var day = 1; day <= last.getDate(); day++) {
      var date = new Date(first.getFullYear(), first.getMonth(), day), iso = toISO(date), dayEntries = byDate[iso] || [];
      var dayTotals = totalsOf(dayEntries);
      var incomplete = dayEntries.some(function(e){return !!e.start !== !!e.finish;});
      var planned = dayEntries.some(function(e){return (e.status || "completed") === "planned";});
      var paid = dayEntries.length && dayEntries.every(function(e){return (e.status || "completed") === "paid";});
      var worked = dayEntries.some(function(e){return !e.leave && hoursOf(e)>0;});
      var firstLeave = dayEntries.filter(function(e){return !!e.leave;})[0];
      var tone = !dayEntries.length ? "empty" : incomplete ? "incomplete" : paid ? "paid" : planned ? "planned" : worked ? "work" : firstLeave ? leaveInfo(firstLeave.leave).tone : "off";
      var summary = !dayEntries.length ? "Add" : dayEntries.length > 1
        ? dayEntries.length + " shifts · " + hm(dayTotals.hours)
        : firstLeave ? esc(leaveInfo(firstLeave.leave).label) : hm(dayTotals.hours);
      html += '<button class="cal-day tone-' + tone + '" data-date="' + iso + '" data-id="' + (dayEntries.length === 1 ? dayEntries[0].id : "") +
        '" data-count="' + dayEntries.length + '" aria-label="' + esc(fmtDateLong(date) + (dayEntries.length ? ", " + dayEntries.length + (dayEntries.length === 1 ? " shift" : " shifts") : "")) + '"><strong>' + day + '</strong><span>' +
        summary + '</span></button>';
    }
    $("tsCalendar").innerHTML = html + '</div>';
    [].forEach.call($("tsCalendar").querySelectorAll("[data-cal-move]"), function (b) {
      b.onclick = function () {
        calendarAnchor = new Date(calendarAnchor.getFullYear(), calendarAnchor.getMonth() + Number(b.dataset.calMove), 1);
        renderCalendar(filteredTimesheet());
      };
    });
    $("tsCalendar").querySelector(".cal-month-pick").onchange = function () {
      calendarAnchor = new Date(calendarAnchor.getFullYear(), Number(this.value), 1);
      renderCalendar(filteredTimesheet());
    };
    $("tsCalendar").querySelector(".cal-year-pick").onchange = function () {
      var year = Math.max(1900, Math.min(2200, parseInt(this.value, 10) || today().getFullYear()));
      calendarAnchor = new Date(year, calendarAnchor.getMonth(), 1);
      renderCalendar(filteredTimesheet());
    };
    $("tsCalendar").querySelector("[data-cal-today]").onclick = function () {
      calendarAnchor = new Date(today().getFullYear(), today().getMonth(), 1);
      renderCalendar(filteredTimesheet());
    };
    [].forEach.call($("tsCalendar").querySelectorAll(".cal-day"), function (b) {
      b.onclick = function () {
        if (b.dataset.id) { openEditor(b.dataset.id); return; }
        if (Number(b.dataset.count) > 1) {
          $("tsFrom").value = b.dataset.date; $("tsTo").value = b.dataset.date;
          tsMode = "list"; $("tsListMode").classList.add("on"); $("tsCalendarMode").classList.remove("on");
          renderTimesheet(); toast("Showing all shifts for " + fmtDate(parseDate(b.dataset.date)));
          return;
        }
        openEditor(null); $("edDate").value = b.dataset.date;
      };
    });
  }

  /* ---------- analytics ---------- */
  var analyticsRange = "month";
  var analyticsSort = { weeks: -1, periods: -1, months: -1 };
  var analyticsModel = null;

  function analyticsDates() {
    var now = today(), from, to = now;
    if (analyticsRange === "week") from = weekStart(now);
    else if (analyticsRange === "month") from = new Date(now.getFullYear(), now.getMonth(), 1);
    else if (analyticsRange === "3months") from = new Date(now.getFullYear(), now.getMonth() - 2, 1);
    else if (analyticsRange === "year") from = new Date(now.getFullYear(), 0, 1);
    else if (analyticsRange === "all") {
      var all = sorted(); from = all.length ? parseDate(all[0].date) : now;
      to = all.length && parseDate(all[all.length - 1].date) > now ? parseDate(all[all.length - 1].date) : now;
    } else {
      from = parseDate($("anFrom").value) || new Date(now.getFullYear(), now.getMonth(), 1);
      to = parseDate($("anTo").value) || now;
    }
    if (from > to) { var swap = from; from = to; to = swap; }
    return { from: from, to: to, fromISO: toISO(from), toISO: toISO(to) };
  }

  function analyticsEntries(range) {
    var emp = $("anEmployer").value || "all";
    return entries().filter(function (e) {
      return e.date >= range.fromISO && e.date <= range.toISO && (emp === "all" || e.employerId === emp);
    });
  }

  function previousRange(range) {
    var days = dayDiff(range.from, range.to) + 1;
    var to = addDays(range.from, -1), from = addDays(to, -days + 1);
    return { from: from, to: to, fromISO: toISO(from), toISO: toISO(to) };
  }

  function deltaText(now, before, noun) {
    var diff = xround(now - before, noun === "pay" ? 2 : 1);
    var pct = before ? Math.round((diff / Math.abs(before)) * 100) : (now ? 100 : 0);
    var mark = diff > 0 ? "↑" : diff < 0 ? "↓" : "→";
    var value = noun === "pay" ? money(Math.abs(diff)) : Math.abs(diff).toFixed(noun === "days" ? 0 : 1);
    return mark + " " + value + " · " + (pct > 0 ? "+" : "") + pct + "% vs previous";
  }

  function groupWeeks(list) {
    var map = {};
    list.forEach(function (e) {
      var start = weekStart(parseDate(e.date)), key = toISO(start);
      if (!map[key]) map[key] = { start: start, end: addDays(start, 6), entries: [] };
      map[key].entries.push(e);
    });
    return Object.keys(map).sort().map(function (k) {
      var w = map[k], t = totalsOf(w.entries), regular = 0, overtime = 0, leave = 0;
      w.entries.forEach(function (e) {
        var h = hoursOf(e);
        if (isLeave(e)) leave += h; else regular += h;
      });
      overtime = Math.max(0, regular - fullWeek());
      return Object.assign(w, t, { regular: xround(regular - overtime, 4), overtime: xround(overtime, 4), leaveHours: xround(leave, 4) });
    });
  }

  function groupMonths(list) {
    var map = {};
    list.forEach(function (e) { var k = e.date.slice(0, 7); (map[k] || (map[k] = [])).push(e); });
    return Object.keys(map).sort().map(function (k) {
      var t = totalsOf(map[k]), first = parseDate(k + "-01"), monthWeeks = groupWeeks(map[k]);
      var activeWeeks = Math.max(1, monthWeeks.length);
      return { key: k, date: first, entries: map[k], days: t.days, hours: t.hours,
        leaveDays: t.leaveDays, pay: t.pay, avg: xround(t.pay / activeWeeks, 2), weeks: monthWeeks };
    });
  }

  function analyticsPeriodModels(source, list, recalculateTotals) {
    var sortedList = list.slice().sort(function (a, b) { return a.date.localeCompare(b.date); });
    var pointer = 0;
    return source.map(function (period) {
      var from = toISO(period.start), to = toISO(period.end), periodEntries = [];
      while (pointer < sortedList.length && sortedList[pointer].date < from) pointer++;
      while (pointer < sortedList.length && sortedList[pointer].date <= to) {
        periodEntries.push(sortedList[pointer]); pointer++;
      }
      var paidLeaveHours = 0;
      periodEntries.forEach(function (e) {
        if (isLeave(e) && leavePct(e.leave) > 0) paidLeaveHours += hoursOf(e);
      });
      var result = Object.assign({}, period, { paidLeaveHours: xround(paidLeaveHours, 4) });
      return recalculateTotals ? Object.assign(result, totalsOf(periodEntries)) : result;
    });
  }

  function chartBars(items, value, label, formatter) {
    if (!items.length) return '<p class="empty-inline">No data in this period</p>';
    var max = Math.max.apply(null, items.map(value).concat([1]));
    return '<div class="bars">' + items.slice(-14).map(function (x) {
      var v = value(x), pct = Math.max(2, Math.round(v / max * 100));
      return '<div class="bar-col" title="' + esc(label(x) + ": " + formatter(v)) + '"><span class="bar-value">' +
        esc(formatter(v)) + '</span><i style="height:' + pct + '%"></i><b>' + esc(label(x)) + '</b></div>';
    }).join("") + '</div>';
  }

  function renderAnalytics() {
    fillEmployerSelects();
    var range = analyticsDates(), list = analyticsEntries(range), total = totalsOf(list);
    var prevList = analyticsEntries(previousRange(range)), prev = totalsOf(prevList);
    var compare = $("anCompare").checked;
    $("anHours").textContent = dec(total.hours); $("anPay").textContent = money(total.pay);
    var takeHomeNow=currentTakeHomeModel(),netRatio=takeHomeNow.grossMonth>0?Math.max(0,Math.min(1,takeHomeNow.netMonth/takeHomeNow.grossMonth)):1;
    $("anNet").textContent=money(total.pay*netRatio);
    $("anDays").textContent = total.days; $("anLeave").textContent = total.leaveDays;
    $("anAvgDay").textContent = money(total.days ? total.pay / total.days : 0);
    $("anHoursDelta").textContent = compare ? deltaText(total.hours, prev.hours, "hours") : "";
    $("anPayDelta").textContent = compare ? deltaText(total.pay, prev.pay, "pay") : "";
    $("anDaysDelta").textContent = compare ? deltaText(total.days, prev.days, "days") : "";
    var spanWeeks = Math.max(1, (dayDiff(range.from, range.to) + 1) / 7);
    $("anAvgWeek").textContent = hmPlain(total.hours / spanWeeks);
    $("anPatternRange").textContent = fmtDateShort(range.from) + " – " + fmtDateShort(range.to);

    var allPeriods = periods(), np = nextPeriod(allPeriods);
    $("anNextPay").textContent = np ? money(np.pay) : money(0);
    $("anNextPayDate").textContent = np ? fmtDateLong(np.payday) : "Not configured";

    var worked = list.filter(function (e) { return !isLeave(e) && hoursOf(e) > 0; });
    var starts = worked.map(function (e) { return minutesOf(e.start); }).filter(function (x) { return x !== null; });
    var finishes = worked.map(function (e) { return minutesOf(e.finish); }).filter(function (x) { return x !== null; });
    function avgTime(xs) { if (!xs.length) return "–"; var n = Math.round(xs.reduce(function (a,b) { return a+b; },0)/xs.length); return pad2(Math.floor(n/60)) + ":" + pad2(n%60); }
    var lengths = worked.map(hoursOf).filter(function (h) { return h > 0; });
    $("anAvgStart").textContent = avgTime(starts); $("anAvgFinish").textContent = avgTime(finishes);
    $("anAvgShift").textContent = lengths.length ? hmPlain(lengths.reduce(function(a,b){return a+b;},0)/lengths.length) : "–";
    $("anLongest").textContent = lengths.length ? hmPlain(Math.max.apply(null,lengths)) : "–";
    $("anShortest").textContent = lengths.length ? hmPlain(Math.min.apply(null,lengths)) : "–";
    var weekdays = {}, leaveCounts = {};
    worked.forEach(function(e){ var d=dayName(parseDate(e.date)); weekdays[d]=(weekdays[d]||0)+1; });
    list.filter(isLeave).forEach(function(e){ leaveCounts[e.leave]=(leaveCounts[e.leave]||0)+1; });
    function maxKey(obj){ return Object.keys(obj).sort(function(a,b){return obj[b]-obj[a];})[0]; }
    $("anBusiest").textContent = maxKey(weekdays) || "–";
    $("anCommonLeave").textContent = leaveCounts[maxKey(leaveCounts)] ? leaveInfo(maxKey(leaveCounts)).label : "–";

    var ws = groupWeeks(list), ms = groupMonths(list);
    $("anHoursChart").innerHTML = chartBars(ws,function(w){return w.hours;},function(w){return fmtDateShort(w.start);},function(v){return dec(v)+"h";});
    $("anPayChart").innerHTML = chartBars(ms.length > 1 ? ms : ws,function(x){return x.pay;},function(x){return x.date ? MONTHS[x.date.getMonth()].slice(0,3) : fmtDateShort(x.start);},money);
    $("anHoursChartTotal").textContent = hmPlain(total.hours); $("anPayChartTotal").textContent = money(total.pay);
    var wd = total.days, ld = total.leaveDays, sumDays = wd + ld, workedPct = sumDays ? Math.round(wd/sumDays*100) : 0;
    $("anDaysChart").innerHTML = '<div class="donut" style="--worked:' + workedPct + '%"><strong>' + sumDays + '</strong><span>logged days</span></div><div class="chart-legend"><span><i class="work-dot"></i>Worked ' + wd + '</span><span><i class="leave-dot"></i>Leave ' + ld + '</span></div>';
    $("anDaysChartTotal").textContent = sumDays + " days";

    $("anGrossNetChart").innerHTML=chartBars([{label:"Gross",value:total.pay},{label:"Take-home",value:total.pay*netRatio}],function(item){return item.value;},function(item){return item.label;},money);
    var byEmployer={};list.forEach(function(entry){var employer=employerFor(entry),key=employer.id;if(!byEmployer[key])byEmployer[key]={label:employer.name,value:0};byEmployer[key].value+=payOf(entry);});
    $("anEmployerChart").innerHTML=chartBars(Object.keys(byEmployer).map(function(key){return byEmployer[key];}),function(item){return item.value;},function(item){return item.label;},money);
    var weekdayOrder=["Mon","Tue","Wed","Thu","Fri","Sat","Sun"],byWeekday={};weekdayOrder.forEach(function(day){byWeekday[day]=0;});list.forEach(function(entry){byWeekday[dayName(parseDate(entry.date))]+=payOf(entry);});
    $("anWeekdayChart").innerHTML=chartBars(weekdayOrder.map(function(day){return {label:day,value:byWeekday[day]};}),function(item){return item.value;},function(item){return item.label;},money);

    var ps = allPeriods.filter(function(p){return toISO(p.end)>=range.fromISO&&toISO(p.start)<=range.toISO;});
    ps = analyticsPeriodModels(ps, list, $("anEmployer").value !== "all");
    $("anPeriodChart").innerHTML = chartBars(ps,function(p){return p.pay;},function(p){return fmtDateShort(p.payday);},money);
    $("anPeriodChartTotal").textContent = ps.length + (ps.length===1?" period":" periods");

    analyticsModel = { weeks: ws, periods: ps, months: ms, entries: list };
    renderAnalyticsTables(ws, ps, ms, list);
    renderProjections(range, list, total, np);
  }

  function renderAnalyticsTables(ws, ps, ms, list) {
    ws = ws.slice().sort(function(a,b){return analyticsSort.weeks*(a.start-b.start);});
    $("anWeeksBody").innerHTML = ws.map(function(w,i){var previous=i<ws.length-1?ws[i+1].pay:0;return '<tr class="analytics-summary-row"><td><button class="analytics-row-toggle" type="button" aria-expanded="false" aria-label="Show days in this week"><span>'+fmtDateShort(w.start)+' – '+fmtDateShort(w.end)+'</span><i aria-hidden="true">⌄</i></button></td><td>'+w.days+'</td><td>'+dec(w.regular)+'</td><td>'+dec(w.overtime)+'</td><td>'+dec(w.leaveHours)+'</td><td>'+money(w.pay)+'</td><td>'+deltaText(w.pay,previous,"pay").split(" · ")[0]+'</td></tr><tr class="detail-row"><td colspan="7">'+w.entries.sort(function(a,b){return a.date.localeCompare(b.date);}).map(function(e){return esc(fmtDateShort(parseDate(e.date))+" · "+leaveInfo(e.leave).label+" · "+hm(hoursOf(e))+" · "+money(payOf(e)));}).join("<br>")+'</td></tr>';}).join("") || emptyTable(7);
    ps = ps.slice().sort(function(a,b){return analyticsSort.periods*(a.payday-b.payday);});
    $("anPeriodsBody").innerHTML = ps.map(function(p){var key=toISO(p.payday),paid=!!state.settings.paidPeriods[key],diff=dayDiff(today(),p.payday);var status=paid?"Paid":diff<0?"Overdue":diff===0?"Due":"Upcoming";return '<tr><td>'+fmtDateShort(p.start)+' – '+fmtDateShort(p.end)+'</td><td>'+fmtDate(p.payday)+'</td><td>'+p.days+'</td><td>'+dec(p.hours)+'</td><td>'+hm(p.paidLeaveHours||0)+'</td><td>'+money(p.pay)+'</td><td><button class="status-btn" data-period="'+key+'">'+status+'</button></td></tr>';}).join("") || emptyTable(7);
    ms = ms.slice().sort(function(a,b){return analyticsSort.months*a.key.localeCompare(b.key);});
    $("anMonthsBody").innerHTML = ms.map(function(m){return '<tr class="analytics-summary-row"><td><button class="analytics-row-toggle" type="button" aria-expanded="false" aria-label="Show weeks in this month"><span>'+MONTHS[m.date.getMonth()]+' '+m.date.getFullYear()+'</span><i aria-hidden="true">⌄</i></button></td><td>'+m.days+'</td><td>'+dec(m.hours)+'</td><td>'+m.leaveDays+'</td><td>'+money(m.pay)+'</td><td>'+money(m.avg)+'</td></tr><tr class="detail-row"><td colspan="6">'+m.weeks.map(function(w){return esc(fmtDateShort(w.start)+" – "+fmtDateShort(w.end)+" · "+hm(w.hours)+" · "+money(w.pay));}).join("<br>")+'</td></tr>';}).join("") || emptyTable(6);
    var lm={};list.filter(isLeave).forEach(function(e){var k=e.leave;if(!lm[k])lm[k]={days:0,hours:0,pay:0};lm[k].days++;lm[k].hours+=hoursOf(e);lm[k].pay+=payOf(e);});
    var keys=Object.keys(lm).sort(function(a,b){return lm[b].days-lm[a].days;});
    $("anLeaveTypes").textContent=keys.length+" type"+(keys.length===1?"":"s");
    $("anLeaveBody").innerHTML=keys.map(function(k){return '<tr><td>'+esc(leaveInfo(k).label)+'</td><td>'+lm[k].days+'</td><td>'+dec(lm[k].hours)+'</td><td>'+leavePct(k)+'%</td><td>'+money(lm[k].pay)+'</td></tr>';}).join("")||emptyTable(5);
    [].forEach.call($("view-analytics").querySelectorAll(".analytics-row-toggle"),function(b){b.onclick=function(){var r=b.closest("tr"),d=r?r.nextElementSibling:null,open=d?d.classList.toggle("open"):false;b.classList.toggle("open",open);b.setAttribute("aria-expanded",open?"true":"false");b.setAttribute("aria-label",(open?"Hide":"Show")+(b.closest("#anMonthsBody")?" weeks in this month":" days in this week"));};});
    [].forEach.call($("anPeriodsBody").querySelectorAll(".status-btn"),function(b){b.onclick=function(){var k=b.dataset.period;state.settings.paidPeriods[k]=!state.settings.paidPeriods[k];state.settingsUpdatedAt=Date.now();save();if(analyticsModel)renderAnalyticsTables(analyticsModel.weeks,analyticsModel.periods,analyticsModel.months,analyticsModel.entries);toast(state.settings.paidPeriods[k]?"Marked as paid":"Marked as unpaid");};});
  }
  function emptyTable(cols){return '<tr><td colspan="'+cols+'" class="empty-cell">No data in this period</td></tr>';}

  function renderProjections(range,list,total,np){
    var observedEnd=range.to<today()?range.to:today();
    var elapsed=Math.max(1,dayDiff(range.from,observedEnd)+1),span=Math.max(1,dayDiff(range.from,range.to)+1),daily=total.pay/elapsed,hDaily=total.hours/elapsed;
    var monthStart=new Date(today().getFullYear(),today().getMonth(),1),monthEnd=new Date(today().getFullYear(),today().getMonth()+1,0);
    var monthRange={from:monthStart,to:today(),fromISO:toISO(monthStart),toISO:toISO(today())};
    var monthTotal=totalsOf(analyticsEntries(monthRange)),monthElapsed=today().getDate(),monthDays=monthEnd.getDate();
    $("anProjectedMonth").textContent=money(monthTotal.pay/Math.max(1,monthElapsed)*monthDays);
    $("anProjectedPayday").textContent=np?money(np.pay+daily*Math.max(0,dayDiff(today(),np.payday))):money(0);
    $("anProjectedHours").textContent=dec(range.to<=today()?total.hours:hDaily*span);
  }

  /* ---------- take-home pay ----------
     Presets are intentionally limited to rules we can keep explainable and
     source from the relevant tax authority. The custom option covers any
     country, local tax, payroll agreement or confirmed exemption not listed. */
  var TAKE_HOME_PRESETS = {
    MT: {
      label: "Malta", currency: "EUR", year: "2026 rules",
      profiles: [
        ["mt-single", "Single"], ["mt-married", "Married"],
        ["mt-married1", "Married, one child"], ["mt-married2", "Married, two or more children"],
        ["mt-parent", "Parent rates"], ["mt-parent1", "Parent, one child"],
        ["mt-parent2", "Parent, two or more children"],
        ["mt-student-under18", "Student employee, under 18"],
        ["mt-student-18", "Student employee, 18 or over"],
        ["mt-student", "Eligible student, 10% part-time scheme"],
        ["mt-stipend", "Not employed: confirmed tax-exempt stipend or grant"]
      ],
      socials: [
        ["mt-prorata", "Employee, 10% pro-rata"],
        ["mt-standard", "Employee age 18 or over"],
        ["mt-under18", "Employee under 18"],
        ["mt-pre1962", "Employee born before 1962"],
        ["mt-apprentice18", "Student-worker, 18+ (max €7.94)"],
        ["mt-apprentice-under18", "Student-worker, under 18 (max €4.38)"],
        ["none", "No contribution (confirmed exempt)"]
      ],
      sources: [
        ["2026 Malta tax rates", "https://mtca.gov.mt/docs/default-source/documents/2026-tax-rates.pdf?sfvrsn=37563fb2_7"],
        ["Student and part-time tax", "https://mtca.gov.mt/personal-tax/individual/faqs/faqs1"],
        ["2026 social security", "https://socialsecurity.gov.mt/en/information-and-applications-for-benefits-and-services/social-security-contributions/social-security-contributions-class-1-2026/"]
      ]
    },
    UK: {
      label: "United Kingdom", currency: "GBP", year: "2026/27 rules",
      profiles: [["uk-standard", "Employee in England, Wales or Northern Ireland"]],
      sources: [["Income Tax rates", "https://www.gov.uk/government/publications/rates-and-allowances-income-tax/income-tax-rates-and-allowances-current-and-past"], ["National Insurance", "https://www.gov.uk/national-insurance/how-much-you-pay"]]
    },
    "UK-SCT": {
      label: "Scotland", currency: "GBP", year: "2026/27 rules",
      profiles: [["uk-scotland", "Employee in Scotland"]],
      sources: [["Scottish Income Tax rates", "https://www.gov.uk/government/publications/rates-and-allowances-income-tax/income-tax-rates-and-allowances-current-and-past"], ["National Insurance", "https://www.gov.uk/national-insurance/how-much-you-pay"]]
    },
    IE: {
      label: "Ireland", currency: "EUR", year: "2026 rules",
      profiles: [["ie-single", "Single"], ["ie-single-parent", "Single parent with child carer credit"], ["ie-married-one", "Married, one income"]],
      sources: [["2026 tax bands and credits", "https://www.revenue.ie/en/personal-tax-credits-reliefs-and-exemptions/tax-relief-charts/index.aspx"], ["2026 USC rates", "https://www.revenue.ie/en/corporate/press-office/budget-information/previous-years/2026/budget-summary.pdf"], ["Class A PRSI", "https://www.gov.ie/en/department-of-social-protection/publications/prsi-class-a-rates/"]]
    },
    NZ: {
      label: "New Zealand", currency: "NZD", year: "2026/27 rules",
      profiles: [["nz-standard", "Employee, main income"], ["nz-ietc", "Main income, eligible for IETC"]],
      sources: [["Income tax rates", "https://www.ird.govt.nz/income-tax/income-tax-for-individuals/tax-codes-and-tax-rates-for-individuals/tax-rates-for-individuals"], ["ACC earners' levy", "https://www.ird.govt.nz/en/income-tax/income-tax-for-individuals/acc-clients-and-carers/acc-earners-levy-rates"]]
    },
    CUSTOM: {
      label: "Custom", currency: "", year: "Custom rates",
      profiles: [["custom", "My own rates"]], sources: []
    }
  };

  function clampRate(n) { return Math.max(0, Math.min(100, Number(n) || 0)); }
  function progressiveTax(income, bands) {
    income = Math.max(0, Number(income) || 0);
    var tax = 0, lower = 0;
    bands.forEach(function (band) {
      var upper = band[0], taxable = Math.max(0, Math.min(income, upper) - lower);
      tax += taxable * band[1]; lower = upper;
    });
    return Math.max(0, tax);
  }
  function maltaTableTax(income, profile) {
    var tables = {
      "mt-single": [[12000,0,0],[16000,.15,1800],[60000,.25,3400],[Infinity,.35,9400]],
      "mt-married": [[15000,0,0],[23000,.15,2250],[60000,.25,4550],[Infinity,.35,10550]],
      "mt-married1": [[17500,0,0],[26500,.15,2625],[60000,.25,5275],[Infinity,.35,11275]],
      "mt-married2": [[22500,0,0],[32000,.15,3375],[60000,.25,6575],[Infinity,.35,12575]],
      "mt-parent": [[13000,0,0],[17500,.15,1950],[60000,.25,3700],[Infinity,.35,9700]],
      "mt-parent1": [[14500,0,0],[21000,.15,2175],[60000,.25,4275],[Infinity,.35,10275]],
      "mt-parent2": [[18500,0,0],[25500,.15,2775],[60000,.25,5325],[Infinity,.35,11325]]
    };
    var rows = tables[profile] || tables["mt-single"];
    for (var i = 0; i < rows.length; i++) if (income <= rows[i][0]) return Math.max(0, income * rows[i][1] - rows[i][2]);
    return 0;
  }
  function maltaIncomeTax(income, profile) {
    if (profile === "mt-stipend") return 0;
    if (profile === "mt-student") {
      return Math.min(income, 10000) * .10 + maltaTableTax(Math.max(0, income - 10000), "mt-single");
    }
    return maltaTableTax(income, profile);
  }
  function maltaSocial(income, social, profile) {
    if (profile === "mt-stipend" || social === "none") return 0;
    var weekly = Math.max(0, income) / 52, amount;
    if (social === "mt-apprentice18") amount = Math.min(weekly * .10, 7.94);
    else if (social === "mt-apprentice-under18") amount = Math.min(weekly * .10, 4.38);
    else if (social === "mt-prorata") amount = Math.min(weekly * .10, 55.93);
    else if (social === "mt-under18" && weekly <= 229.44) amount = 6.62;
    else if (social === "mt-standard" && weekly <= 229.44) amount = 22.94;
    else if (social === "mt-pre1962") amount = Math.min(weekly * .10, 49.04);
    else amount = Math.min(weekly * .10, 55.93);
    return Math.min(income, Math.max(0, amount * 52));
  }
  function ukAllowance(income) { return Math.max(0, 12570 - Math.max(0, income - 100000) / 2); }
  function ukTax(income, scotland) {
    var taxable = Math.max(0, income - ukAllowance(income));
    return scotland
      ? progressiveTax(taxable, [[3967,.19],[16956,.20],[31092,.21],[62430,.42],[125140,.45],[Infinity,.48]])
      : progressiveTax(taxable, [[37700,.20],[125140,.40],[Infinity,.45]]);
  }
  function ukNationalInsurance(income) {
    return progressiveTax(income, [[12570,0],[50270,.08],[Infinity,.02]]);
  }
  function irelandIncomeTax(income, profile) {
    var band = profile === "ie-single-parent" ? 48000 : profile === "ie-married-one" ? 53000 : 44000;
    var credits = profile === "ie-married-one" ? 6000 : profile === "ie-single-parent" ? 5900 : 4000;
    return Math.max(0, progressiveTax(income, [[band,.20],[Infinity,.40]]) - credits);
  }
  function irelandUsc(income) {
    if (income <= 13000) return 0;
    return progressiveTax(income, [[12012,.005],[28700,.02],[70044,.03],[Infinity,.08]]);
  }
  function irelandPrsi(income) {
    var weekly = income / 52;
    if (weekly <= 352) return 0;
    var rate = .042375; /* weighted 2026 average: 4.20% Jan-Sep, 4.35% Oct-Dec */
    var credit = weekly < 424 ? Math.max(0, 12 - (weekly - 352) / 6) : 0;
    return Math.max(0, weekly * rate - credit) * 52;
  }
  function nzTax(income) {
    return progressiveTax(income, [[15600,.105],[53500,.175],[78100,.30],[180000,.33],[Infinity,.39]]);
  }
  function nzIetc(income) {
    if (income < 24000 || income > 70000) return 0;
    return income <= 66000 ? 520 : Math.max(0, 520 - (income - 66000) * .13);
  }
  function taxAttributable(fn, workIncome, otherIncome) {
    return Math.max(0, fn(workIncome + otherIncome) - fn(otherIncome));
  }
  function expectedPaydayPay() {
    var now = today(), cycle = Math.max(1, parseInt(state.settings.cycleDays,10) || 28);
    var period = nextPeriod();
    /* firstPayday is normally always valid, but old or hand-edited backups may
       not have one. Keep the page useful while Settings asks for a real date. */
    if (!period) period = { start: now, end: addDays(now, cycle - 1), payday: addDays(now, cycle - 1) };
    var start = period.start, end = period.end, observedEnd = now < end ? now : end;
    var fullEntries = inRange(toISO(start), toISO(end));
    var currentEntries = observedEnd >= start ? inRange(toISO(start), toISO(observedEnd)) : [];
    var sourceEntries = currentEntries;
    var full = totalsOf(fullEntries);
    var toDate = totalsOf(currentEntries);
    var workDays = state.settings.logReminderDays || [1,2,3,4,5];
    var elapsedWorkDays = 0, periodWorkDays = 0;
    for (var d = new Date(start.getTime()); d <= end; d = addDays(d,1)) {
      if (workDays.indexOf(d.getDay()) !== -1) {
        periodWorkDays++;
        if (d <= observedEnd) elapsedWorkDays++;
      }
    }

    /* Use the worker's real average rate when this period has usable hours.
       Otherwise their default job is the most honest rate available. */
    var defaultEmployer = employerFor({ employerId: state.settings.takeHome.estimateEmployerId || state.settings.defaultEmployerId });
    var defaultRate = Math.max(0, Number(defaultEmployer.rate) || 0);
    var rateSource = toDate.hours > 0 ? toDate : full;
    var automaticRate = rateSource.hours > 0 ? rateSource.pay / rateSource.hours : defaultRate;
    var automaticHours = toDate.hours > 0
      ? toDate.hours / Math.max(1, elapsedWorkDays) * Math.max(1, periodWorkDays) : 0;
    if (!automaticHours) {
      var lookbackCycles = 3;
      var priorEntries = inRange(toISO(addDays(start, -cycle * lookbackCycles)), toISO(addDays(start, -1)));
      var prior = totalsOf(priorEntries);
      sourceEntries = priorEntries;
      automaticHours = prior.hours > 0 ? prior.hours / lookbackCycles : full.hours;
      if (prior.hours > 0 && prior.pay > 0) automaticRate = prior.pay / prior.hours;
    }
    var enteredHours = Math.max(0, Number(state.settings.takeHome.estimatedHours) || 0);
    var expectedHours = enteredHours > 0 ? enteredHours : automaticHours;
    var hourlyRate = enteredHours > 0 ? defaultRate : automaticRate;
    expectedHours = Math.max(full.hours, expectedHours);
    var loggedWeeks={};sourceEntries.forEach(function(e){if(hoursOf(e)>0)loggedWeeks[toISO(weekStart(parseDate(e.date)))]=true;});
    return {
      logged: full.pay, loggedHours: full.hours,
      projected: xround(Math.max(full.pay, expectedHours * hourlyRate), 2),
      projectedHours: xround(expectedHours, 2), hourlyRate: hourlyRate,
      customHours: enteredHours > 0, enteredHours: enteredHours, period: period,
      sourceWeeks: Object.keys(loggedWeeks).length,
      label: fmtDateShort(start) + " – " + fmtDateLong(end)
    };
  }
  function takeHomeEstimate(grossAnnual, cfg, workMonths) {
    workMonths = Math.max(1, Math.min(12, Number(workMonths) || 12));
    var activeAnnualRate = grossAnnual * 12 / workMonths;
    var other = Math.max(0, Number(cfg.otherIncome) || 0), incomeTax = 0, social = 0, levy = 0, fixed = 0;
    if (cfg.country === "MT") {
      incomeTax = taxAttributable(function(x){return maltaIncomeTax(x,cfg.profile);},grossAnnual,other);
      social = maltaSocial(activeAnnualRate,cfg.social,cfg.profile) * workMonths / 12;
    } else if (cfg.country === "UK" || cfg.country === "UK-SCT") {
      incomeTax = taxAttributable(function(x){return ukTax(x,cfg.country === "UK-SCT");},grossAnnual,other);
      social = ukNationalInsurance(activeAnnualRate) * workMonths / 12;
    } else if (cfg.country === "IE") {
      incomeTax = taxAttributable(function(x){return irelandIncomeTax(x,cfg.profile);},grossAnnual,other);
      levy = taxAttributable(irelandUsc,grossAnnual,other);
      social = irelandPrsi(activeAnnualRate) * workMonths / 12;
    } else if (cfg.country === "NZ") {
      incomeTax = taxAttributable(nzTax,grossAnnual,other);
      if (cfg.profile === "nz-ietc") incomeTax = Math.max(0, incomeTax - Math.max(0, nzIetc(grossAnnual + other) - nzIetc(other)));
      social = Math.min(grossAnnual,156641) * .0175;
    } else {
      incomeTax = grossAnnual * clampRate(cfg.customTax) / 100;
      social = grossAnnual * clampRate(cfg.customSocial) / 100;
      levy = grossAnnual * clampRate(cfg.customOther) / 100;
      fixed = Math.max(0, Number(cfg.customFixed) || 0) * workMonths;
    }
    var deductions = Math.min(grossAnnual, Math.max(0, incomeTax + social + levy + fixed));
    return { incomeTax: incomeTax, social: social, levy: levy, fixed: fixed,
      deductions: deductions, net: Math.max(0, grossAnnual - deductions) };
  }
  function profileLabel(preset, value) {
    var found = (preset.profiles || []).filter(function(x){return x[0] === value;})[0];
    return found ? found[1] : preset.label;
  }
  function takeHomeHelp(cfg) {
    if (cfg.country === "MT" && cfg.profile === "mt-student-under18") return "For employment while studying and under 18. Income tax uses the normal single bands. The student-worker or apprenticeship social-security category is 10% of weekly pay, capped at €4.38 a week; an ordinary under-18 employee can use a different category.";
    if (cfg.country === "MT" && cfg.profile === "mt-student-18") return "For employment while studying and age 18 or over. Income tax uses the normal single bands. Choose the social-security category shown by your employer or payslip.";
    if (cfg.country === "MT" && cfg.profile === "mt-student") return "For eligible full-time students or apprentices using Malta's 10% part-time-work scheme on the first €10,000. This is not a general student exemption.";
    if (cfg.country === "MT" && cfg.profile === "mt-stipend") return "Use this only when you are not employed and the stipend or grant is confirmed tax-exempt. Do not use it for income shown on an employer's payslip.";
    if (cfg.country === "MT") return "Uses Malta's 2026 resident employee rates. The social-security choice matters for low-paid, part-time and apprenticeship work.";
    if (cfg.country === "UK" || cfg.country === "UK-SCT") return "Uses employee Income Tax and Class 1 National Insurance for 6 April 2026 to 5 April 2027.";
    if (cfg.country === "IE") return "Uses 2026 PAYE bands and standard credits, USC and an annual estimate of Class A employee PRSI.";
    if (cfg.country === "NZ") return "Uses the current individual rates and the 2026/27 ACC earners' levy." + (cfg.profile === "nz-ietc" ? " Includes the Independent Earner Tax Credit, assuming you meet its eligibility rules." : " KiwiSaver, student loans and the Independent Earner Tax Credit are not included in this option.");
    return "Enter the percentages normally taken from your gross pay. Use the fixed amount for a monthly pension, union fee or another regular deduction.";
  }
  function renderTakeHomeControls() {
    var cfg = state.settings.takeHome, preset = TAKE_HOME_PRESETS[cfg.country] || TAKE_HOME_PRESETS.CUSTOM;
    $("thCountry").value = cfg.country;
    $("thProfile").innerHTML = preset.profiles.map(function(x){return '<option value="'+x[0]+'">'+esc(x[1])+'</option>';}).join("");
    if (!preset.profiles.some(function(x){return x[0] === cfg.profile;})) cfg.profile = preset.profiles[0][0];
    $("thProfile").value = cfg.profile;
    var isMalta = cfg.country === "MT", custom = cfg.country === "CUSTOM";
    $("thSocialWrap").classList.toggle("hidden", !isMalta || cfg.profile === "mt-stipend");
    if (isMalta) {
      $("thSocial").innerHTML = preset.socials.map(function(x){return '<option value="'+x[0]+'">'+esc(x[1])+'</option>';}).join("");
      if (!preset.socials.some(function(x){return x[0] === cfg.social;})) cfg.social = "mt-prorata";
      $("thSocial").value = cfg.social;
    }
    $("thCustomFields").classList.toggle("hidden", !custom);
    $("thOtherIncome").value = Number(cfg.otherIncome) || 0;
    $("thCustomTax").value = Number(cfg.customTax) || 0;
    $("thCustomSocial").value = Number(cfg.customSocial) || 0;
    $("thCustomOther").value = Number(cfg.customOther) || 0;
    $("thCustomFixed").value = Number(cfg.customFixed) || 0;
    $("thWorkMonths").value = Math.max(1, Math.min(12, Number(cfg.workMonths) || 12));
    $("thEstimatedHours").value = Number(cfg.estimatedHours) > 0 ? Number(cfg.estimatedHours) : "";
    $("thEstimateEmployer").innerHTML = (state.settings.employers || []).map(function(e){
      return '<option value="'+esc(e.id)+'">'+esc(e.name)+'</option>';
    }).join("");
    $("thEstimateEmployer").value = cfg.estimateEmployerId || state.settings.defaultEmployerId;
    $("thProfileHelp").textContent = takeHomeHelp(cfg);
    var payslip=cfg.payslip||{};
    [["thPayslipGross","gross"],["thPayslipTax","tax"],["thPayslipSocial","social"],["thPayslipOther","other"],["thPayslipNet","net"]].forEach(function(pair){
      $(pair[0]).value=Number(payslip[pair[1]])>0?Number(payslip[pair[1]]):"";
    });
  }
  function currentTakeHomeModel() {
    var cfg = state.settings.takeHome;
    var preset = TAKE_HOME_PRESETS[cfg.country] || TAKE_HOME_PRESETS.CUSTOM;
    var pay = expectedPaydayPay();
    var workMonths = Math.max(1, Math.min(12, Number(cfg.workMonths) || 12));
    var cycle = Math.max(1, parseInt(state.settings.cycleDays,10) || 28);
    var periodsPerYear = 365.2425 / cycle;
    var workPeriods = Math.max(1, periodsPerYear * workMonths / 12);
    var grossPeriod = pay.projected;
    var grossYear = grossPeriod * workPeriods;
    var result = takeHomeEstimate(grossYear, cfg, workMonths);
    return { cfg: cfg, preset: preset, pay: pay, workMonths: workMonths,
      workPeriods: workPeriods, periodsPerYear: periodsPerYear,
      grossPeriod: grossPeriod, grossYear: grossYear, result: result,
      netPeriod: result.net / workPeriods,
      deductPeriod: result.deductions / workPeriods,
      /* Keep these aliases for spending plans and older widget code. Their
         budget now follows the upcoming payday rather than a calendar month. */
      grossMonth: grossPeriod, netMonth: result.net / workPeriods,
      deductMonth: result.deductions / workPeriods };
  }
  function renderTakeHome() {
    renderTakeHomeControls();
    var model = currentTakeHomeModel(), cfg = model.cfg, preset = model.preset;
    var pay = model.pay, grossPeriod = model.grossPeriod, workMonths = model.workMonths;
    var grossYear = model.grossYear, result = model.result;
    var netPeriod = model.netPeriod, deductPeriod = model.deductPeriod, workPeriods = model.workPeriods;
    $("thRulesYear").textContent = preset.year; $("thPresetName").textContent = profileLabel(preset,cfg.profile);
    $("thCurrencyMatch").textContent = preset.currency || (state.settings.currency === "custom" ? "Custom" : state.settings.currency);
    $("thLogged").textContent = money(pay.logged); $("thGross").textContent = money(grossPeriod);
    $("thGrossNote").textContent = "expected by " + fmtDateLong(pay.period.payday);
    $("thNet").textContent = money(netPeriod); $("thDeductions").textContent = money(deductPeriod); $("thYearNet").textContent = money(result.net);
    $("thYearNetNote").textContent = "across " + workMonths + (workMonths === 1 ? " work month" : " work months");
    var effective = grossYear ? result.deductions / grossYear * 100 : 0;
    $("thRate").textContent = effective.toFixed(1) + "% of gross";
    $("thNetBar").style.width = (grossYear ? Math.max(0,Math.min(100,result.net/grossYear*100)) : 0) + "%";
    $("thNetCaption").textContent = grossPeriod ? "From " + money(grossPeriod) + " expected gross pay on " + fmtDateLong(pay.period.payday) + "." : "Log some work or enter your expected hours to create an estimate.";
    $("thConfidence").textContent = pay.customHours ? "Using your hours" : pay.sourceWeeks
      ? "Based on " + pay.sourceWeeks + (pay.sourceWeeks===1?" logged week":" logged weeks") : "Waiting for work hours";
    $("thHoursHelp").textContent = pay.customHours && pay.enteredHours < pay.loggedHours
      ? "You entered " + dec(pay.enteredHours) + " hours, but " + dec(pay.loggedHours) + " are already logged for " + pay.label + ", so the higher total is kept."
      : pay.customHours
        ? "Using " + dec(pay.projectedHours) + " hours at " + money(pay.hourlyRate) + " an hour for " + pay.label + ". Clear the field to estimate from your shifts."
        : "Estimated from your shifts for " + pay.label + ". Enter a total if you already know the hours you expect to work.";
    var socialLabel = (cfg.country === "UK" || cfg.country === "UK-SCT") ? "National Insurance" : cfg.country === "IE" ? "PRSI" : cfg.country === "NZ" ? "ACC earners' levy" : cfg.social === "mt-apprentice-under18" ? "Social security (student under 18)" : "Social security";
    var rows = [["Expected gross pay",grossYear,"gross"],["Income tax",-result.incomeTax,"deduction"],[socialLabel,-result.social,"deduction"]];
    if (result.levy || cfg.country === "IE" || cfg.country === "CUSTOM") rows.push([cfg.country === "IE" ? "Universal Social Charge" : "Other deductions",-result.levy,"deduction"]);
    if (result.fixed || cfg.country === "CUSTOM") rows.push(["Fixed monthly deductions",-result.fixed,"deduction"]);
    rows.push(["Estimated take-home",result.net,"net"]);
    $("thBreakdown").innerHTML = rows.map(function(r){return '<tr class="tax-row-'+r[2]+'"><td>'+esc(r[0])+'</td><td>'+money(r[1]/workPeriods)+'</td><td>'+money(r[1])+'</td></tr>';}).join("");
    var mismatch = preset.currency && state.settings.currency !== preset.currency;
    $("thCurrencyWarning").classList.toggle("hidden", !mismatch);
    $("thCurrencyWarning").textContent = mismatch ? "Your app uses " + state.settings.currency + ", but this preset is for " + preset.currency + ". Change the app currency or choose a matching country before relying on this estimate." : "";
    $("thStudentNote").classList.toggle("hidden", cfg.country !== "MT");
    $("thAssumptions").textContent = takeHomeHelp(cfg) + " The estimate follows your " + state.settings.cycleDays + "-day pay cycle ending " + fmtDateLong(pay.period.payday) + ". Annual tax is estimated from that pay and " + workMonths + (workMonths === 1 ? " working month" : " working months") + ", then shared across about " + workPeriods.toFixed(1) + " paydays.";
    var ps=cfg.payslip||{},actualNet=Number(ps.net)||Math.max(0,(Number(ps.gross)||0)-(Number(ps.tax)||0)-(Number(ps.social)||0)-(Number(ps.other)||0));
    if (Number(ps.gross)>0) {
      var difference=xround(actualNet-netPeriod,2),close=Math.abs(difference)<0.02;
      $("thPayslipPill").textContent=close?"matches":difference>0?"higher":"lower";
      $("thPayslipResult").textContent="This payslip paid "+money(actualNet)+" net. That is "+money(Math.abs(difference))+" "+(close?"from the estimate.":difference>0?"more than the estimate.":"less than the estimate.");
    } else {
      $("thPayslipPill").textContent="optional";
      $("thPayslipResult").textContent="Enter the figures printed on your payslip.";
    }
    $("thSources").innerHTML = (preset.sources || []).map(function(s){return '<button class="btn tax-source" type="button" data-url="'+esc(s[1])+'">'+esc(s[0])+' ↗</button>';}).join("");
    [].forEach.call($("thSources").querySelectorAll(".tax-source"),function(button){button.onclick=function(){var url=button.dataset.url;if(window.WPLBridge&&window.WPLBridge.openExternal)window.WPLBridge.openExternal(url);else if(window.WPLDesktop&&window.WPLDesktop.openExternal)window.WPLDesktop.openExternal(url);else window.open(url,"_blank");};});
  }

  /* ---------- spending plan ---------- */
  function spendingModel() {
    var plan = state.settings.spendingPlan;
    var takeHome = currentTakeHomeModel();
    var budget = plan.incomeMode === "custom" ? Math.max(0, Number(plan.customIncome) || 0) : takeHome.netMonth;
    var period = nextPeriod(), fromISO = period ? toISO(period.start) : toISO(new Date(today().getFullYear(),today().getMonth(),1));
    var toISOd = period ? toISO(period.end) : toISO(new Date(today().getFullYear(),today().getMonth()+1,0));
    var transactions = (plan.transactions || []).filter(function (transaction) {
      return transaction.date >= fromISO && transaction.date <= toISOd;
    });
    var categories = plan.categories.map(function (category) {
      var amount = category.enabled === false ? 0 : category.mode === "percent" ? budget * Math.max(0, Number(category.value) || 0) / 100 : Math.max(0, Number(category.value) || 0);
      var spent = transactions.filter(function (transaction) { return transaction.categoryId === category.id; })
        .reduce(function (sum, transaction) { return sum + transaction.amount; }, 0);
      return { id: category.id, name: category.name, color: category.color,
        mode: category.mode, value: category.value, enabled: category.enabled !== false,
        amount: xround(amount,2), spent: xround(spent,2) };
    });
    var allocated = categories.reduce(function (sum, category) { return sum + category.amount; }, 0);
    var spent = transactions.reduce(function (sum, transaction) { return sum + transaction.amount; }, 0);
    var dueBills=(plan.bills||[]).filter(function(bill){return bill.enabled!==false&&bill.nextDue<=toISOd;});
    var reservedBills=dueBills.reduce(function(sum,bill){return sum+bill.amount;},0);
    return { plan: plan, takeHome: takeHome, budget: budget, categories: categories,
      transactions:transactions,spent:xround(spent,2),dueBills:dueBills,reservedBills:xround(reservedBills,2),safeToSpend:xround(Math.max(0,budget-spent-reservedBills),2),
      fromISO: fromISO, toISO: toISOd, allocated: allocated, remaining: budget - allocated };
  }

  /* ---------- Android home-screen widgets ----------
     The launcher receives calculated summaries rather than the full work log.
     A short range is included so midnight, Monday and payday rollovers remain
     correct even when the app has not been opened that morning. */
  var widgetRefreshTimer = null;
  function scheduleWidgetRefresh() {
    if (!window.WPLBridge || !window.WPLBridge.updateWidgetData) return;
    clearTimeout(widgetRefreshTimer);
    widgetRefreshTimer = setTimeout(pushWidgetSnapshot, 120);
  }

  function widgetModelFor(employerId, baseDate, netRatio, saving) {
    var list = entries().filter(function (entry) {
      return employerId === "all" || entry.employerId === employerId;
    });
    var days = {};
    for (var offset = -1; offset <= 14; offset++) {
      var date = addDays(baseDate, offset), iso = toISO(date);
      var dayEntries = list.filter(function (entry) { return entry.date === iso; });
      var total = totalsOf(dayEntries), first = dayEntries[0] || null;
      days[iso] = {
        logged: !!first,
        start: first && first.start || "",
        finish: first && first.finish || "",
        breakHours: first ? Number(employerFor(first).breakHours || 0) : 0,
        hours: total.hours,
        pay: total.pay,
        label: !first ? "Not logged yet" : isLeave(first) ? leaveInfo(first.leave).label
          : total.hours > 0 ? first.start + " – " + first.finish : "Day off"
      };
    }

    var weekModels = [];
    var firstWeek = weekStart(baseDate);
    for (var weekOffset = 0; weekOffset < 9; weekOffset++) {
      var start = addDays(firstWeek, weekOffset * 7), end = addDays(start, 6);
      var totalWeek = totalsOf(list.filter(function (entry) {
        return entry.date >= toISO(start) && entry.date <= toISO(end);
      }));
      weekModels.push({ start: toISO(start), end: toISO(end), hours: totalWeek.hours,
        pay: totalWeek.pay, days: totalWeek.days, target: fullWeek() });
    }

    var paydayModels = periods().filter(function (period) {
      return period.payday >= baseDate;
    }).slice(0, 8).map(function (period) {
      var periodTotal = totalsOf(list.filter(function (entry) {
        return entry.date >= toISO(period.start) && entry.date <= toISO(period.end);
      }));
      return { date: toISO(period.payday), start: toISO(period.start), end: toISO(period.end),
        gross: periodTotal.pay, net: xround(periodTotal.pay * netRatio, 2),
        hours: periodTotal.hours, days: periodTotal.days };
    });

    return { days: days, weeks: weekModels, paydays: paydayModels,
      savingLabel: saving ? saving.name : "", savingAmount: saving ? saving.amount : 0 };
  }

  function widgetSnapshot() {
    var baseDate = today(), takeHome = currentTakeHomeModel(), spending = spendingModel();
    var netRatio = takeHome.grossMonth > 0
      ? Math.max(0, Math.min(1, takeHome.netMonth / takeHome.grossMonth)) : 1;
    var saving = spending.categories.filter(function (category) {
      return category.id === "savings" || /sav|hold/i.test(category.name || "");
    })[0] || null;
    var selected = APP_COLORS[state.settings.appColor] ? state.settings.appColor : "orange";
    var dark = document.documentElement.classList.contains("dark");
    var colours = dark ? APP_COLORS[selected].dark : APP_COLORS[selected].light;
    var employers = (state.settings.employers || []).map(function (employer) {
      return { id: employer.id, name: employer.name || "My job" };
    });
    var models = { all: widgetModelFor("all", baseDate, netRatio, saving) };
    employers.forEach(function (employer) {
      models[employer.id] = widgetModelFor(employer.id, baseDate, netRatio, saving);
    });
    return { version: 1, createdAt: Date.now(), currency: currencySymbol(), dark: dark,
      accent: colours[0], accentSoft: colours[2], employers: employers, models: models };
  }

  function pushWidgetSnapshot() {
    widgetRefreshTimer = null;
    if (!window.WPLBridge || !window.WPLBridge.updateWidgetData) return;
    try { window.WPLBridge.updateWidgetData(JSON.stringify(widgetSnapshot())); }
    catch (err) { /* the next save or resume will retry */ }
  }
  function spendingGradient(model) {
    var total = Math.max(model.budget, model.allocated);
    if (!total) return "conic-gradient(var(--line) 0 100%)";
    var cursor = 0, parts = [];
    model.categories.forEach(function (category) {
      if (category.amount <= 0) return;
      var end = Math.min(100, cursor + category.amount / total * 100);
      parts.push(category.color + " " + cursor.toFixed(3) + "% " + end.toFixed(3) + "%");
      cursor = end;
    });
    if (cursor < 100) parts.push("var(--line) " + cursor.toFixed(3) + "% 100%");
    return "conic-gradient(" + (parts.length ? parts.join(",") : "var(--line) 0 100%") + ")";
  }
  function renderSpending() {
    var model = spendingModel(), plan = model.plan;
    $("spIncomeMode").value = plan.incomeMode;
    $("spCustomIncome").value = Number(plan.customIncome) || 0;
    $("spCustomIncomeWrap").classList.toggle("hidden", plan.incomeMode !== "custom");
    $("spSourcePill").textContent = plan.incomeMode === "custom" ? "Your amount" : "Take-home estimate";
    $("spIncomeNote").innerHTML = plan.incomeMode === "custom"
      ? "This amount is only used for your spending plan."
      : (model.budget ? "Using " + money(model.budget) + " from your Take-home pay estimate."
        : "Your take-home estimate is empty. Log some work, check your tax setup, or choose an amount to enter.");
    $("spBudget").textContent = money(model.budget);
    $("spPlanned").textContent = money(model.allocated);
    var activeCount=model.categories.filter(function(category){return category.enabled;}).length;
    $("spPlannedNote").textContent = "across " + activeCount + (activeCount === 1 ? " category" : " categories");
    $("spSpent").textContent=money(model.spent);
    $("spSpentNote").textContent=money(model.safeToSpend)+" left before payday";
    var over = model.remaining < -0.005, balance = Math.abs(model.remaining);
    $("spBalanceLabel").textContent = over ? "Over budget" : "Left to plan";
    $("spBalance").textContent = money(balance);
    $("spBalanceNote").textContent = over ? "more than is available" : "unallocated";
    $("spBalanceCard").classList.toggle("is-over", over);
    $("spChartPill").textContent = activeCount + (activeCount === 1 ? " category" : " categories");
    $("spDonut").style.background = "radial-gradient(circle at center,var(--surface) 0 55%,transparent 56%)," + spendingGradient(model);
    $("spDonutValue").textContent = money(balance);
    $("spDonutLabel").textContent = over ? "over" : "left";
    var legend = model.categories.filter(function (category) { return category.amount > 0; }).map(function (category) {
      return '<div><span><i style="background:'+category.color+'"></i>'+esc(category.name)+'</span><strong>'+money(category.spent)+' / '+money(category.amount)+'</strong></div>';
    });
    if (model.remaining > 0.005) legend.push('<div><span><i class="sp-unallocated"></i>Not planned yet</span><strong>'+money(model.remaining)+'</strong></div>');
    $("spLegend").innerHTML = legend.join("") || '<p class="muted tiny">Add a category to see your split.</p>';
    var scale = Math.max(model.budget, model.allocated, 1);
    $("spBars").innerHTML = model.categories.map(function (category) {
      var width = Math.min(100, category.amount / scale * 100);
      var share = model.budget ? category.amount / model.budget * 100 : 0;
      return '<div class="sp-bar-row"><div class="sp-bar-head"><span>'+esc(category.name)+'</span><strong>'+money(category.amount)+'</strong></div><div class="sp-bar-track"><i style="width:'+width.toFixed(2)+'%;background:'+category.color+'"></i></div><small>'+share.toFixed(1)+'% of your budget</small></div>';
    }).join("") || '<p class="muted tiny">Your category bars will appear here.</p>';
    $("spWarning").classList.toggle("hidden", !over);
    $("spWarning").textContent = over ? "Your plan is " + money(balance) + " over budget. Lower a category or increase the money available." : "";
    $("spEmpty").classList.toggle("hidden", model.categories.length > 0);
    $("spTransactionRange").textContent=fmtDateShort(parseDate(model.fromISO))+" – "+fmtDateShort(parseDate(model.toISO));
    $("spTxCategory").innerHTML=model.categories.filter(function(category){return category.enabled;}).map(function(category){return '<option value="'+esc(category.id)+'">'+esc(category.name)+'</option>';}).join("");
    if (!$("spTxDate").value) $("spTxDate").value=toISO(today());
    $("spTransactionList").innerHTML=model.transactions.slice().sort(function(a,b){return b.date.localeCompare(a.date)||b.updatedAt-a.updatedAt;}).map(function(transaction){var category=model.categories.filter(function(item){return item.id===transaction.categoryId;})[0];return '<div class="transaction-row" data-tx-id="'+esc(transaction.id)+'"><span class="transaction-dot" style="background:'+(category?category.color:'var(--line)')+'"></span><span><strong>'+esc(transaction.note||(category?category.name:"Spending"))+'</strong><small>'+esc(fmtDateLong(parseDate(transaction.date)))+(category?' · '+esc(category.name):'')+'</small></span><b>'+money(transaction.amount)+'</b><button class="icon-btn tx-delete" type="button" aria-label="Remove spending">×</button></div>';}).join("")||'<p class="empty-inline">No spending added in this pay period.</p>';
    var categoryOptions=model.categories.map(function(category){return '<option value="'+esc(category.id)+'">'+esc(category.name)+'</option>';}).join("");
    $("spBillCategory").innerHTML=categoryOptions;$("spBillsTotal").textContent=money(model.reservedBills)+" due";
    if(!$("spBillDue").value)$("spBillDue").value=toISO(today());
    $("spBillList").innerHTML=(plan.bills||[]).map(function(bill){var due=bill.enabled!==false&&bill.nextDue<=model.toISO;return '<div class="money-plan-row" data-bill-id="'+esc(bill.id)+'"><span><strong>'+esc(bill.name)+'</strong><small>'+money(bill.amount)+' · '+(bill.enabled===false?'paused':(due?'due '+fmtDateShort(parseDate(bill.nextDue)):'next '+fmtDateShort(parseDate(bill.nextDue))))+'</small></span><button class="btn bill-paid" type="button"'+(!due?' disabled':'')+'>Paid</button><button class="icon-btn bill-toggle" type="button" aria-label="'+(bill.enabled===false?'Use':'Pause')+' '+esc(bill.name)+'">'+(bill.enabled===false?'○':'●')+'</button><button class="icon-btn bill-delete" type="button" aria-label="Remove '+esc(bill.name)+'">×</button></div>';}).join("")||'<p class="empty-inline">No recurring bills.</p>';
    $("spGoalList").innerHTML=(plan.goals||[]).map(function(goal){var pct=Math.min(100,goal.target?goal.saved/goal.target*100:0);return '<div class="goal-row" data-goal-id="'+esc(goal.id)+'"><div class="goal-head"><span><strong>'+esc(goal.name)+'</strong><small>'+money(goal.saved)+' of '+money(goal.target)+(goal.targetDate?' · by '+fmtDateShort(parseDate(goal.targetDate)):'')+'</small></span><button class="icon-btn goal-delete" type="button" aria-label="Remove '+esc(goal.name)+'">×</button></div><div class="sp-bar-track"><i style="width:'+pct.toFixed(2)+'%;background:var(--accent)"></i></div><label class="goal-saved"><span>Saved</span><input type="number" min="0" step="1" value="'+Number(goal.saved||0)+'" inputmode="decimal"></label></div>';}).join("")||'<p class="empty-inline">No savings goals.</p>';
    $("spCategoryList").innerHTML = model.categories.map(function (category, index) {
      var displayValue = category.mode === "percent" ? Math.round(Number(category.value) * 100) / 100 : Number(category.value);
      var canFill = model.remaining > 0.004 && model.budget > 0;
      var fillButton = '<button class="btn sp-fill-category" type="button" title="'+(canFill ? 'Add '+money(model.remaining)+' to '+esc(category.name) : 'There is no money left to plan')+'"'+(canFill ? '' : ' disabled')+'>Fill rest</button>';
      return '<div class="sp-category'+(category.enabled?'':' is-disabled')+'" data-sp-id="'+esc(category.id)+'">'+
        '<label class="sp-color-field"><span>Colour</span><input class="sp-color" type="color" value="'+category.color+'" aria-label="Colour for '+esc(category.name)+'"></label>'+
        '<label class="field sp-name-field"><span>Name</span><input class="sp-name" maxlength="40" value="'+esc(category.name)+'"></label>'+
        '<label class="field sp-mode-field"><span>Plan by</span><select class="sp-mode"><option value="percent"'+(category.mode === "percent" ? " selected" : "")+'>Percentage</option><option value="amount"'+(category.mode === "amount" ? " selected" : "")+'>Exact amount</option></select></label>'+
        '<label class="field sp-value-field"><span>'+(category.mode === "percent" ? "Percent" : "Amount")+'</span><div class="input-suffix"><input class="sp-value" type="number" min="0" step="'+(category.mode === "percent" ? ".01" : "1")+'" inputmode="decimal" value="'+displayValue+'"><b>'+(category.mode === "percent" ? "%" : esc(currencySymbol()))+'</b></div></label>'+
        '<div class="sp-category-total"><span>This plan</span><strong>'+money(category.amount)+'</strong></div>'+
        '<div class="sp-category-actions"><label class="sp-enabled"><input class="sp-enabled-input" type="checkbox"'+(category.enabled?' checked':'')+'><span>Use</span></label>'+fillButton+'<button class="icon-btn sp-up" type="button" title="Move up" aria-label="Move '+esc(category.name)+' up"'+(index === 0 ? " disabled" : "")+'>↑</button><button class="icon-btn sp-down" type="button" title="Move down" aria-label="Move '+esc(category.name)+' down"'+(index === model.categories.length - 1 ? " disabled" : "")+'>↓</button><button class="icon-btn sp-delete" type="button" title="Remove" aria-label="Remove '+esc(category.name)+'">×</button></div></div>';
    }).join("");
    bindSpendingRows();
  }
  function saveSpending() {
    state.settingsUpdatedAt = Date.now();
    save();
    renderSpending();
  }
  function bindSpendingRows() {
    [].forEach.call($("spCategoryList").querySelectorAll(".sp-category"), function (row) {
      var index = state.settings.spendingPlan.categories.findIndex(function (category) { return category.id === row.dataset.spId; });
      if (index < 0) return;
      var category = state.settings.spendingPlan.categories[index];
      row.querySelector(".sp-color").onchange = function () { category.color = this.value; saveSpending(); };
      row.querySelector(".sp-name").onchange = function () { category.name = this.value.trim().slice(0,40) || "New category"; saveSpending(); };
      row.querySelector(".sp-mode").onchange = function () { category.mode = this.value === "amount" ? "amount" : "percent"; saveSpending(); };
      row.querySelector(".sp-value").onchange = function () { category.value = Math.max(0, Number(this.value) || 0); saveSpending(); };
      row.querySelector(".sp-enabled-input").onchange=function(){category.enabled=this.checked;saveSpending();toast(this.checked?category.name+" is in your plan":category.name+" is paused");};
      var fill = row.querySelector(".sp-fill-category");
      if (fill) fill.onclick = function () {
        var model = spendingModel();
        if (model.remaining <= 0.004 || model.budget <= 0) return;
        var targetModel = model.categories.filter(function (item) { return item.id === category.id; })[0];
        var everythingElse = model.allocated - (targetModel ? targetModel.amount : 0);
        if (category.mode === "amount") {
          category.value = Math.round(Math.max(0, model.budget - everythingElse) * 100) / 100;
        } else {
          category.value = Math.round((Math.max(0, model.budget - everythingElse) / model.budget * 100) * 1000000000) / 1000000000;
        }
        saveSpending();
        toast("Remaining money added to " + category.name);
      };
      row.querySelector(".sp-up").onclick = function () { if (index < 1) return; var list = state.settings.spendingPlan.categories; list.splice(index - 1, 0, list.splice(index,1)[0]); saveSpending(); };
      row.querySelector(".sp-down").onclick = function () { var list = state.settings.spendingPlan.categories; if (index >= list.length - 1) return; list.splice(index + 1, 0, list.splice(index,1)[0]); saveSpending(); };
      row.querySelector(".sp-delete").onclick = function () {
        var removed = Object.assign({}, category), removedAt = index;
        state.settings.spendingPlan.categories.splice(index,1); saveSpending();
        toastUndo(removed.name + " removed", function () {
          state.settings.spendingPlan.categories.splice(Math.min(removedAt,state.settings.spendingPlan.categories.length),0,removed);
          saveSpending(); toast("Category restored");
        });
      };
    });
    [].forEach.call($("spTransactionList").querySelectorAll(".transaction-row"),function(row){row.querySelector(".tx-delete").onclick=function(){var list=state.settings.spendingPlan.transactions,index=list.findIndex(function(transaction){return transaction.id===row.dataset.txId;});if(index<0)return;var removed=list.splice(index,1)[0];saveSpending();toastUndo("Spending removed",function(){list.splice(Math.min(index,list.length),0,removed);saveSpending();});};});
    [].forEach.call($("spBillList").querySelectorAll(".money-plan-row"),function(row){var list=state.settings.spendingPlan.bills,index=list.findIndex(function(bill){return bill.id===row.dataset.billId;}),bill=list[index];if(!bill)return;row.querySelector(".bill-paid").onclick=function(){state.settings.spendingPlan.transactions.push({id:"transaction-"+uid(),date:toISO(today()),categoryId:bill.categoryId,amount:bill.amount,note:bill.name,updatedAt:Date.now()});var due=parseDate(bill.nextDue),lastNextMonth=new Date(due.getFullYear(),due.getMonth()+2,0).getDate();bill.nextDue=toISO(new Date(due.getFullYear(),due.getMonth()+1,Math.min(due.getDate(),lastNextMonth)));saveSpending();toast(bill.name+" marked paid");};row.querySelector(".bill-toggle").onclick=function(){bill.enabled=bill.enabled===false;saveSpending();};row.querySelector(".bill-delete").onclick=function(){var removed=list.splice(index,1)[0];saveSpending();toastUndo("Bill removed",function(){list.splice(Math.min(index,list.length),0,removed);saveSpending();});};});
    [].forEach.call($("spGoalList").querySelectorAll(".goal-row"),function(row){var list=state.settings.spendingPlan.goals,index=list.findIndex(function(goal){return goal.id===row.dataset.goalId;}),goal=list[index];if(!goal)return;row.querySelector(".goal-saved input").onchange=function(){goal.saved=Math.max(0,Number(this.value)||0);saveSpending();};row.querySelector(".goal-delete").onclick=function(){var removed=list.splice(index,1)[0];saveSpending();toastUndo("Goal removed",function(){list.splice(Math.min(index,list.length),0,removed);saveSpending();});};});
  }

  /* ---------- weekly ---------- */
  function renderWeekly() {
    var ws = weeks(), curStart = toISO(weekStart(today()));
    var th = 0, tp = 0;
    $("wkList").innerHTML = ws.map(function (w) {
      th += w.hours; tp += w.pay;
      var isCur = toISO(w.start) === curStart;
      var pct = Math.min(100, (w.hours / fullWeek()) * 100);
      return '<div class="wcard' + (isCur ? " is-current" : "") + (w.hours === 0 ? " is-empty" : "") + '">' +
        '<div class="wcard-top"><div><div class="wcard-title">' +
        fmtDateShort(w.start) + " – " + fmtDateShort(w.end) + '</div>' +
        '<div class="wcard-sub">' + fmtDate(w.start) + " to " + fmtDate(w.end) + '</div></div>' +
        (isCur ? '<span class="pill pill-accent">This week</span>'
               : '<span class="pill">' + w.days + (w.days === 1 ? " day" : " days") + '</span>') +
        '</div>' +
        '<div class="bar"><div class="bar-fill" style="width:' + pct + '%"></div></div>' +
        '<div class="wcard-grid">' +
        '<div><span class="k">Hours</span><span class="v">' + dec(w.hours) + '</span></div>' +
        '<div><span class="k">Time worked</span><span class="v">' + hm(w.hours) + '</span></div>' +
        '<div><span class="k">Pay</span><span class="v accent">' + money(w.pay) + '</span></div>' +
        '</div></div>';
    }).join("") || '<p class="muted tiny">Nothing logged yet.</p>';

    $("wkTotalHours").textContent = dec(xround(th, 4));
    $("wkTotalTime").textContent = hmPlain(th);
    $("wkTotalPay").textContent = money(xround(tp, 2));
  }

  /* ---------- pay periods ---------- */
  function renderPeriods() {
    var ps = periods(), np = nextPeriod(ps);
    var th = 0, tp = 0;
    $("ppNote").textContent = "Each pay period is the " + state.settings.cycleDays +
      " days ending on the payday itself. First payday " +
      fmtDateLong(parseDate(state.settings.firstPayday)) + ".";

    $("ppList").innerHTML = ps.map(function (p) {
      th += p.hours; tp += p.pay;
      var isNext = np && +p.payday === +np.payday;
      var past = p.payday < today();
      var d = dayDiff(today(), p.payday);
      return '<div class="wcard' + (isNext ? " is-current" : "") + (p.hours === 0 ? " is-empty" : "") + '">' +
        '<div class="wcard-top"><div><div class="wcard-title">Payday ' + fmtDateLong(p.payday) + '</div>' +
        '<div class="wcard-sub">' + DAYS[p.payday.getDay()] + " · period " + fmtDate(p.start) + " to " + fmtDate(p.end) + '</div></div>' +
        (isNext ? '<span class="pill pill-accent">' + (d > 1 ? "in " + d + " days" : d === 1 ? "tomorrow" : "today") + '</span>'
                : past ? '<span class="pill pill-good">paid</span>'
                       : '<span class="pill">upcoming</span>') +
        '</div>' +
        '<div class="wcard-grid">' +
        '<div><span class="k">Days worked</span><span class="v">' + p.days + '</span></div>' +
        '<div><span class="k">Time worked</span><span class="v">' + hm(p.hours) + '</span></div>' +
        '<div><span class="k">Pay</span><span class="v accent">' + money(p.pay) + '</span></div>' +
        '</div></div>';
    }).join("") || '<p class="muted tiny">Set a first payday in Settings.</p>';

    $("ppTotalHours").textContent = dec(xround(th, 4));
    $("ppTotalTime").textContent = hmPlain(th);
    $("ppTotalPay").textContent = money(xround(tp, 2));
    $("ppNotes").innerHTML = WPL_SEED.notes.map(function (n) { return "<li>" + esc(n) + "</li>"; }).join("");
  }

  /* ---------- settings ---------- */
  function fillEmployerSelects() {
    var list = state.settings.employers || [];
    [["tsEmployer",true],["anEmployer",true],["edEmployer",false]].forEach(function(pair){
      var el=$(pair[0]); if(!el)return; var old=el.value;
      var signature=(pair[1]?"all|":"")+list.map(function(x){return x.id+":"+x.name;}).join("|");
      if(el.dataset.employerSignature!==signature){
        el.innerHTML=(pair[1]?'<option value="all">All employers</option>':"")+list.map(function(x){return '<option value="'+esc(x.id)+'">'+esc(x.name)+'</option>';}).join("");
        el.dataset.employerSignature=signature;
      }
      el.value=Array.prototype.some.call(el.options,function(o){return o.value===old;})?old:(pair[1]?"all":state.settings.defaultEmployerId);
    });
  }

  function renderEmployers() {
    var list=state.settings.employers||[];
    $("employerList").innerHTML=list.map(function(e){return '<div class="manage-row" data-employer="'+e.id+'"><input class="manage-color" type="color" value="'+esc(e.color||"#e85d0f")+'" aria-label="Job colour"><label><span>Job name</span><input data-key="name" value="'+esc(e.name)+'"></label><label><span>Hourly rate</span><input data-key="rate" type="number" min="0" step="0.01" value="'+Number(e.rate||0)+'"></label><label><span>Usual start</span><input data-key="start" type="time" value="'+esc(e.start||"")+'"></label><label><span>Usual finish</span><input data-key="finish" type="time" value="'+esc(e.finish||"")+'"></label><label><span>Unpaid break</span><input data-key="breakHours" type="number" min="0" step="0.25" value="'+Number(e.breakHours||0)+'"></label>'+(e.id===state.settings.defaultEmployerId?'<span class="pill pill-good">Default</span>':'<button class="icon-btn employer-delete" aria-label="Delete job">✕</button>')+'</div>';}).join("");
    [].forEach.call($("employerList").querySelectorAll(".manage-row"),function(row){
      var e=list.filter(function(x){return x.id===row.dataset.employer;})[0];
      [].forEach.call(row.querySelectorAll("input"),function(inp){inp.onchange=function(){var key=inp.dataset.key||"color",v=inp.value;if(key==="rate"||key==="breakHours")v=Math.max(0,parseFloat(v)||0);e[key]=v;if(e.id===state.settings.defaultEmployerId){state.settings.rate=e.rate;state.settings.breakHours=e.breakHours;state.settings.defaultStart=e.start;state.settings.defaultFinish=e.finish;}state.settingsUpdatedAt=Date.now();save();renderSettings();toast("Employer updated");};});
      var del=row.querySelector(".employer-delete");if(del)del.onclick=function(){confirmAsk("Delete this employer?","Its existing days will move to your default employer.",function(){var now=Date.now();state.entries.forEach(function(x){if(x.employerId===e.id){x.employerId=state.settings.defaultEmployerId;x.updatedAt=now;}});state.settings.employers=list.filter(function(x){return x.id!==e.id;});state.settingsUpdatedAt=now;save();renderSettings();toast("Employer removed");});};
    });
  }

  function renderTemplates() {
    var list=state.settings.shiftTemplates||[];
    $("templateList").innerHTML=list.map(function(t){return '<div class="manage-row template-row" data-template="'+t.id+'"><label><span>Name</span><input data-key="name" value="'+esc(t.name)+'"></label><label><span>Start</span><input data-key="start" type="time" value="'+esc(t.start)+'"></label><label><span>Finish</span><input data-key="finish" type="time" value="'+esc(t.finish)+'"></label><button class="icon-btn template-delete" aria-label="Delete template">✕</button></div>';}).join("")||'<p class="empty-inline">No templates yet.</p>';
    [].forEach.call($("templateList").querySelectorAll(".manage-row"),function(row){var t=list.filter(function(x){return x.id===row.dataset.template;})[0];[].forEach.call(row.querySelectorAll("input"),function(inp){inp.onchange=function(){t[inp.dataset.key]=inp.value;state.settingsUpdatedAt=Date.now();save();renderSettings();toast("Template updated");};});row.querySelector(".template-delete").onclick=function(){state.settings.shiftTemplates=list.filter(function(x){return x.id!==t.id;});state.settingsUpdatedAt=Date.now();save();renderSettings();toast("Template removed");};});
  }

  function renderSettings() {
    var s = state.settings;
    $("stFirstPayday").value = s.firstPayday;
    $("stCycle").value = s.cycleDays;

    fillCurrencies($("stCurrency"), s.currency);
    $("stSymbol").value = s.currencySymbol || "";
    $("stSymbolField").classList.toggle("hidden", s.currency !== "custom");
    $("stWeek").value = s.weekHours || 40;
    renderLeaveRates();
    fillEmployerSelects();
    renderEmployers();
    renderTemplates();
    renderBackup();
    renderCloud();

    $("historyHint").textContent = (state.history || []).length
      ? (state.history.length >= WPLSync.HISTORY_LIMIT ? "The latest " : "") +
        state.history.length + " change" + (state.history.length === 1 ? "" : "s") + " recorded."
      : "No changes recorded yet.";
    /* Show the version plainly — otherwise there is no way to tell from inside
       the app which build you are on, or whether an update actually landed. */
    $("aboutLine").textContent = "Work Payment Log " + appVersion() + " · " +
      entries().length + " days on this " + (window.WPLDesktop ? "Mac" : "phone");
    renderSync();
    renderNotifications();
    renderDataHealth();
    renderSectionSummaries();
  }

  /* Each collapsed section shows its own current value, so the page can be read
     without opening anything. */
  function renderSectionSummaries() {
    var s = state.settings;

    var jobs=s.employers||[],primary=employerFor({employerId:s.defaultEmployerId});
    setSum("sumPay", (s.currency||"EUR") + " · " + dec(Number(s.weekHours||40)) + "h full week");
    setSum("sumEmployers", jobs.length===1
      ? primary.name + " · " + money(Number(primary.rate||0)) + " an hour"
      : jobs.length + " jobs · " + primary.name + " is default");
    setSum("sumTemplates", (s.shiftTemplates||[]).length + ((s.shiftTemplates||[]).length===1?" template":" templates"));
    var enabled=Object.keys(s.notifications||{}).filter(function(k){return s.notifications[k];}).length;
    setSum("sumNotifications",enabled?enabled+" enabled":"All optional");
    var colour = APP_COLORS[s.appColor] || APP_COLORS.orange;
    var theme = localStorage.getItem(THEME_KEY) || "auto";
    setSum("sumAppearance", colour.label + " with " + (theme === "auto" ? "automatic" : theme) + " theme");
    [].forEach.call(document.querySelectorAll("[data-app-color]"), function (button) {
      var selected = button.dataset.appColor === (s.appColor || "orange");
      button.classList.toggle("on", selected);
      button.setAttribute("aria-pressed", selected ? "true" : "false");
    });

    setSum("sumPaydays", s.cycleDays
      ? "Every " + s.cycleDays + " days"
      : "Not set");

    var rates = s.leaveRates || {};
    var paid = 0, known = 0;
    for (var k in rates) { known++; if (Number(rates[k]) > 0) paid++; }
    setSum("sumLeave", known ? paid + " of " + known + " paid" : "");

    var hist = (state.history || []).length;
    setSum("sumHistory", hist
      ? hist + " change" + (hist === 1 ? "" : "s")
      : "Nothing recorded yet");
  }

  function setSum(id, text) {
    var el = $(id);
    if (el) el.textContent = text;
  }

  var reminderScheduleSignature = "";
  function completeTodayEntry() {
    return entries().some(function (e) {
      return e.date === toISO(today()) && (!!e.leave || (!!e.start && !!e.finish));
    });
  }
  function reminderDaysText(days) {
    var sorted = (days || []).slice().sort();
    if (sorted.join(",") === "1,2,3,4,5") return "weekdays";
    if (sorted.join(",") === "0,6") return "weekends";
    if (sorted.length === 7) return "every day";
    return sorted.map(function (d) { return DAYS[d]; }).join(", ");
  }
  function syncWorkReminderSchedule(force) {
    var s = state.settings, n = s.notifications || {};
    var loggedDate = completeTodayEntry() ? toISO(today()) : "";
    var days = (s.logReminderDays || [1,2,3,4,5]).join(",");
    var allowed=!!n.logHours&&!timeInQuietHours(s.logReminderTime||"18:00");
    var signature = [allowed, s.logReminderTime || "18:00", days, loggedDate,s.quietStart,s.quietEnd].join("|");
    if (!force && signature === reminderScheduleSignature) return;
    reminderScheduleSignature = signature;
    try {
      if (window.WPLBridge && window.WPLBridge.scheduleWorkReminder) {
        window.WPLBridge.scheduleWorkReminder(allowed, s.logReminderTime || "18:00", days, loggedDate);
      } else if (window.WPLDesktop && window.WPLDesktop.scheduleWorkReminder) {
        window.WPLDesktop.scheduleWorkReminder(allowed, s.logReminderTime || "18:00", days, loggedDate);
      }
    } catch (err) { /* reminders will be scheduled next time the app opens */ }
  }
  function renderNotifications(){
    var s=state.settings,n=s.notifications||{},days=s.logReminderDays||[1,2,3,4,5];
    [].forEach.call(document.querySelectorAll("[data-notify]"),function(inp){inp.checked=!!n[inp.dataset.notify];});
    $("stLogReminder").checked=!!n.logHours;
    $("stLogReminderTime").value=s.logReminderTime||"18:00";
    $("stQuietStart").value=s.quietStart||"22:00";$("stQuietEnd").value=s.quietEnd||"07:00";
    $("logReminderOptions").classList.toggle("hidden",!n.logHours);
    [].forEach.call(document.querySelectorAll("[data-reminder-day]"),function(button){
      var on=days.indexOf(Number(button.dataset.reminderDay))!==-1;
      button.classList.toggle("on",on);button.setAttribute("aria-pressed",on?"true":"false");
    });
    $("logReminderHint").textContent=timeInQuietHours(s.logReminderTime||"18:00")?"This time is inside quiet hours. Choose another time to turn the reminder on.":"At "+(s.logReminderTime||"18:00")+" on "+reminderDaysText(days)+".";
  }
  function renderDataHealth(){
    var cloud=(window.WPLCloud&&WPLCloud.cfg)?WPLCloud.cfg():{};
    var cloudCoverage=cloudSyncCoverage();
    $("healthBackup").textContent=backupCfg.lastAt?fmtAgo(backupCfg.lastAt):"Not yet";
    $("healthCloud").textContent=cloud.lastAt?fmtAgo(cloud.lastAt):(window.WPLCloud&&WPLCloud.signedIn()?"Waiting":"Not signed in");
    $("healthMac").textContent=window.WPLDesktop?"This Mac":cloudCoverage.synced?"Via cloud":!syncCfg.enabled?"Off":syncCfg.lastAt?fmtAgo(syncCfg.lastAt):"Not yet";
    $("healthPending").textContent=cloudCoverage.synced?0:(window.WPLDesktop||syncCfg.enabled?pendingCount():0); $("healthVersion").textContent=appVersion()||"Web";
    $("healthUpdate").textContent=isIOSApp()?(iosDistribution()==="appstore"?"App Store":"Personal build"):availableUpdateVersion||"Up to date";
    var safe=!!backupCfg.lastAt||!!cloud.lastAt||!entries().length;
    $("healthPill").textContent=safe?"Backed up":"Back up soon";
    $("healthPill").className="pill "+(safe?"pill-good":"pill-accent");
  }

  function timeInQuietHours(time){var s=state.settings,start=minutesOf(s.quietStart||"22:00"),end=minutesOf(s.quietEnd||"07:00"),value=minutesOf(time);if(start===null||end===null||value===null||start===end)return false;return start<end?value>=start&&value<end:value>=start||value<end;}
  function inQuietHours(){var now=new Date();return timeInQuietHours(pad2(now.getHours())+":"+pad2(now.getMinutes()));}
  function notifyOnce(key,title,text){
    if(inQuietHours())return;
    var stamp=toISO(today())+":"+key,storeKey="wpl.notice."+key;
    try{if(localStorage.getItem(storeKey)===stamp)return;localStorage.setItem(storeKey,stamp);}catch(err){/* still show */}
    if(window.WPLBridge&&window.WPLBridge.notifyUser){try{window.WPLBridge.notifyUser(key,title,text);return;}catch(err){/* fall through */}}
    if(window.Notification&&Notification.permission==="granted"){try{new Notification(title,{body:text,tag:key});}catch(err){/* in-app only */}}
  }
  function runFriendlyReminders(){
    var n=state.settings.notifications||{},todayEntry=entries().filter(function(e){return e.date===toISO(today());})[0];
    if(n.logHours&&!window.WPLBridge&&!window.WPLDesktop&&!completeTodayEntry()){
      var now=new Date(),parts=(state.settings.logReminderTime||"18:00").split(":"),due=now.getHours()*60+now.getMinutes()>=Number(parts[0])*60+Number(parts[1]);
      if(due&&(state.settings.logReminderDays||[]).indexOf(now.getDay())!==-1)notifyOnce("log-hours","Time to log your hours","Add today's start and finish times.");
    }
    if((n.missingFinish||n.incomplete)&&todayEntry&&!!todayEntry.start!==!!todayEntry.finish){notifyOnce("missing-finish","Finish time missing","Add the other time when your shift is complete.");}
    var np=nextPeriod(),days=np?dayDiff(today(),np.payday):99;
    if(n.payday&&days>=0&&days<=3)notifyOnce("payday-"+toISO(np.payday),"Payday "+(days===0?"is today":"is coming up"),money(np.pay)+" is currently expected.");
    if(n.payslip&&days===0)notifyOnce("payslip-"+toISO(np.payday),"Check your payslip","Compare it with the estimate in Take-home pay.");
    if(n.backup&&entries().length&&(!backupCfg.lastAt||Date.now()-backupCfg.lastAt>8*86400e3))notifyOnce("backup","Your data needs a backup","Open Settings to make a safe local copy.");
    if(n.sync&&!cloudSyncCoverage().synced&&(syncFailed||pendingCount()>2))notifyOnce("sync","Sync needs attention",pendingCount()+" changes are still waiting to sync.");
    if(n.updates&&availableUpdateVersion)notifyOnce("update-"+availableUpdateVersion,"App update available",availableUpdateVersion+" and is ready from Settings.");
    var spending=spendingModel();if(n.spending&&spending.budget>0&&spending.spent>spending.budget)notifyOnce("spending-"+spending.toISO,"Spending is over the plan",money(spending.spent-spending.budget)+" is over the money available.");
  }

  /* ============================================================
     EDITOR
     ============================================================ */
  var editingId = null;
  var savingEditor = false;
  var editorExtraBreaks = [];

  function renderEditorExtraBreaks() {
    $("edExtraBreaks").innerHTML=editorExtraBreaks.map(function(item,index){return '<div class="extra-break" data-break-index="'+index+'"><label class="field"><span>Break '+(index+2)+' (hours)</span><input class="extra-break-duration" type="number" min="0" step="0.25" inputmode="decimal" value="'+Number(item.duration||0)+'"></label><label class="break-paid"><input class="extra-break-paid" type="checkbox"'+(item.paid?' checked':'')+'><span>Paid</span></label><button class="icon-btn extra-break-remove" type="button" aria-label="Remove break">×</button></div>';}).join("");
    [].forEach.call($("edExtraBreaks").querySelectorAll(".extra-break"),function(row){var index=Number(row.dataset.breakIndex);row.querySelector(".extra-break-duration").oninput=updatePreview;row.querySelector(".extra-break-paid").onchange=updatePreview;row.querySelector(".extra-break-remove").onclick=function(){editorExtraBreaks.splice(index,1);renderEditorExtraBreaks();updatePreview();};});
  }

  function editorBreakItems() {
    var items=[],first=Math.max(0,Number($("edBreak").value)||0);
    if(first>0)items.push({id:"break-1",duration:first,paid:$("edBreakPaid").checked});
    [].forEach.call($("edExtraBreaks").querySelectorAll(".extra-break"),function(row,index){var duration=Math.max(0,Number(row.querySelector(".extra-break-duration").value)||0);if(duration>0)items.push({id:"break-"+(index+2),duration:duration,paid:row.querySelector(".extra-break-paid").checked});});
    return items;
  }

  function loadBreaksIntoEditor(entry) {
    var saved=Array.isArray(entry&&entry.breaks)?entry.breaks:[];
    if(saved.length){$("edBreak").value=Number(saved[0].duration)||0;$("edBreakPaid").checked=!!saved[0].paid;editorExtraBreaks=saved.slice(1).map(function(item){return {duration:Number(item.duration)||0,paid:!!item.paid};});}
    else{$("edBreak").value=typeof (entry&&entry.breakHours)==="number"?entry.breakHours:Number(employerFor(entry||{}).breakHours||0);$("edBreakPaid").checked=false;editorExtraBreaks=[];}
    renderEditorExtraBreaks();
  }

  function openEditor(id) {
    editingId = id || null;
    var e = id ? state.entries.filter(function (x) { return x.id === id; })[0] : null;
    $("sheetTitle").textContent = e ? "Edit day" : "Log a day";
    $("edDelete").classList.toggle("hidden", !e);
    $("edRepeatFields").classList.toggle("hidden",!!e);
    $("edHistory").classList.toggle("hidden",
      !e || !(state.history || []).some(function (r) { return r.date === e.date; }));

    if (!$("edType").options.length) {
      $("edType").innerHTML = LEAVE_TYPES.map(function (t) {
        return '<option value="' + t.key + '">' + esc(t.label) + "</option>";
      }).join("");
    }
    fillEmployerSelects();
    $("edTemplate").innerHTML='<option value="">Choose a template</option>'+(state.settings.shiftTemplates||[]).map(function(t){return '<option value="'+t.id+'">'+esc(t.name)+'</option>';}).join("");

    if (e) {
      $("edDate").value = e.date;
      $("edStart").value = e.start;
      $("edFinish").value = e.finish;
      $("edNotes").value = e.notes || "";
      $("edType").value = e.leave || "";
      $("edEmployer").value = e.employerId || state.settings.defaultEmployerId;
      $("edRate").value = typeof e.rateSnapshot === "number" ? e.rateSnapshot : Number(employerFor(e).rate || 0);
      loadBreaksIntoEditor(e);
      $("edStatus").value = ["planned", "completed", "approved", "paid"].indexOf(e.status) !== -1 ? e.status : (e.start && e.finish ? "completed" : "planned");
      $("edRateType").value=["standard","overtime","night","sunday","public","custom"].indexOf(e.rateType)!==-1?e.rateType:"standard";
      $("edMultiplier").value=e.payMultiplier==null?1:Math.max(0,Number(e.payMultiplier)||0);
    } else {
      var defaultEmployer = employerFor({ employerId: state.settings.defaultEmployerId });
      $("edDate").value = toISO(today());
      $("edStart").value = state.settings.defaultStart;
      $("edFinish").value = state.settings.defaultFinish;
      $("edNotes").value = "";
      $("edType").value = "";
      $("edEmployer").value = state.settings.defaultEmployerId;
      $("edRate").value = Number(defaultEmployer.rate || 0);
      $("edBreak").value = Number(defaultEmployer.breakHours || 0);
      $("edBreakPaid").checked=false;editorExtraBreaks=[];
      $("edStatus").value = "completed";
      $("edRateType").value="standard";$("edMultiplier").value="1";
    }
    savingEditor = false;
    $("edRepeat").value="none";$("edRepeatWeeks").value="4";$("edRepeatWeeksWrap").classList.add("hidden");
    renderEditorExtraBreaks();
    $("edMultiplierWrap").classList.toggle("hidden",$("edRateType").value!=="custom");
    $("edSave").disabled = false;
    updatePreview();
    document.body.classList.add("editor-open");
    $("scrim").classList.remove("hidden");
    $("sheet").classList.remove("hidden");
  }

  function openTodayEditor() {
    var iso=toISO(today()),e=entries().filter(function(x){return x.date===iso&&!!x.start!==!!x.finish;})[0];
    openEditor(e?e.id:null);
    $("edDate").value=iso;
  }

  function closeEditor() {
    $("scrim").classList.add("hidden");
    $("sheet").classList.add("hidden");
    document.body.classList.remove("editor-open");
    editingId = null;
    savingEditor = false;
    $("edSave").disabled = false;
  }

  function updatePreview() {
    var editorBreaks=editorBreakItems();
    var draft = { start: $("edStart").value, finish: $("edFinish").value,
      employerId: $("edEmployer").value || state.settings.defaultEmployerId,
      rateSnapshot: Math.max(0, Number($("edRate").value) || 0),
      breakHours: editorBreaks.reduce(function(sum,item){return sum+(item.paid?0:item.duration);},0), breaks:editorBreaks,
      status: ["planned", "completed", "approved", "paid"].indexOf($("edStatus").value) !== -1 ? $("edStatus").value : "completed",
      payMultiplier:Math.min(10,Math.max(0,Number($("edMultiplier").value)||0)),rateType:$("edRateType").value||"standard",
      leave: $("edType").value || "" };
    var calc = calculateEntry(draft), h = calc.hours, p = calc.pay;
    $("edTime").textContent = hm(h);
    $("edHours").textContent = dec(h);
    $("edPay").textContent = money(p);

    var info = leaveInfo($("edType").value);
    var sm = minutesOf(draft.start), fm = minutesOf(draft.finish);
    var note = "";
    if (info.key) {
      $("edBreakNote").textContent = info.pct <= 0
        ? info.label + " is set to unpaid, so this day is worth nothing."
        : h > 0
          ? info.label + ". " + hm(h) + " at " + info.pct + "% of a normal day is " + money(p) +
            (info.pct === 100 ? "." : ". Change that rate in Settings.")
          : info.label + ". Add times to be paid for it.";
      return;
    }
    if (sm !== null && fm !== null) {
      if (calc.invalidBreak) note = "The breaks are longer than the shift. Shorten them before saving.";
      else note = hmPlain(calc.shiftHours) + " on site" + (calc.overnight ? " over midnight" : "") +
        ", less " + hmPlain(calc.breakHours) + " unpaid break"+(calc.paidBreakHours?" and "+hmPlain(calc.paidBreakHours)+" paid break":"")+", at " + money(calc.hourlyRate) + " an hour"+(calc.payMultiplier!==1?" × "+calc.payMultiplier.toFixed(2):"")+".";
    } else if (sm === null && fm === null) {
      note = "No times means 0 hours. Use this for a day off.";
    } else {
      note = "Both a start and a finish are needed for the day to count.";
    }
    $("edBreakNote").textContent = note;
  }

  function shiftInterval(entry) {
    var day = parseDate(entry && entry.date);
    var start = minutesOf(entry && entry.start), finish = minutesOf(entry && entry.finish);
    if (!day || start === null || finish === null || start === finish) return null;
    if (finish < start) finish += 24 * 60;
    var base = day.getTime();
    return { start: base + start * 60000, end: base + finish * 60000 };
  }

  function overlappingShift(entry, ignoreId) {
    var candidate = shiftInterval(entry);
    if (!candidate || entry.leave) return null;
    return entries().filter(function (other) {
      if (other.id === ignoreId || other.leave) return false;
      var interval = shiftInterval(other);
      return interval && candidate.start < interval.end && interval.start < candidate.end;
    })[0] || null;
  }

  function saveEditor() {
    if (savingEditor) return;
    var date = $("edDate").value;
    if (!date) { toast("Pick a date first"); return; }
    if (!validStoredDate(date)) { toast("That date is not valid"); return; }
    var editorBreaks=editorBreakItems();
    var rec = {
      id: editingId || uid(),
      date: date,
      start: $("edStart").value || "",
      finish: $("edFinish").value || "",
      notes: ($("edNotes").value || "").trim(),
      leave: $("edType").value || "", employerId: $("edEmployer").value || state.settings.defaultEmployerId,
      rateSnapshot: Math.max(0, Number($("edRate").value) || 0),
      breakHours: xround(editorBreaks.reduce(function(sum,item){return sum+(item.paid?0:item.duration);},0),4), breaks:editorBreaks,
      status: ["planned", "completed", "approved", "paid"].indexOf($("edStatus").value) !== -1 ? $("edStatus").value : "completed",
      payMultiplier:Math.min(10,Math.max(0,Number($("edMultiplier").value)||0)),rateType:$("edRateType").value||"standard",
      updatedAt: Date.now(),
      deleted: false
    };
    var calc = calculateEntry(rec);
    if (calc.invalidBreak) { toast("The break cannot be longer than the shift"); return; }
    var overlap = overlappingShift(rec, editingId);
    if (overlap) { toast("This overlaps another shift on " + fmtDate(parseDate(overlap.date))); return; }

    savingEditor = true;
    $("edSave").disabled = true;
    var existing = editingId ? state.entries.filter(function (x) { return x.id === editingId; })[0] : null;
    var previous = snapshotOf(existing);
    if (existing) Object.assign(existing, rec);
    else state.entries.push(rec);
    recordHistory("day", date, previous, snapshotOf(existing || rec));
    var repeated=0,skipped=0;
    if(!existing&&$("edRepeat").value==="weekly"){
      var repeatWeeks=Math.max(1,Math.min(52,parseInt($("edRepeatWeeks").value,10)||1));
      for(var repeatIndex=1;repeatIndex<=repeatWeeks;repeatIndex++){
        var repeatedRec=Object.assign({},rec,{id:uid(),date:toISO(addDays(parseDate(rec.date),repeatIndex*7)),status:"planned",updatedAt:Date.now()+repeatIndex,breaks:rec.breaks.map(function(item){return Object.assign({},item,{id:"break-"+uid()});})});
        if(overlappingShift(repeatedRec,null)){skipped++;continue;}
        state.entries.push(repeatedRec);recordHistory("day",repeatedRec.date,null,snapshotOf(repeatedRec),"repeat");repeated++;
      }
    }
    save();
    closeEditor();
    render();
    var h = hoursOf(existing || rec);
    var kind = leaveInfo(rec.leave);
    var savedMessage=rec.leave
      ? kind.label + (h > 0 ? " · " + money(payOf(existing || rec)) : " · unpaid") + " saved"
      : h > 0 ? hm(h) + " · " + money(payOf(existing || rec)) + " saved"
              : "Day off saved";
    if(repeated)savedMessage+=" · "+repeated+" planned";
    if(skipped)savedMessage+=" · "+skipped+" overlap skipped";
    toast(savedMessage);
  }

  function deleteEditor() {
    var targetId = editingId;
    confirmAsk("Delete this day?", "The entry will be removed from the timesheet.", function () {
      var deletedCopy = null;
      /* tombstone rather than drop, so the delete reaches the other device */
      state.entries.forEach(function (x) {
        if (x.id === targetId) {
          deletedCopy = Object.assign({}, x);
          recordHistory("day", x.date, snapshotOf(x),
            { id: x.id, date: x.date, start: "", finish: "", notes: x.notes || "", leave: x.leave || "",
              employerId: x.employerId || state.settings.defaultEmployerId,
              rateSnapshot: x.rateSnapshot, breakHours: x.breakHours, breaks:x.breaks, status:x.status, payMultiplier:x.payMultiplier,rateType:x.rateType,deleted: true }, "delete");
          x.deleted = true; x.updatedAt = Date.now();
        }
      });
      save(); closeEditor(); render();
      toastUndo("Day deleted", function () {
        var current = state.entries.filter(function (x) { return x.id === targetId; })[0];
        if (!current || !deletedCopy) return;
        var before = snapshotOf(current);
        Object.assign(current, deletedCopy, { deleted: false, updatedAt: Date.now() });
        recordHistory("day", current.date, before, snapshotOf(current), "undo");
        save(); render(); toast("Delete undone");
      });
    });
  }

  /* ============================================================
     TOAST + CONFIRM
     ============================================================ */
  var toastTimer;
  function toast(msg) {
    var t = $("toast");
    t.textContent = msg;
    t.classList.remove("hidden");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.classList.add("hidden"); }, 2200);
  }
  function toastUndo(msg,cb){var t=$("toast");t.innerHTML=esc(msg)+' <button type="button">Undo</button>';t.classList.remove("hidden");clearTimeout(toastTimer);t.querySelector("button").onclick=function(){t.classList.add("hidden");cb();};toastTimer=setTimeout(function(){t.classList.add("hidden");},6000);}

  var confirmCb = null;
  function confirmAsk(title, text, cb) {
    $("confirmTitle").textContent = title;
    $("confirmText").textContent = text;
    confirmCb = cb;
    $("confirmWrap").classList.remove("hidden");
  }

  /* ============================================================
     EXPORT / IMPORT
     ============================================================ */
  function csvCell(v) {
    v = String(v == null ? "" : v);
    return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
  }

  function buildCsv() {
    var s = state.settings;
    var lines = [];
    lines.push(["WORK PAYMENT LOG"].join(","));
    lines.push([]);
    lines.push(["Hourly rate (EUR)", s.rate, "Default start", s.defaultStart, "Default finish", s.defaultFinish].map(csvCell).join(","));
    lines.push(["Unpaid break per day (hrs)", s.breakHours, "First payday", s.firstPayday, "Cycle length (days)", s.cycleDays].map(csvCell).join(","));
    lines.push([]);
    lines.push(["Date", "Day", "Type", "Start", "Finish", "Hours (decimal)", "Time worked", "Pay (EUR)", "Notes"].join(","));
    sorted().forEach(function (e) {
      var d = parseDate(e.date), h = hoursOf(e);
      lines.push([fmtDate(d), dayName(d), leaveInfo(e.leave).label, e.start, e.finish,
        dec(h), hm(h), dec(payOf(e)), e.notes].map(csvCell).join(","));
    });
    var g = grandTotals();
    lines.push(["TOTAL", "", "", "", "", dec(g.hours), hmPlain(g.hours), dec(g.pay), ""].map(csvCell).join(","));
    lines.push([]);
    lines.push(["WEEKLY SUMMARY"].join(","));
    lines.push(["Week starting (Mon)", "Week ending (Sun)", "Hours (decimal)", "Time worked", "Pay (EUR)"].join(","));
    weeks().forEach(function (w) {
      lines.push([fmtDate(w.start), fmtDate(w.end), dec(w.hours), hm(w.hours), dec(w.pay)].map(csvCell).join(","));
    });
    lines.push([]);
    lines.push(["PAY PERIODS"].join(","));
    lines.push(["Payday", "Day", "Period start", "Period end", "Days worked", "Hours (decimal)", "Time worked", "Pay (EUR)"].join(","));
    periods().forEach(function (p) {
      lines.push([fmtDate(p.payday), DAYS[p.payday.getDay()], fmtDate(p.start), fmtDate(p.end),
        p.days, dec(p.hours), hm(p.hours), dec(p.pay)].map(csvCell).join(","));
    });
    return "﻿" + lines.join("\n");
  }

  function buildAnalyticsCsv() {
    var range=analyticsDates(), list=analyticsEntries(range), t=totalsOf(list), ws=groupWeeks(list), ms=groupMonths(list), lines=[];
    lines.push("WORK PAYMENT LOG ANALYTICS");
    lines.push(["From",range.fromISO,"To",range.toISO].map(csvCell).join(","));
    lines.push(["Total hours",dec(t.hours),"Gross pay",dec(t.pay),"Days worked",t.days,"Leave days",t.leaveDays].map(csvCell).join(","));
    lines.push("");lines.push("WEEKS");
    lines.push(["Week","Days","Regular hours","Overtime hours","Leave hours","Gross pay"].join(","));
    ws.forEach(function(w){lines.push([toISO(w.start)+" to "+toISO(w.end),w.days,dec(w.regular),dec(w.overtime),dec(w.leaveHours),dec(w.pay)].map(csvCell).join(","));});
    lines.push("");lines.push("MONTHS");lines.push(["Month","Days","Hours","Leave days","Gross earnings","Average weekly earnings"].join(","));
    ms.forEach(function(m){lines.push([m.key,m.days,dec(m.hours),m.leaveDays,dec(m.pay),dec(m.avg)].map(csvCell).join(","));});
    lines.push("");lines.push("ENTRIES");lines.push(["Date","Employer","Type","Start","Finish","Hours","Pay","Notes"].join(","));
    list.sort(function(a,b){return a.date.localeCompare(b.date);}).forEach(function(e){lines.push([e.date,employerFor(e).name,leaveInfo(e.leave).label,e.start,e.finish,dec(hoursOf(e)),dec(payOf(e)),e.notes].map(csvCell).join(","));});
    return "﻿"+lines.join("\n");
  }

  /* Lightweight valid PDF: plain text keeps the mobile bundle dependency-free. */
  function buildAnalyticsPdf() {
    var r=analyticsDates(),list=analyticsEntries(r),t=totalsOf(list),lines=["Work Payment Log Analytics",r.fromISO+" to "+r.toISO,"", "Total hours: "+dec(t.hours),"Gross pay: "+dec(t.pay)+" "+(state.settings.currency||"EUR"),"Days worked: "+t.days,"Leave days: "+t.leaveDays,"", "Weeks"];
    groupWeeks(list).forEach(function(w){lines.push(toISO(w.start)+" - "+toISO(w.end)+"   "+dec(w.hours)+" h   "+dec(w.pay)+" "+(state.settings.currency||"EUR"));});
    lines.push("","Months");groupMonths(list).forEach(function(m){lines.push(m.key+"   "+m.days+" days   "+dec(m.hours)+" h   "+dec(m.pay)+" "+(state.settings.currency||"EUR"));});
    lines=lines.slice(0,52).map(function(x){return String(x).replace(/[^\x20-\x7E]/g," ").replace(/[()\\]/g,"\\$&");});
    var content="BT /F1 11 Tf 48 790 Td 15 TL "+lines.map(function(x,i){return (i?"T* ":"")+"("+x+") Tj";}).join(" ")+" ET";
    var objs=["<< /Type /Catalog /Pages 2 0 R >>","<< /Type /Pages /Kids [3 0 R] /Count 1 >>","<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>","<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>","<< /Length "+content.length+" >>\nstream\n"+content+"\nendstream"];
    var pdf="%PDF-1.4\n",offsets=[0];objs.forEach(function(o,i){offsets.push(pdf.length);pdf+=(i+1)+" 0 obj\n"+o+"\nendobj\n";});var xref=pdf.length;pdf+="xref\n0 "+(objs.length+1)+"\n0000000000 65535 f \n";for(var i=1;i<=objs.length;i++)pdf+=("0000000000"+offsets[i]).slice(-10)+" 00000 n \n";return pdf+="trailer << /Size "+(objs.length+1)+" /Root 1 0 R >>\nstartxref\n"+xref+"\n%%EOF";
  }

  function saveFile(name, text, mime) {
    /* Android WebView bridge */
    if (window.WPLBridge && window.WPLBridge.saveFile) {
      try {
        var res = window.WPLBridge.saveFile(name, text);
        toast(res || "Saved to Downloads");
      } catch (err) { toast("Could not save the file"); }
      return;
    }
    /* Electron bridge */
    if (window.WPLDesktop && window.WPLDesktop.saveFile) {
      window.WPLDesktop.saveFile(name, text).then(function (res) {
        if (res) toast("Saved to " + res);
      });
      return;
    }
    /* browser fallback */
    var blob = new Blob([text], { type: mime || "text/plain;charset=utf-8" });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1500);
    toast("Downloaded " + name);
  }

  function applyImport(text) {
    var data;
    try { data = JSON.parse(text); } catch (err) { toast("That file is not a valid backup"); return; }
    var problem = validateBackup(data);
    if (problem) { toast(problem); return; }
    var live = data.entries.filter(function(e){return !e.deleted;});
    var dates = live.map(function(e){return e.date;}).sort();
    var range = dates.length ? " from " + fmtDate(parseDate(dates[0])) + " to " + fmtDate(parseDate(dates[dates.length-1])) : "";
    confirmAsk("Restore this backup?", live.length + (live.length === 1 ? " shift" : " shifts") + range + " will replace the data on this device.", function () {
      if (entries().length) runBackup(false); // best-effort safety copy before replacement
      var now = Date.now();
      state.settings = Object.assign({}, WPL_SEED.settings, data.settings || {});
      state.settingsUpdatedAt = now;
      state.settingsFieldUpdatedAt = {};
      Object.keys(state.settings).forEach(function(key){state.settingsFieldUpdatedAt[key]=now;});
      state.entries = data.entries.map(function (e, index) {
        return copyEntry(e, index, now);
      });
      state.history = Array.isArray(data.history) ? data.history.slice(-WPLSync.HISTORY_LIMIT) : [];
      ensureProductData(state);
      save(); render();
      toast("Restored " + entries().length + (entries().length === 1 ? " shift" : " shifts"));
    });
  }

  function validStoredDate(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ""))) return false;
    var parsed = parseDate(value);
    return !!parsed && toISO(parsed) === value;
  }

  function validStoredTime(value) {
    return value === "" || /^([01]\d|2[0-3]):[0-5]\d$/.test(String(value || ""));
  }

  function validateBackup(data) {
    if (!data || typeof data !== "object" || !data.settings || typeof data.settings !== "object" || !Array.isArray(data.entries)) {
      return "That file is not a Work Payment Log backup";
    }
    if (data.entries.length > 100000) return "That backup contains too many shifts to open safely";
    if(data.integrity&&data.integrity!==backupIntegrity(data))return "That backup is damaged and was not restored";
    var ids = {};
    for (var i = 0; i < data.entries.length; i++) {
      var e = data.entries[i];
      if (!e || typeof e !== "object" || !validStoredDate(e.date) || !validStoredTime(e.start || "") || !validStoredTime(e.finish || "")) {
        return "That backup has a damaged shift at item " + (i + 1);
      }
      if (e.id && ids[e.id]) return "That backup contains duplicate shift identifiers";
      if (e.id) ids[e.id] = true;
      if (e.rateSnapshot != null && (!isFinite(Number(e.rateSnapshot)) || Number(e.rateSnapshot) < 0)) return "That backup has an invalid hourly rate";
      if (e.breakHours != null && (!isFinite(Number(e.breakHours)) || Number(e.breakHours) < 0)) return "That backup has an invalid break";
      if (e.breaks != null && !Array.isArray(e.breaks)) return "That backup has damaged break details";
      if (Array.isArray(e.breaks) && e.breaks.some(function(item){return !item||!isFinite(Number(item.duration))||Number(item.duration)<0;})) return "That backup has an invalid break";
      if (e.payMultiplier != null && (!isFinite(Number(e.payMultiplier)) || Number(e.payMultiplier)<0 || Number(e.payMultiplier)>10)) return "That backup has an invalid pay multiplier";
    }
    return "";
  }

  function copyEntry(e, index, fallbackUpdatedAt) {
    return {
      id: String(e.id || legacyEntryId(e, index)), date: e.date,
      start: e.start || "", finish: e.finish || "",
      notes: String(e.notes || "").slice(0, 120), leave: e.leave || "",
      employerId: e.employerId || state.settings.defaultEmployerId,
      rateSnapshot: typeof e.rateSnapshot === "number" ? Math.max(0, e.rateSnapshot) : undefined,
      breakHours: typeof e.breakHours === "number" ? Math.max(0, e.breakHours) : undefined,
      breaks: Array.isArray(e.breaks) ? e.breaks.map(function(item){return {id:String(item.id||uid()),duration:Math.max(0,Number(item.duration)||0),paid:!!item.paid};}) : undefined,
      status: ["planned", "completed", "approved", "paid"].indexOf(e.status) !== -1 ? e.status : (e.start && e.finish ? "completed" : "planned"),
      payMultiplier:e.payMultiplier==null?1:Math.min(10,Math.max(0,Number(e.payMultiplier)||0)),rateType:["standard","overtime","night","sunday","public","custom"].indexOf(e.rateType)!==-1?e.rateType:"standard",
      updatedAt: Number(e.updatedAt) || fallbackUpdatedAt || 0, deleted: !!e.deleted
    };
  }

  /* ============================================================
     SYNC
     The Mac hosts; the phone talks to it. Both sides send their whole
     store and take back the merged result — with 50-odd days that is a
     few kilobytes, and it removes every "what changed since when" bug.
     ============================================================ */
  function storeForSync() {
    return {
      settings: state.settings,
      settingsUpdatedAt: state.settingsUpdatedAt || 0,
      settingsFieldUpdatedAt: state.settingsFieldUpdatedAt || {},
      entries: state.entries,
      history: state.history || []
    };
  }

  /* Take what came back, without re-triggering a push.
     Always merge rather than replace: an empty or stale answer — a host with no
     file yet, a phone that has never synced — must never wipe what is here. */
  function adoptMerged(incoming, peerLabel) {
    if (!incoming || !Array.isArray(incoming.entries)) return false;
    var merged = WPLSync.mergeStores(storeForSync(), incoming);
    var before = WPLSync.signature(storeForSync());
    var after = WPLSync.signature(merged);
    if (before === after) return false;

    applyingRemote = true;
    state.settings = Object.assign({}, WPL_SEED.settings, merged.settings || {});
    state.settingsUpdatedAt = merged.settingsUpdatedAt || 0;
    state.settingsFieldUpdatedAt = merged.settingsFieldUpdatedAt || {};
    state.entries = merged.entries.map(function (e, index) { return copyEntry(e, index, 0); });
    state.history = merged.history || [];
    ensureProductData(state);
    save();
    applyingRemote = false;
    render();
    flashTotals();
    backupIfDue();
    if (peerLabel) toast("Synced with " + peerLabel);
    return true;
  }

  /* When figures change because the other device sent something, say so visually
     rather than letting numbers silently rewrite themselves. */
  function flashTotals() {
    if (reduceMotion) return;
    ["ovTotalPay", "tsPay", "wkTotalPay", "ppTotalPay"].forEach(function (id) {
      var el = $(id);
      if (!el) return;
      el.classList.remove("flash");
      void el.offsetWidth;              // restart the animation
      el.classList.add("flash");
    });
  }

  /* Mac: every local save goes straight into the shared store the phone reads.
     Phone: every local save nudges the Mac a moment later, so a day logged here
     is on the Mac within seconds rather than at the next scheduled check. */
  var pushTimer = null;
  function pushToHost() {
    if (window.WPLDesktop && window.WPLDesktop.push) {
      window.WPLDesktop.push(storeForSync()).then(function (merged) {
        adoptMerged(merged, null);
      }).catch(function () { /* the host will catch up on the next save */ });
      return;
    }
    if (!syncCfg.enabled || !syncCfg.auto || !syncCfg.host || !syncCfg.code) return;
    clearTimeout(pushTimer);
    pushTimer = setTimeout(function () { syncNow(false); }, 1500);
  }

  /* Phone: reach the Mac over the local network. */
  var syncing = false;
  var syncFailed = false;             // last attempt could not reach the Mac
  var failStreak = 0;                 // consecutive misses, drives re-discovery
  var hostApk = null;                 // the build the Mac is carrying, if newer
  var availableUpdateVersion = "";
  var nativeUpdateInfo = null;
  var updateChecking = false;

  function isIOSApp() {
    try {
      return !!(window.WPLBridge && window.WPLBridge.platform && window.WPLBridge.platform() === "ios");
    } catch (err) { return false; }
  }

  function iosDistribution() {
    if (!isIOSApp()) return "";
    try {
      return window.WPLBridge.distribution ? String(window.WPLBridge.distribution()) : "personal";
    } catch (err) { return "personal"; }
  }

  function automaticUpdatesEnabled() {
    try {
      if (window.WPLBridge && window.WPLBridge.autoUpdatesEnabled) {
        return !!window.WPLBridge.autoUpdatesEnabled();
      }
    } catch (err) { /* native host may still be starting */ }
    return nativeUpdateInfo && typeof nativeUpdateInfo.autoEnabled === "boolean"
      ? nativeUpdateInfo.autoEnabled : true;
  }

  function refreshNativeUpdateInfo() {
    if (window.WPLDesktop && window.WPLDesktop.updateInfo) {
      return window.WPLDesktop.updateInfo().then(function (info) {
        nativeUpdateInfo = info || null;
        updateChecking = !!(info && info.checking);
        renderUpdate();
        return info;
      }).catch(function () { return null; });
    }
    renderUpdate();
    return Promise.resolve(null);
  }

  function checkForUpdate(manual) {
    /* App Store releases are checked and installed by iOS. The app cannot
       inspect or replace its own signed bundle, so do not pretend a web check
       can force the system updater. */
    if (isIOSApp()) {
      if (manual) toast(iosDistribution() === "appstore"
        ? "Updates are managed by the App Store"
        : "Free personal builds must be reinstalled after seven days");
      return Promise.resolve(null);
    }
    updateChecking = true;
    renderUpdate();
    if (window.WPLDesktop && window.WPLDesktop.checkUpdate) {
      return window.WPLDesktop.checkUpdate(!!manual).then(function (info) {
        nativeUpdateInfo = info || null;
        updateChecking = false;
        renderUpdate();
        if (manual) toast(info && info.available ? "Update found" : "You are up to date");
        return info;
      }).catch(function () {
        updateChecking = false;
        renderUpdate();
        if (manual) toast("Could not check for updates");
      });
    }
    try {
      if (window.WPLBridge && window.WPLBridge.checkForOnlineUpdate) {
        window.WPLBridge.checkForOnlineUpdate();
        if (manual) toast("Checking for an update in the background");
        setTimeout(function () {
          updateChecking = false;
          renderUpdate();
        }, 1800);
      }
    } catch (err) {
      updateChecking = false;
      renderUpdate();
      if (manual) toast("Could not check for updates");
    }
    return Promise.resolve(null);
  }
  function syncNow(manual, skipCloudFallback) {
    if (syncing) return Promise.resolve(false);
    if (!syncCfg.enabled) {
      if (manual) toast("Turn on Mac sync first");
      return Promise.resolve(false);
    }
    if (!syncCfg.host || !syncCfg.code) {
      if (manual) toast("Add your Mac's address and code first");
      return Promise.resolve(false);
    }
    syncing = true;
    var owedAtStart = syncCfg.dirtyCount || 0;
    setSyncStatus("Syncing…");

    var ctrl = typeof AbortController !== "undefined" ? new AbortController() : null;
    var timer = setTimeout(function () { if (ctrl) ctrl.abort(); }, 8000);

    return fetch(hostUrl(syncCfg.host) + "/sync", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": "Bearer " + syncCfg.code },
      body: JSON.stringify(WPLSync.payload(storeForSync(), device)),
      signal: ctrl ? ctrl.signal : undefined
    }).then(function (res) {
      if (res.status === 401) throw new Error("bad-code");
      if (!res.ok) throw new Error("http-" + res.status);
      return res.json();
    }).then(function (merged) {
      syncCfg.lastAt = Date.now();
      syncCfg.lastPeer = merged.hostName || "Mac";
      if (merged.streamToken && merged.streamToken !== streamToken) {
        streamToken = merged.streamToken;
        closeStream();                            // reopen against the new token
      }
      hostApk = merged.apk || null;
      openStream();
      /* anything saved while the request was in flight stays owed */
      syncCfg.dirtyCount = Math.max(0, (syncCfg.dirtyCount || 0) - owedAtStart);
      saveSyncCfg();
      var changed = adoptMerged(merged, manual ? syncCfg.lastPeer : null);
      if (manual && !changed) toast("Already up to date");
      syncFailed = false;
      failStreak = 0;
      setSyncStatus("");
      renderSync();
      return true;
    }).catch(function (err) {
      var msg = err && err.message === "bad-code"
        ? "That code was not accepted"
        : "Could not reach your Mac";
      if (manual) toast(msg);
      syncFailed = true;
      failStreak++;
      setSyncStatus(msg);
      /* Two misses in a row usually means the Mac is on a different address,
         not that it is off. Go and look, once, instead of nagging. */
      if (failStreak === 2 && err && err.message !== "bad-code") relocateHost();
      /* The Mac being unavailable must not strand the change. A signed-in
         cloud account is the fallback peer and can take over immediately. */
      if (!skipCloudFallback) cloudSync(false);
      return false;
    }).then(function (ok) {
      clearTimeout(timer);
      syncing = false;
      return ok;
    });
  }

  /* ---- live channel ----
     Polling every few minutes is fine for catching up, but a change made on
     the Mac while you are looking at the phone should appear now, not in three
     minutes. The phone keeps a stream open to the Mac and syncs on each nudge. */
  var stream = null;
  var streamToken = "";

  function openStream() {
    if (stream || !streamToken || typeof EventSource === "undefined") return;
    if (!syncCfg.enabled || !syncCfg.host || !syncCfg.auto) return;
    try {
      stream = new EventSource(hostUrl(syncCfg.host) + "/events?t=" + encodeURIComponent(streamToken));
    } catch (err) { stream = null; return; }

    stream.onmessage = function (ev) {
      var msg;
      try { msg = JSON.parse(ev.data); } catch (err) { return; }
      if (msg.by === device.id) return;          // our own push coming back round
      if (msg.by === "hello") return;            // just the handshake
      syncNow(false);
    };
    stream.onerror = function () {
      closeStream();
      /* EventSource retries by itself, but the token may have died with the Mac
         app; a fresh sync gets a new one and reopens. */
      setTimeout(function () { if (syncCfg.enabled && syncCfg.auto) syncNow(false); }, 5000);
    };
  }

  function closeStream() {
    if (stream) { try { stream.close(); } catch (err) { /* already gone */ } }
    stream = null;
  }

  function hostUrl(host) {
    host = String(host || "").trim().replace(/^https?:\/\//, "").replace(/\/+$/, "");
    if (host.indexOf(":") === -1) host += ":" + WPLSync.PORT;
    return "http://" + host;
  }

  function setSyncStatus(text) {
    var el = $("syncStatus");
    if (!el) return;
    if (text === "Syncing…") {
      el.innerHTML = '<span class="syncing-dot"></span>Syncing…';
    } else {
      el.textContent = text;
    }
  }

  /* Your Mac's address changes whenever either device joins a different Wi-Fi,
     and re-typing it is the only reason pairing ever needs touching again. So
     when the saved address stops answering, go and find it: sweep the phone's
     own subnet, and adopt the first host that accepts the code we already hold.
     The code itself is unchanged, so nothing is asked of you. */
  var relocating = false;
  function relocateHost() {
    if (relocating || !window.WPLBridge || !window.WPLBridge.localIp) return Promise.resolve(false);
    if (!syncCfg.code) return Promise.resolve(false);

    var ip = "";
    try { ip = window.WPLBridge.localIp(); } catch (err) { ip = ""; }
    if (!ip || ip.indexOf(".") === -1) return Promise.resolve(false);

    relocating = true;
    setSyncStatus("Your Mac moved. Looking for it…");
    var base = ip.split(".").slice(0, 3).join(".");
    var found = false;
    var checks = [];

    for (var i = 1; i < 255; i++) checks.push(probe(base + "." + i));

    return Promise.all(checks).then(function () {
      relocating = false;
      if (!found) setSyncStatus("Could not find your Mac on this network");
      return found;
    });

    function probe(addr) {
      if (addr === syncCfg.host) return Promise.resolve();
      var ctrl = typeof AbortController !== "undefined" ? new AbortController() : null;
      var t = setTimeout(function () { if (ctrl) ctrl.abort(); }, 2500);
      return fetch("http://" + addr + ":" + WPLSync.PORT + "/ping",
        { signal: ctrl ? ctrl.signal : undefined })
        .then(function (r) { return r.ok ? r.json() : null; })
        .then(function (info) {
          if (found || !info || info.app !== "work-payment-log") return;
          found = true;
          syncCfg.host = addr;
          saveSyncCfg();
          if ($("syncHost")) $("syncHost").value = addr;
          setSyncStatus("Found your Mac again at " + addr);
          toast("Reconnected to " + (info.hostName || "your Mac"));
          return syncNow(false);
        })
        .catch(function () { /* nothing listening there */ })
        .then(function () { clearTimeout(t); });
    }
  }

  /* Sweep the phone's own /24 looking for a Mac that answers /ping. */
  function discoverHost() {
    if (!window.WPLBridge || !window.WPLBridge.localIp) {
      toast("Type your Mac's address instead");
      return;
    }
    var ip = "";
    try { ip = window.WPLBridge.localIp(); } catch (err) { ip = ""; }
    if (!ip || ip.indexOf(".") === -1) { toast("Connect to the same wireless network as your Mac"); return; }

    var base = ip.split(".").slice(0, 3).join(".");
    setSyncStatus("Looking for your Mac on " + base + ".*");
    $("btnSyncFind").disabled = true;

    var found = false;
    var checks = [];
    for (var i = 1; i < 255; i++) {
      checks.push(ping(base + "." + i));
    }
    Promise.all(checks).then(function () {
      $("btnSyncFind").disabled = false;
      if (!found) setSyncStatus("No Mac found. Check that the app is open on it.");
    });

    function ping(addr) {
      var ctrl = typeof AbortController !== "undefined" ? new AbortController() : null;
      var t = setTimeout(function () { if (ctrl) ctrl.abort(); }, 2500);
      return fetch("http://" + addr + ":" + WPLSync.PORT + "/ping",
        { signal: ctrl ? ctrl.signal : undefined })
        .then(function (r) { return r.ok ? r.json() : null; })
        .then(function (info) {
          if (found || !info || info.app !== "work-payment-log") return;
          found = true;
          syncCfg.host = addr;
          saveSyncCfg();
          $("syncHost").value = addr;
          setSyncStatus("Found " + (info.hostName || "your Mac") + " at " + addr);
          toast("Found " + (info.hostName || "your Mac"));
        })
        .catch(function () { /* nothing there */ })
        .then(function () { clearTimeout(t); });
    }
  }

  function fmtAgo(ts) {
    if (!ts) return "never";
    var mins = Math.floor((Date.now() - ts) / 60000);
    if (mins < 1) return "just now";
    if (mins < 60) return mins + " min ago";
    var hrs = Math.floor(mins / 60);
    if (hrs < 24) return hrs + (hrs === 1 ? " hour ago" : " hours ago");
    return fmtDate(new Date(ts));
  }

  /* How much this device is still holding that the Mac has not seen. */
  function pendingCount() {
    if (window.WPLDesktop) return 0;          // the Mac writes into the shared store directly
    return syncCfg.dirtyCount || 0;
  }

  /* ---- app updates ----
     Android checks a public HTTPS release channel independently of Mac sync.
     The Mac-carried APK remains as a fallback for older/local-only setups. */
  var desktopVersion = "";            // filled in once by the Electron host

  function appVersion() {
    if (window.WPLBridge && window.WPLBridge.versionName) {
      try { return "v" + window.WPLBridge.versionName(); } catch (err) { /* fall through */ }
    }
    return desktopVersion ? "v" + desktopVersion : "";
  }

  function myVersionCode() {
    if (window.WPLBridge && window.WPLBridge.versionCode) {
      try { return window.WPLBridge.versionCode(); } catch (err) { return 0; }
    }
    return 0;
  }

  function renderUpdate() {
    var card = $("updateCard");
    if (!card) return;
    if (isIOSApp()) {
      card.classList.add("hidden");
      availableUpdateVersion = "";
      var storeBuild = iosDistribution() === "appstore";
      setSum("sumAppUpdates", storeBuild ? "App Store" : "Personal build");
      var toggle = $("autoUpdateToggle");
      if (toggle && toggle.closest(".toggle-row")) toggle.closest(".toggle-row").classList.add("hidden");
      if ($("updateSettingHelp")) {
        $("updateSettingHelp").textContent = storeBuild
          ? "Your iPhone installs new versions automatically when App Updates is on in iPhone Settings."
          : "This app uses Apple's free personal signature and must be rebuilt and reinstalled every seven days.";
      }
      if ($("btnCheckUpdate")) {
        $("btnCheckUpdate").classList.add("hidden");
      }
      if ($("updateCheckStatus")) {
        $("updateCheckStatus").textContent = storeBuild
          ? "App updates do not remove your saved work logs."
          : "Automatic updates are not available for free-signed IPA files.";
      }
      return;
    }
    var mine = myVersionCode();
    var onlineApk = null;
    if (window.WPLBridge && window.WPLBridge.onlineUpdateInfo) {
      try {
        var raw = window.WPLBridge.onlineUpdateInfo();
        onlineApk = raw ? JSON.parse(raw) : null;
      } catch (err) { onlineApk = null; }
    }
    var candidate = onlineApk && onlineApk.versionCode > mine ? onlineApk : null;
    if (hostApk && hostApk.versionCode > mine &&
        (!candidate || hostApk.versionCode > candidate.versionCode)) candidate = hostApk;
    var fromInternet = candidate === onlineApk;
    var offered = candidate && candidate.versionCode;
    var macOffer = window.WPLDesktop && nativeUpdateInfo && nativeUpdateInfo.available
      ? nativeUpdateInfo : null;
    var show = !!((window.WPLBridge && offered && mine && offered > mine) || macOffer);
    var shownName = macOffer ? macOffer.versionName : candidate && (candidate.versionName || offered);
    availableUpdateVersion = show ? "v" + shownName + " available" : "";
    card.classList.toggle("hidden", !show);
    var autoOn = automaticUpdatesEnabled();
    if ($("autoUpdateToggle")) $("autoUpdateToggle").checked = autoOn;
    setSum("sumAppUpdates", updateChecking ? "Checking…" : autoOn ? "Automatic" : "Manual");
    if ($("btnCheckUpdate")) $("btnCheckUpdate").disabled = updateChecking;
    if ($("btnCheckUpdate")) $("btnCheckUpdate").textContent = updateChecking ? "Checking…" : "Check now";
    if ($("updateCheckStatus")) {
      var status = nativeUpdateInfo && nativeUpdateInfo.error ? nativeUpdateInfo.error
        : updateChecking ? "Looking for a newer version…"
        : show ? "A newer version is ready"
        : autoOn ? "Checks happen quietly in the background" : "Automatic checks are off";
      $("updateCheckStatus").textContent = status;
    }
    if (!show) return;
    var name = String(shownName || offered);
    var ready = macOffer ? !!macOffer.ready : !!(fromInternet && candidate.ready);
    if (!fromInternet) {
      try { if (!macOffer) ready = !!(window.WPLBridge.isUpdateReady && window.WPLBridge.isUpdateReady(name)); }
      catch (err) { ready = false; }
    }

    /* Start pulling it down as soon as we know about it, so that by the time the
       button is pressed there is usually nothing left to wait for. */
    if (!macOffer && !fromInternet && !ready && window.WPLBridge.prefetchUpdate && syncCfg.host && syncCfg.code) {
      try {
        window.WPLBridge.prefetchUpdate(hostUrl(syncCfg.host) + "/apk", syncCfg.code, name);
      } catch (err) { /* it gets fetched on demand instead */ }
    }

    $("updateVersion").textContent = "v" + name;
    $("btnUpdate").textContent = ready ? "Install now" : "Update now";
    $("updateText").textContent = macOffer
      ? "Version " + name + " is available. " + (ready
        ? "The verified Mac installer is ready to open. " : "It is downloading now. ") +
        "Your logged days will stay."
      : "Version " + name + " is available. This phone has " +
        (window.WPLBridge.versionName ? window.WPLBridge.versionName() : mine) + ". " +
        (ready ? "Ready to install. " : "Downloading now. ") +
        "Your logged days will stay.";
    runFriendlyReminders();
  }

  /* ---------- sync UI ---------- */
  function renderSync() {
    renderUpdate();
    var isHost = !!window.WPLDesktop;
    var secureWeb = !isHost && !window.WPLBridge && location.protocol === "https:";
    var cloudCoverage = cloudSyncCoverage();
    $("syncHostCard").classList.toggle("hidden", !isHost);
    /* An HTTPS Home Screen app cannot call the Mac's plain HTTP address.
       Cloud sync is the reliable route here, so do not offer a control that
       the browser will always block as mixed content. */
    $("syncClientCard").classList.toggle("hidden", isHost || secureWeb);

    if (isHost) {
      window.WPLDesktop.syncInfo().then(function (info) {
        $("syncToggle").checked = !!info.enabled;
        $("syncCode").textContent = info.code || "------";
        if (info.rotate) $("syncRotate").value = info.rotate;
        $("syncRotateHint").textContent = info.nextRotation
          ? "The code changes on " + fmtDateLong(new Date(info.nextRotation)) + "."
          : "This code will not change.";
        $("syncAddresses").innerHTML = (info.addresses || []).length
          ? info.addresses.map(function (a) {
              return '<code class="addr">' + esc(a) + "</code>";
            }).join("")
          : '<span class="muted tiny">Not on a network right now</span>';
        var live = info.listeners > 0 ? " · phone connected live" : "";
        $("syncHostState").textContent = info.lastPeer && !info.error
          ? "Last direct sync with " + info.lastPeer + " " + fmtAgo(info.lastSyncAt) + live
          : cloudCoverage.synced
            ? "Your data is synced through the cloud"
            : info.error ? info.error
            : info.enabled ? "Waiting for your phone to pair" : "Phone sync is off";
        $("syncHostState").classList.toggle("is-pending", !!info.error && !cloudCoverage.synced);
        setSum("sumSyncHost", info.listeners > 0 ? "Phone connected"
          : info.lastPeer ? "Paired"
          : cloudCoverage.synced ? "Using cloud"
          : !info.enabled ? "Off" : "Waiting to pair");
      });
    } else if (!secureWeb) {
      $("syncUseMac").checked = !!syncCfg.enabled;
      $("syncHost").value = syncCfg.host || "";
      $("syncCode2").value = syncCfg.code || "";
      $("syncAuto").checked = !!syncCfg.auto;
      $("syncClientControls").classList.toggle("sync-disabled", !syncCfg.enabled);
      $("syncLast").textContent = syncCfg.lastAt
        ? "Last synced " + fmtAgo(syncCfg.lastAt) + (syncCfg.lastPeer ? " with " + syncCfg.lastPeer : "")
        : cloudCoverage.synced ? "Your data is synced through the cloud"
        : syncCfg.enabled ? "Not synced yet" : "Mac sync is off";

      var pending = pendingCount();
      var el = $("syncPending");
      el.textContent = cloudCoverage.synced
        ? "Your latest changes are saved to the cloud"
        : pending
        ? pending + (pending === 1 ? " change" : " changes") + " waiting for your Mac"
        : syncCfg.enabled ? "Phone and Mac are synced" : "Direct Mac sync is optional";
      el.classList.toggle("is-pending", pending > 0 && !cloudCoverage.synced);

      setSum("sumSyncClient", !syncCfg.enabled ? (cloudCoverage.synced ? "Using cloud" : "Off")
        : !syncCfg.host ? (cloudCoverage.synced ? "Using cloud" : "Not paired")
        : pending && !cloudCoverage.synced ? pending + " waiting"
        : syncCfg.lastAt ? "Synced " + fmtAgo(syncCfg.lastAt)
        : cloudCoverage.synced ? "Using cloud" : "Paired");
    }
  }

  /* ============================================================
     VERSION HISTORY VIEW
     ============================================================ */
  var historyFocusDate = "";     // set when opened from a single day

  function describeSnapshot(s) {
    if (!s) return "nothing";
    if (s.deleted) return "deleted";
    var kind = s.leave ? leaveInfo(s.leave).label : "";
    if (!s.start || !s.finish) return kind ? kind + " (unpaid)" : "day off";
    return (kind ? kind + " " : "") + s.start + "–" + s.finish;
  }

  function historyTitle(r) {
    if (r.kind === "settings") return "Settings";
    var d = parseDate(r.date);
    return d ? dayName(d) + " " + fmtDate(d) : r.date;
  }

  function settingsDiff(before, after) {
    var labels = {
      rate: "Hourly rate", breakHours: "Unpaid break", defaultStart: "Default start",
      defaultFinish: "Default finish", firstPayday: "First payday", cycleDays: "Cycle length",
      appColor: "App colour"
    };
    var out = [];
    Object.keys(labels).forEach(function (k) {
      var b = before ? before[k] : undefined;
      var a = after ? after[k] : undefined;
      if (String(b) !== String(a)) out.push(labels[k] + " " + b + " → " + a);
    });
    return out.length ? out.join(", ") : "no change";
  }

  function renderHistory() {
    var q = ($("hSearch").value || "").trim().toLowerCase();
    var kind = $("hFilter").value;

    var list = (state.history || []).slice().reverse().filter(function (r) {
      if (historyFocusDate && r.date !== historyFocusDate) return false;
      if (kind !== "all" && r.kind !== kind) return false;
      if (!q) return true;
      var hay = (historyTitle(r) + " " + (r.after && r.after.notes || "") + " " +
        (r.before && r.before.notes || "") + " " + r.device).toLowerCase();
      return hay.indexOf(q) !== -1;
    });

    $("hNote").textContent = historyFocusDate
      ? "Changes to " + fmtDate(parseDate(historyFocusDate)) + " only."
      : "Newest changes first.";
    $("hEmpty").classList.toggle("hidden", list.length > 0);

    $("hList").innerHTML = list.map(function (r) {
      var detail = r.kind === "settings"
        ? esc(settingsDiff(r.before, r.after))
        : esc(describeSnapshot(r.before) + " → " + describeSnapshot(r.after));
      var note = r.kind === "day" && r.after && r.after.notes
        ? '<div class="h-note">“' + esc(r.after.notes) + '”</div>' : "";
      var badge = r.action === "delete" ? "deleted"
        : r.action === "add" ? "added"
        : r.action === "undo" ? "undone"
        : r.action === "settings" ? "settings"
        : "edited";
      return '<div class="wcard h-card">' +
        '<div class="wcard-top"><div>' +
        '<div class="wcard-title">' + esc(historyTitle(r)) + '</div>' +
        '<div class="wcard-sub">' + esc(r.device) + " · " + fmtAgo(r.at) + '</div>' +
        '</div><span class="pill">' + badge + '</span></div>' +
        '<div class="h-change">' + detail + '</div>' + note +
        '<div class="btn-row"><button class="btn h-undo" data-id="' + r.id + '">Undo this change</button></div>' +
        '</div>';
    }).join("");

    [].forEach.call($("hList").querySelectorAll(".h-undo"), function (b) {
      b.onclick = function () { undoRecord(b.dataset.id); };
    });
  }

  function undoRecord(id) {
    var r = (state.history || []).filter(function (x) { return x.id === id; })[0];
    if (!r) return;

    var what = r.kind === "settings"
      ? "Settings will go back to: " + settingsDiff(r.after, r.before) + "."
      : fmtDate(parseDate(r.date)) + " will go back to " + describeSnapshot(r.before) + ".";

    confirmAsk("Undo this change?", what,
      function () {
        var now = Date.now();
        if (r.kind === "settings") {
          var beforeSettings = Object.assign({}, state.settings);
          state.settings = Object.assign({}, WPL_SEED.settings, r.before || {});
          state.settingsUpdatedAt = now;
          recordHistory("settings", "", beforeSettings, Object.assign({}, state.settings), "undo");
        } else {
          var e = state.entries.filter(function (x) {
            return r.entryId ? x.id === r.entryId : x.date === r.date;
          })[0];
          var was = snapshotOf(e);
          var target = r.before;
          if (!target) {                              /* undoing an add removes the day */
            if (e) { e.deleted = true; e.updatedAt = now; }
          } else if (e) {
            e.start = target.start; e.finish = target.finish;
            e.notes = target.notes; e.leave = target.leave || "";
            e.employerId = target.employerId || state.settings.defaultEmployerId;
            e.rateSnapshot = Math.max(0, Number(target.rateSnapshot) || 0);
            e.breakHours = Math.max(0, Number(target.breakHours) || 0);
            e.breaks = Array.isArray(target.breaks) ? target.breaks.map(function(item){return {id:item.id,duration:Number(item.duration)||0,paid:!!item.paid};}) : undefined;
            e.status = ["planned", "completed", "approved", "paid"].indexOf(target.status) !== -1 ? target.status : "completed";
            e.payMultiplier=target.payMultiplier==null?1:Math.max(0,Number(target.payMultiplier)||0);e.rateType=target.rateType||"standard";
            e.deleted = !!target.deleted; e.updatedAt = now;
          } else {
            state.entries.push({
              id: target.id || r.entryId || uid(), date: target.date || r.date, start: target.start, finish: target.finish,
              notes: target.notes, leave: target.leave || "",
              employerId: target.employerId || state.settings.defaultEmployerId,
              rateSnapshot: Math.max(0, Number(target.rateSnapshot) || 0),
              breakHours: Math.max(0, Number(target.breakHours) || 0),
              breaks: Array.isArray(target.breaks) ? target.breaks.map(function(item){return {id:item.id,duration:Number(item.duration)||0,paid:!!item.paid};}) : undefined,
              status: ["planned", "completed", "approved", "paid"].indexOf(target.status) !== -1 ? target.status : "completed",
              payMultiplier:target.payMultiplier==null?1:Math.max(0,Number(target.payMultiplier)||0),rateType:target.rateType||"standard",
              updatedAt: now, deleted: !!target.deleted
            });
          }
          recordHistory("day", r.date, was,
            target || { start: "", finish: "", notes: "", deleted: true }, "undo");
        }
        save();
        render();
        toast("Change undone");
      });
  }

  /* ============================================================
     THEME
     ============================================================ */
  function applyTheme() {
    var mode = localStorage.getItem(THEME_KEY) || "auto";
    var dark = mode === "dark" ||
      (mode === "auto" && window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches);
    document.documentElement.classList.toggle("dark", dark);
    document.body.classList.toggle("dark", dark);
    applyAppColor();
    scheduleWidgetRefresh();
  }

  function applyAppColor() {
    var selected = APP_COLORS[state.settings.appColor] ? state.settings.appColor : "orange";
    var colour = APP_COLORS[selected];
    var dark = document.documentElement.classList.contains("dark");
    var values = dark ? colour.dark : colour.light;
    var names = ["--accent","--accent-2","--accent-soft","--bg","--bg-2","--surface","--surface-2","--line","--line-soft","--text","--text-2","--muted"];
    names.forEach(function (name, index) {
      document.documentElement.style.setProperty(name, values[index]);
      /* body also carries the .dark class, whose defaults otherwise override
         inherited values from the document root. Set both so the live colour
         choice wins in automatic, light and dark mode on every platform. */
      if (document.body) document.body.style.setProperty(name, values[index]);
    });
    document.documentElement.dataset.appColor = selected;
    if (document.body) document.body.dataset.appColor = selected;
    var meta = document.querySelector('meta[name="theme-color"]');
    if (meta) meta.setAttribute("content", values[3]);

    /* Keep the Android system bars and WebView canvas in the selected colour.
       Older phone builds still understand the boolean theme fallback. */
    if (window.WPLBridge) {
      try {
        if (window.WPLBridge.setNativeColours) window.WPLBridge.setNativeColours(dark, values[3], values[5]);
        else if (window.WPLBridge.setNativeTheme) window.WPLBridge.setNativeTheme(dark);
      } catch (err) { /* older shell */ }
    }
  }

  /* ============================================================
     WIRING
     ============================================================ */
  function bind() {
    [].forEach.call(document.querySelectorAll(".nav-item,.tab"), function (b) {
      b.onclick = function () { showView(b.dataset.view); };
    });
    [].forEach.call(document.querySelectorAll("[data-goto]"), function (b) {
      b.onclick = function () { showView(b.dataset.goto); };
    });

    $("thCountry").onchange = function () {
      var cfg = state.settings.takeHome, preset = TAKE_HOME_PRESETS[this.value] || TAKE_HOME_PRESETS.CUSTOM;
      cfg.country = this.value; cfg.profile = preset.profiles[0][0];
      if (cfg.country === "MT" && !cfg.social) cfg.social = "mt-prorata";
      state.settingsUpdatedAt = Date.now(); save(); renderTakeHome();
    };
    $("thProfile").onchange = function () {
      state.settings.takeHome.profile = this.value;
      if (this.value === "mt-stipend") state.settings.takeHome.social = "none";
      else if (this.value === "mt-student-under18") state.settings.takeHome.social = "mt-apprentice-under18";
      else if (this.value === "mt-student-18") state.settings.takeHome.social = "mt-apprentice18";
      else if (state.settings.takeHome.country === "MT" && state.settings.takeHome.social === "none") state.settings.takeHome.social = "mt-prorata";
      state.settingsUpdatedAt = Date.now(); save(); renderTakeHome();
    };
    $("thSocial").onchange = function () { state.settings.takeHome.social = this.value; state.settingsUpdatedAt = Date.now(); save(); renderTakeHome(); };
    $("thEstimateEmployer").onchange = function () { state.settings.takeHome.estimateEmployerId=this.value;state.settingsUpdatedAt=Date.now();save();renderTakeHome(); };
    [["thPayslipGross","gross"],["thPayslipTax","tax"],["thPayslipSocial","social"],["thPayslipOther","other"],["thPayslipNet","net"]].forEach(function(pair){
      $(pair[0]).onchange=function(){state.settings.takeHome.payslip[pair[1]]=Math.max(0,Number(this.value)||0);state.settingsUpdatedAt=Date.now();save();renderTakeHome();};
    });
    [["thOtherIncome","otherIncome"],["thWorkMonths","workMonths"],["thEstimatedHours","estimatedHours"],["thCustomTax","customTax"],["thCustomSocial","customSocial"],["thCustomOther","customOther"],["thCustomFixed","customFixed"]].forEach(function(pair){
      $(pair[0]).onchange = function () {
        var value = Math.max(0, Number(this.value) || 0);
        if (pair[1] === "workMonths") value = Math.max(1, Math.min(12, Math.round(value) || 1));
        else if (pair[1] !== "otherIncome" && pair[1] !== "customFixed" && pair[1] !== "estimatedHours") value = clampRate(value);
        state.settings.takeHome[pair[1]] = value; state.settingsUpdatedAt = Date.now(); save(); renderTakeHome();
      };
    });

    $("spIncomeMode").onchange = function () {
      state.settings.spendingPlan.incomeMode = this.value === "custom" ? "custom" : "takehome";
      saveSpending();
    };
    $("spCustomIncome").onchange = function () {
      state.settings.spendingPlan.customIncome = Math.max(0, Number(this.value) || 0);
      saveSpending();
    };
    $("spAddCategory").onclick = function () {
      var palette = ["#ec4899","#06b6d4","#84cc16","#ef4444","#6366f1"];
      var list = state.settings.spendingPlan.categories;
      list.push({ id: "category-" + uid(), name: "New category",
        color: palette[list.length % palette.length], mode: "percent", value: 0, enabled:true });
      saveSpending();
      var rows = $("spCategoryList").querySelectorAll(".sp-category");
      if (rows.length) { var input = rows[rows.length - 1].querySelector(".sp-name"); input.focus(); input.select(); }
    };
    $("spAddTransaction").onclick=function(){var date=$("spTxDate").value,categoryId=$("spTxCategory").value,amount=Math.max(0,Number($("spTxAmount").value)||0);if(!validStoredDate(date)){toast("Pick a valid date");return;}if(!categoryId){toast("Add or turn on a category first");return;}if(amount<=0){toast("Enter the amount spent");return;}state.settings.spendingPlan.transactions.push({id:"transaction-"+uid(),date:date,categoryId:categoryId,amount:xround(amount,2),note:($("spTxNote").value||"").trim().slice(0,60),updatedAt:Date.now()});$("spTxAmount").value="";$("spTxNote").value="";saveSpending();toast("Spending added");};
    $("spAddBill").onclick=function(){var name=($("spBillName").value||"").trim(),amount=Math.max(0,Number($("spBillAmount").value)||0),due=$("spBillDue").value;if(!name){toast("Name the bill");return;}if(amount<=0){toast("Enter the bill amount");return;}if(!validStoredDate(due)){toast("Pick the next due date");return;}state.settings.spendingPlan.bills.push({id:"bill-"+uid(),name:name.slice(0,40),amount:xround(amount,2),nextDue:due,categoryId:$("spBillCategory").value||"",enabled:true});$("spBillName").value="";$("spBillAmount").value="";saveSpending();toast("Bill added");};
    $("spAddGoal").onclick=function(){var name=($("spGoalName").value||"").trim(),target=Math.max(0,Number($("spGoalTarget").value)||0),saved=Math.max(0,Number($("spGoalSaved").value)||0),date=$("spGoalDate").value;if(!name){toast("Name the goal");return;}if(target<=0){toast("Enter the target amount");return;}if(date&&!validStoredDate(date)){toast("Pick a valid target date");return;}state.settings.spendingPlan.goals.push({id:"goal-"+uid(),name:name.slice(0,40),target:xround(target,2),saved:xround(saved,2),targetDate:date||""});$("spGoalName").value="";$("spGoalTarget").value="";$("spGoalSaved").value="";$("spGoalDate").value="";saveSpending();toast("Goal added");};

    $("fab").onclick = function () { openEditor(null); };
    $("topLogBtn").onclick = function () { openEditor(null); };
    $("sideLogBtn").onclick = function () { openEditor(null); };
    $("qLogToday").onclick = function () { openEditor(null); };
    $("qClock").onclick = toggleClock;

    $("sheetClose").onclick = closeEditor;
    $("scrim").onclick = closeEditor;
    $("edSave").onclick = saveEditor;
    $("edDelete").onclick = deleteEditor;
    ["edStart", "edFinish", "edRate", "edBreak", "edMultiplier"].forEach(function (id) {
      $(id).addEventListener("input", updatePreview);
      $(id).addEventListener("change", updatePreview);
    });
    $("edBreakPaid").onchange=updatePreview;
    $("edAddBreak").onclick=function(){editorExtraBreaks.push({duration:.25,paid:false});renderEditorExtraBreaks();updatePreview();};
    $("edRepeat").onchange=function(){$("edRepeatWeeksWrap").classList.toggle("hidden",this.value!=="weekly");};
    $("edRateType").onchange=function(){var rates={standard:1,overtime:1.5,night:1.25,sunday:2,public:2};if(this.value!=="custom")$("edMultiplier").value=String(rates[this.value]||1);$("edMultiplierWrap").classList.toggle("hidden",this.value!=="custom");updatePreview();};

    $("chipDefault").onclick = function () {
      var emp = employerFor({ employerId: $("edEmployer").value });
      $("edStart").value = emp.start || state.settings.defaultStart;
      $("edFinish").value = emp.finish || state.settings.defaultFinish;
      $("edRate").value = Number(emp.rate || 0);
      loadBreaksIntoEditor({employerId:emp.id,breakHours:Number(emp.breakHours||0)});
      updatePreview();
    };
    $("chipNow").onclick = function () {
      var n = new Date();
      $("edFinish").value = pad2(n.getHours()) + ":" + pad2(n.getMinutes());
      if (!$("edStart").value) $("edStart").value = employerFor({ employerId: $("edEmployer").value }).start || state.settings.defaultStart;
      updatePreview();
    };
    $("chipOff").onclick = function () {
      $("edStart").value = ""; $("edFinish").value = "";
      updatePreview();
    };
    function copyShiftFrom(e){$("edStart").value=e.start;$("edFinish").value=e.finish;$("edType").value=e.leave||"";$("edEmployer").value=e.employerId||state.settings.defaultEmployerId;$("edRate").value=typeof e.rateSnapshot==="number"?e.rateSnapshot:Number(employerFor(e).rate||0);loadBreaksIntoEditor(e);$("edStatus").value="completed";$("edRateType").value=e.rateType||"standard";$("edMultiplier").value=e.payMultiplier==null?1:e.payMultiplier;$("edMultiplierWrap").classList.toggle("hidden",$("edRateType").value!=="custom");updatePreview();}
    $("chipYesterday").onclick=function(){var d=parseDate($("edDate").value)||today(),iso=toISO(addDays(d,-1)),e=entries().filter(function(x){return x.date===iso;})[0];if(!e){toast("Yesterday has no saved shift");return;}copyShiftFrom(e);};
    $("chipLastWeek").onclick=function(){var d=parseDate($("edDate").value)||today(),iso=toISO(addDays(d,-7)),e=entries().filter(function(x){return x.date===iso;})[0];if(!e){toast("No shift was saved on this day last week");return;}copyShiftFrom(e);};
    $("edTemplate").onchange=function(){var id=this.value,t=(state.settings.shiftTemplates||[]).filter(function(x){return x.id===id;})[0];if(t){$("edStart").value=t.start;$("edFinish").value=t.finish;updatePreview();}};
    $("edEmployer").onchange=function(){var emp=employerFor({employerId:this.value});$("edRate").value=Number(emp.rate||0);loadBreaksIntoEditor({employerId:emp.id,breakHours:Number(emp.breakHours||0)});updatePreview();};

    /* Choosing a type fills in what that type normally pays: a full day for
       paid leave, nothing for unpaid. Adjust the times afterwards for a half
       day and the hours follow. */
    $("edType").onchange = function () {
      var info = leaveInfo(this.value);
      if (info.paid || !this.value) {
        var emp = employerFor({ employerId: $("edEmployer").value });
        if (!$("edStart").value) $("edStart").value = emp.start || state.settings.defaultStart;
        if (!$("edFinish").value) $("edFinish").value = emp.finish || state.settings.defaultFinish;
      } else {
        $("edStart").value = ""; $("edFinish").value = "";
      }
      updatePreview();
    };
    [].forEach.call(document.querySelectorAll(".note-chip"), function (c) {
      c.onclick = function () { $("edNotes").value = c.textContent; };
    });

    $("tsSearch").addEventListener("input", renderTimesheet);
    $("tsFilter").addEventListener("change", renderTimesheet);
    ["tsFrom","tsTo","tsEmployer","tsSort"].forEach(function(id){$(id).addEventListener("change",renderTimesheet);});
    $("tsListMode").onclick=function(){tsMode="list";this.classList.add("on");$("tsCalendarMode").classList.remove("on");renderTimesheet();};
    $("tsCalendarMode").onclick=function(){tsMode="calendar";if(!calendarAnchor)calendarAnchor=$("tsTo").value?parseDate($("tsTo").value):today();this.classList.add("on");$("tsListMode").classList.remove("on");renderTimesheet();};
    $("btnBulkClear").onclick=function(){tsSelected={};renderTimesheet();};
    $("btnSelectShown").onclick=function(){filteredTimesheet().forEach(function(e){tsSelected[e.id]=true;});renderTimesheet();};
    $("btnBulkShift").onclick=function(){var changed=[];state.entries.forEach(function(e){if(tsSelected[e.id]){changed.push({id:e.id,start:e.start,finish:e.finish,leave:e.leave});var emp=employerFor(e);e.start=emp.start;e.finish=emp.finish;e.leave="";e.updatedAt=Date.now();}});save();tsSelected={};render();toastUndo("Usual shift applied",function(){changed.forEach(function(x){var e=state.entries.filter(function(y){return y.id===x.id;})[0];if(e){e.start=x.start;e.finish=x.finish;e.leave=x.leave;e.updatedAt=Date.now();}});save();render();toast("Change undone");});};
    $("btnBulkLeave").onclick=function(){var value=window.prompt("Leave type: annual, public, sick, unpaid or off","annual");if(value===null)return;value=value.trim().toLowerCase();if(!LEAVE_TYPES.some(function(t){return t.key===value&&t.key;})){toast("That leave type was not recognised");return;}state.entries.forEach(function(e){if(tsSelected[e.id]){e.leave=value;if(!leaveInfo(value).paid){e.start="";e.finish="";}else if(!e.start&&!e.finish){var emp=employerFor(e);e.start=emp.start;e.finish=emp.finish;}e.updatedAt=Date.now();}});save();tsSelected={};render();toast("Leave applied");};
    $("btnBulkDelete").onclick=function(){var ids=Object.keys(tsSelected),copies=state.entries.filter(function(e){return tsSelected[e.id];}).map(function(e){return Object.assign({},e);});confirmAsk("Delete "+ids.length+" days?","You can undo this for a few seconds.",function(){var now=Date.now();state.entries.forEach(function(e){if(tsSelected[e.id]){e.deleted=true;e.updatedAt=now;}});tsSelected={};save();render();toastUndo("Days deleted",function(){copies.forEach(function(c){var e=state.entries.filter(function(x){return x.id===c.id;})[0];if(e)Object.assign(e,c,{updatedAt:Date.now(),deleted:false});});save();render();toast("Delete undone");});});};

    [].forEach.call($("anPresets").querySelectorAll("button"),function(b){b.onclick=function(){analyticsRange=b.dataset.range;[].forEach.call($("anPresets").querySelectorAll("button"),function(x){x.classList.toggle("on",x===b);});var custom=analyticsRange==="custom";$("anFromWrap").classList.toggle("hidden",!custom);$("anToWrap").classList.toggle("hidden",!custom);renderAnalytics();};});
    ["anFrom","anTo","anEmployer","anCompare"].forEach(function(id){$(id).addEventListener("change",renderAnalytics);});
    [].forEach.call(document.querySelectorAll(".table-sort"),function(b){b.onclick=function(){analyticsSort[b.dataset.table]*=-1;b.textContent=analyticsSort[b.dataset.table]<0?"Sort oldest":"Sort newest";if(analyticsModel)renderAnalyticsTables(analyticsModel.weeks,analyticsModel.periods,analyticsModel.months,analyticsModel.entries);else renderAnalytics();};});
    $("btnAnalyticsCsv").onclick=function(){var r=analyticsDates();saveFile("work-payment-log-analytics-"+r.fromISO+"-"+r.toISO+".csv",buildAnalyticsCsv(),"text/csv;charset=utf-8");};
    $("btnAnalyticsPdf").onclick=function(){var r=analyticsDates();saveFile("work-payment-log-analytics-"+r.fromISO+"-"+r.toISO+".pdf",buildAnalyticsPdf(),"application/pdf");};
    $("hSearch").addEventListener("input", renderHistory);
    $("hFilter").addEventListener("change", renderHistory);

    $("backBtn").onclick = function () { historyFocusDate = ""; showView("settings"); };
    $("btnHistory").onclick = function () { historyFocusDate = ""; showView("history"); };
    $("btnTour").onclick = function () { openOnboarding(true); };
    $("btnAddEmployer").onclick=function(){var id="job-"+uid();state.settings.employers.push({id:id,name:"New job",color:"#3b82f6",rate:Number(state.settings.rate||0),start:state.settings.defaultStart,finish:state.settings.defaultFinish,breakHours:Number(state.settings.breakHours||0)});state.settingsUpdatedAt=Date.now();save();renderSettings();$("sectEmployers").open=true;};
    $("btnAddTemplate").onclick=function(){state.settings.shiftTemplates.push({id:"shift-"+uid(),name:"New shift",start:state.settings.defaultStart,finish:state.settings.defaultFinish});state.settingsUpdatedAt=Date.now();save();renderSettings();$("sectTemplates").open=true;};
    $("stLogReminder").onchange=function(){
      state.settings.notifications.logHours=this.checked;state.settingsUpdatedAt=Date.now();save();renderNotifications();renderSectionSummaries();toast(this.checked?"Reminder set":"Reminder turned off");
    };
    $("stLogReminderTime").onchange=function(){
      if(!this.value)return;state.settings.logReminderTime=this.value;state.settingsUpdatedAt=Date.now();save();renderNotifications();toast("Reminder set for "+this.value);
    };
    [["stQuietStart","quietStart"],["stQuietEnd","quietEnd"]].forEach(function(pair){$(pair[0]).onchange=function(){if(!this.value)return;state.settings[pair[1]]=this.value;state.settingsUpdatedAt=Date.now();save();renderNotifications();toast("Quiet hours updated");};});
    [].forEach.call(document.querySelectorAll("[data-reminder-day]"),function(button){button.onclick=function(){
      var day=Number(button.dataset.reminderDay),days=(state.settings.logReminderDays||[]).slice(),at=days.indexOf(day);
      if(at===-1)days.push(day);else if(days.length>1)days.splice(at,1);else{toast("Choose at least one day");return;}
      state.settings.logReminderDays=days;state.settingsUpdatedAt=Date.now();save();renderNotifications();
    };});
    [].forEach.call(document.querySelectorAll("[data-notify]"),function(inp){inp.onchange=function(){state.settings.notifications[inp.dataset.notify]=inp.checked;state.settingsUpdatedAt=Date.now();save();renderSectionSummaries();toast(inp.checked?"Reminder enabled":"Reminder disabled");};});
    [].forEach.call(document.querySelectorAll("[data-app-color]"),function(button){button.onclick=function(){
      var next=button.dataset.appColor;
      if(!APP_COLORS[next]||state.settings.appColor===next)return;
      var before=Object.assign({},state.settings);
      state.settings.appColor=next;
      state.settingsUpdatedAt=Date.now();
      recordHistory("settings","",before,Object.assign({},state.settings),"settings");
      applyAppColor();
      save();
      renderSectionSummaries();
      toast("Colour changed to "+APP_COLORS[next].label);
    };});
    $("btnThemeSettings").onclick=function(){$("themeBtn").click();renderSectionSummaries();};

    $("stBackupEvery").onchange = function () {
      backupCfg.every = this.value;
      saveBackupCfg();
      renderBackup();
      toast(this.value === "off" ? "Automatic backups off"
        : "Backing up " + this.selectedOptions[0].textContent.toLowerCase());
      backupIfDue();
    };
    $("stBackupKeep").onchange = function () {
      backupCfg.keep = Math.max(1, Math.min(60, parseInt(this.value, 10) || 10));
      this.value = backupCfg.keep;
      saveBackupCfg();
    };
    $("btnBackupNow").onclick = function () { runBackup(true); };
    $("btnOpenBackups").onclick = function () {
      if (window.WPLDesktop && window.WPLDesktop.openBackups) window.WPLDesktop.openBackups();
    };
    $("edHistory").onclick = function () {
      historyFocusDate = $("edDate").value;
      closeEditor();
      showView("history");
    };

    /* settings inputs */
    var map = {
      stFirstPayday: ["firstPayday", "str"], stCycle: ["cycleDays", "int"],
      stCurrency: ["currency", "str"], stWeek: ["weekHours", "num"],
      stSymbol: ["currencySymbol", "sym"]
    };
    Object.keys(map).forEach(function (id) {
      $(id).addEventListener("change", function () {
        var key = map[id][0], kind = map[id][1], v = $(id).value;
        if (kind === "num") v = Math.max(0, parseFloat(v) || 0);
        else if (kind === "int") v = Math.max(1, parseInt(v, 10) || 1);
        else if (kind === "sym") v = (v || "").trim().slice(0, 4);
        else if (!v) return;
        var beforeSettings = Object.assign({}, state.settings);
        state.settings[key] = v;
        state.settingsUpdatedAt = Date.now();
        recordHistory("settings", "", beforeSettings,
          Object.assign({}, state.settings), "settings");
        save(); renderSettings();
        toast("Settings updated");
      });
    });

    $("btnExportCsv").onclick = function () {
      saveFile("work-payment-log.csv", buildCsv(), "text/csv;charset=utf-8");
    };
    $("btnExportJson").onclick = function () {
      saveFile("work-payment-log-backup.json",
        JSON.stringify({ app: "Work Payment Log", saved: new Date().toISOString(), settings: state.settings,
          settingsUpdatedAt: state.settingsUpdatedAt || 0, settingsFieldUpdatedAt: state.settingsFieldUpdatedAt || {},
          entries: state.entries, history: state.history || [] }, null, 2),
        "application/json");
    };
    $("btnImportJson").onclick = function () {
      if (window.WPLDesktop && window.WPLDesktop.openFile) {
        window.WPLDesktop.openFile().then(function (text) { if (text) applyImport(text); });
      } else {
        $("fileInput").click();
      }
    };
    $("fileInput").onchange = function () {
      var f = this.files && this.files[0];
      if (!f) return;
      var r = new FileReader();
      r.onload = function () { applyImport(String(r.result)); };
      r.readAsText(f);
      this.value = "";
    };

    $("btnWipe").onclick = function () {
      confirmAsk("Delete everything?", "All logged days will be removed. This cannot be undone.",
        function () {
          var now = Date.now();
          state.entries.forEach(function (e) { e.deleted = true; e.updatedAt = now; });
          save(); render(); toast("All days deleted");
        });
    };

    /* ---- sync ---- */
    if (window.WPLDesktop) {
      $("syncToggle").onchange = function () {
        window.WPLDesktop.setSync($("syncToggle").checked).then(renderSync);
      };
      $("btnMacRefresh").onclick = function () {
        window.WPLDesktop.pull().then(function (merged) {
          var changed = adoptMerged(merged, null);
          toast(changed ? "Picked up changes from your phone" : "Already up to date");
          renderSync();
        });
      };
      $("syncRotate").onchange = function () {
        window.WPLDesktop.setRotate($("syncRotate").value).then(function () {
          renderSync();
          toast($("syncRotate").value === "never"
            ? "The code will stay the same"
            : "The code will change " + $("syncRotate").selectedOptions[0].textContent.toLowerCase());
        });
      };
      $("btnNewCode").onclick = function () {
        confirmAsk("New pairing code?",
          "Your phone will need the new code before it can sync again.",
          function () { window.WPLDesktop.newCode().then(renderSync); });
      };
      if (window.WPLDesktop.onAuthCallback) {
        window.WPLDesktop.onAuthCallback(function (url) { window.WPLAuthCallback(url); });
      }
      if (window.WPLDesktop.onRemote) {
        window.WPLDesktop.onRemote(function (merged, peer) {
          adoptMerged(merged, peer || "your phone");
          renderSync();
        });
      }
    } else {
      $("syncUseMac").onchange = function () {
        syncCfg.enabled = this.checked;
        saveSyncCfg();
        if (!syncCfg.enabled) {
          closeStream();
          syncFailed = false;
          setSyncStatus("");
        } else if (syncCfg.auto) {
          syncNow(false);
        }
        renderSync();
        toast(syncCfg.enabled ? "Mac sync is on" : "Mac sync is off");
      };
      $("syncHost").addEventListener("change", function () {
        syncCfg.host = this.value.trim(); saveSyncCfg();
      });
      $("syncCode2").addEventListener("change", function () {
        syncCfg.code = this.value.replace(/\D/g, "").slice(0, 6); this.value = syncCfg.code; saveSyncCfg();
      });
      $("syncAuto").onchange = function () {
        syncCfg.auto = this.checked; saveSyncCfg();
        if (syncCfg.enabled && syncCfg.auto) openStream(); else closeStream();
        if (!syncCfg.auto) cloudSync(false);
        toast(syncCfg.auto ? "Will sync in the background" : "Background sync off");
      };
      $("btnSyncNow").onclick = function () { syncNow(true); };
      $("btnSyncFind").onclick = discoverHost;
    }

    $("autoUpdateToggle").onchange = function () {
      var on = this.checked;
      if (window.WPLDesktop && window.WPLDesktop.setAutoUpdates) {
        window.WPLDesktop.setAutoUpdates(on).then(function (info) {
          nativeUpdateInfo = info || nativeUpdateInfo;
          renderUpdate();
          toast(on ? "Automatic update checks are on" : "Automatic update checks are off");
        });
      } else if (window.WPLBridge && window.WPLBridge.setAutoUpdatesEnabled) {
        try { window.WPLBridge.setAutoUpdatesEnabled(on); } catch (err) { /* retry on next open */ }
        renderUpdate();
        toast(on ? "Automatic update checks are on" : "Automatic update checks are off");
      }
    };
    $("btnCheckUpdate").onclick = function () { checkForUpdate(true); };
    $("btnUpdate").onclick = function () {
      if (window.WPLDesktop && window.WPLDesktop.installUpdate) {
        window.WPLDesktop.installUpdate().then(function (result) {
          if (result && result.error) toast(result.error);
          else if (result && result.opened) toast("The Mac installer is open");
        });
        return;
      }
      if (!window.WPLBridge) return;
      var online = "";
      try { online = window.WPLBridge.onlineUpdateInfo && window.WPLBridge.onlineUpdateInfo(); }
      catch (err) { online = ""; }
      if (online && window.WPLBridge.installOnlineUpdate) {
        window.WPLBridge.installOnlineUpdate();
        return;
      }
      if (!hostApk || !window.WPLBridge.installUpdate) return;
      window.WPLBridge.installUpdate(hostUrl(syncCfg.host) + "/apk", syncCfg.code,
        hostApk.versionName || String(hostApk.versionCode));
    };

    $("confirmNo").onclick = function () { $("confirmWrap").classList.add("hidden"); confirmCb = null; };
    $("confirmYes").onclick = function () {
      $("confirmWrap").classList.add("hidden");
      if (confirmCb) confirmCb();
      confirmCb = null;
    };

    $("themeBtn").onclick = function () {
      var order = ["auto", "light", "dark"];
      var cur = localStorage.getItem(THEME_KEY) || "auto";
      var next = order[(order.indexOf(cur) + 1) % 3];
      localStorage.setItem(THEME_KEY, next);
      applyTheme();
      toast((next === "auto" ? "Automatic" : next.charAt(0).toUpperCase() + next.slice(1)) + " theme");
    };

    document.addEventListener("keydown", function (ev) {
      if (ev.key === "Escape") {
        if (!$("confirmWrap").classList.contains("hidden")) { $("confirmWrap").classList.add("hidden"); confirmCb = null; }
        else if (!$("sheet").classList.contains("hidden")) closeEditor();
      }
      if (ev.key === "Enter" && !$("sheet").classList.contains("hidden") &&
        document.activeElement && document.activeElement.tagName === "INPUT") saveEditor();
      if ((ev.metaKey || ev.ctrlKey) && ev.key.toLowerCase() === "n") { ev.preventDefault(); openEditor(null); }
    });

    if (window.matchMedia) {
      var mq = window.matchMedia("(prefers-color-scheme: dark)");
      if (mq.addEventListener) mq.addEventListener("change", applyTheme);
    }
  }

  /* One refresh covers every peer the phone has. Keep the Wi-Fi and cloud
     passes sequential: anything received from the Mac is then included in the
     cloud push, instead of two merges racing one another. */
  var refreshing = false;
  function syncAllNow() {
    if (refreshing) return Promise.resolve(false);
    /* Pull-to-sync is also the most natural explicit update check. Native code
       performs the network work; this just asks it to bypass the next timer. */
    try {
      if (window.WPLBridge && window.WPLBridge.checkForOnlineUpdate) {
        window.WPLBridge.checkForOnlineUpdate();
      }
    } catch (err) { /* the background worker remains the safety net */ }
    var hasMac = !!(syncCfg.enabled && syncCfg.host && syncCfg.code);
    var hasCloud = !!(window.WPLCloud && WPLCloud.signedIn());
    if (!hasMac && !hasCloud) {
      toast("Connect your Mac or sign in to cloud sync first");
      return Promise.resolve(false);
    }

    refreshing = true;
    var macOk = false, cloudOk = false;
    var work = hasMac
      ? syncNow(false, true).then(function (ok) { macOk = ok; })
      : Promise.resolve();

    return work.then(function () {
      return hasCloud ? cloudSync(false).then(function (ok) { cloudOk = ok; }) : null;
    }).then(function () {
      var ok = macOk || cloudOk;
      if (ok) toast(hasMac && hasCloud && macOk && cloudOk ? "Mac and cloud synced" : "Sync complete");
      else toast("Could not sync right now");
      return ok;
    }).catch(function () {
      toast("Could not sync right now");
      return false;
    }).then(function (ok) {
      refreshing = false;
      renderSync();
      renderCloud();
      return ok;
    });
  }

  /* Pull down at the top of any list to sync — the gesture people already try. */
  function bindPullToRefresh() {
    if (!window.WPLBridge) return;
    var sc = $("scroll"), ptr = $("ptr");
    var startX = 0, startY = 0, pulling = false, dist = 0, armed = false;
    var TRIGGER = 52;

    function resetPull() {
      pulling = false;
      armed = false;
      ptr.style.height = "0px";
      ptr.style.setProperty("--pull-progress", "0");
      ptr.classList.remove("is-armed", "is-pulling");
      $("ptrText").textContent = "Pull down to sync";
    }

    sc.addEventListener("touchstart", function (ev) {
      if (sc.scrollTop > 1 || syncing || refreshing || ev.touches.length !== 1) return;
      startX = ev.touches[0].clientX;
      startY = ev.touches[0].clientY;
      pulling = true;
      dist = 0;
      armed = false;
      ptr.classList.add("is-pulling");
    }, { passive: true });

    sc.addEventListener("touchmove", function (ev) {
      if (!pulling || ev.touches.length !== 1) return;
      var dx = ev.touches[0].clientX - startX;
      dist = ev.touches[0].clientY - startY;
      /* A sideways gesture belongs to a chart/table, not refresh. */
      if (Math.abs(dx) > Math.abs(dist) && Math.abs(dx) > 10) { resetPull(); return; }
      if (dist <= 0) { resetPull(); return; }
      ev.preventDefault();
      var shown = Math.min(62, dist * 0.66);           // short, forgiving pull
      var progress = Math.min(1, dist / TRIGGER);
      ptr.style.height = shown + "px";
      ptr.style.setProperty("--pull-progress", String(progress));
      var nowArmed = dist >= TRIGGER;
      ptr.classList.toggle("is-armed", nowArmed);
      if (nowArmed) $("ptrText").textContent = "Release to sync";
      else $("ptrText").textContent = progress > .45 ? "Keep pulling…" : "Pull down to sync";
      armed = nowArmed;
    }, { passive: false });

    function end() {
      if (!pulling) return;
      pulling = false;
      var go = armed;
      armed = false;
      ptr.classList.remove("is-armed");
      ptr.classList.remove("is-pulling");
      if (!go) { resetPull(); return; }
      ptr.style.height = "42px";
      ptr.classList.add("is-syncing");
      $("ptrText").textContent = "Syncing…";
      syncAllNow().then(function (ok) {
        $("ptrText").textContent = ok ? "Synced" : "Sync unavailable";
        setTimeout(function () {
          ptr.style.height = "0px";
          ptr.classList.remove("is-syncing");
          ptr.style.setProperty("--pull-progress", "0");
          $("ptrText").textContent = "Pull down to sync";
        }, 550);
      });
    }
    sc.addEventListener("touchend", end, { passive: true });
    sc.addEventListener("touchcancel", function () {
      resetPull();
    }, { passive: true });
  }

  /* Android WebView turns overflow-y:auto into a two-axis scroller if any
     child is even a few pixels too wide. Keep the app on the vertical axis so
     diagonal swipes cannot slide the whole interface sideways. */
  function bindVerticalScrollLock() {
    var sc = $("scroll");
    if (!sc) return;
    sc.addEventListener("scroll", function () {
      if (sc.scrollLeft) sc.scrollLeft = 0;
    }, { passive: true });
  }

  /* Horizontal swipes move between the six phone tabs. The direction is only
     claimed after it is clearly more horizontal than vertical, so normal page
     scrolling and pull-to-sync remain natural. Reports that scroll sideways
     keep their own gesture. */
  function bindPageSwipes() {
    var sc = $("scroll");
    if (!sc) return;
    var startX = 0, startY = 0, lastX = 0, startedAt = 0;
    var axis = "", tracking = false, blocked = false, activeView = null;

    function phoneLayout() {
      return window.matchMedia && window.matchMedia("(max-width: 820px)").matches;
    }
    function overlayOpen() {
      return !$('onboard').classList.contains('hidden') ||
        !$('sheet').classList.contains('hidden') ||
        !$('confirmWrap').classList.contains('hidden');
    }
    function ownsHorizontalGesture(target) {
      return !!(target && target.closest && target.closest(
        ".table-scroll,.range-presets,input,select,textarea,[contenteditable=true]"
      ));
    }
    function clearTracking(keepIndicator) {
      document.body.classList.remove("is-page-swiping");
      if (!keepIndicator) setPhoneTabPosition(PHONE_VIEWS.indexOf(currentView), false);
      tracking = false; axis = ""; activeView = null;
    }

    sc.addEventListener("touchstart", function (ev) {
      blocked = !phoneLayout() || overlayOpen() || ev.touches.length !== 1 ||
        PHONE_VIEWS.indexOf(currentView) < 0 || ownsHorizontalGesture(ev.target);
      if (blocked) return;
      clearTracking();
      startX = lastX = ev.touches[0].clientX;
      startY = ev.touches[0].clientY;
      startedAt = Date.now();
      activeView = $("view-" + currentView);
    }, { passive: true });

    sc.addEventListener("touchmove", function (ev) {
      if (blocked || !activeView || ev.touches.length !== 1) return;
      var dx = ev.touches[0].clientX - startX;
      var dy = ev.touches[0].clientY - startY;
      lastX = ev.touches[0].clientX;
      if (!axis) {
        if (Math.max(Math.abs(dx), Math.abs(dy)) < 10) return;
        axis = Math.abs(dx) > Math.abs(dy) * 1.15 ? "x" : "y";
        if (axis === "y") { activeView = null; return; }
        tracking = true;
        document.body.classList.add("is-page-swiping");
      }
      if (axis !== "x") return;
      ev.preventDefault();
      var index = PHONE_VIEWS.indexOf(currentView);
      var indicatorPosition = Math.max(0, Math.min(PHONE_VIEWS.length - 1,
        index - dx / Math.max(1, sc.clientWidth)));
      setPhoneTabPosition(indicatorPosition, true);
      /* Keep the current page solid while deciding whether this is a swipe.
         Moving a large, scrollable WebView layer exposed the bare background
         and forced Android to reraster it, producing the visible flash. */
    }, { passive: false });

    function finishSwipe(cancelled) {
      if (blocked || !activeView || axis !== "x" || !tracking) { clearTracking(); return; }
      var dx = lastX - startX;
      var velocity = dx / Math.max(1, Date.now() - startedAt);
      var index = PHONE_VIEWS.indexOf(currentView);
      var nextIndex = dx < 0 ? index + 1 : index - 1;
      var farEnough = Math.abs(dx) >= Math.min(72, sc.clientWidth * .18);
      var quickEnough = Math.abs(dx) >= 28 && Math.abs(velocity) >= .35;
      if (cancelled || nextIndex < 0 || nextIndex >= PHONE_VIEWS.length || (!farEnough && !quickEnough)) {
        clearTracking(); return;
      }
      var target = PHONE_VIEWS[nextIndex];
      clearTracking(true);
      showView(target);
    }
    sc.addEventListener("touchend", function () { finishSwipe(false); }, { passive: true });
    sc.addEventListener("touchcancel", function () { finishSwipe(true); }, { passive: true });
  }

  /* The FAB floats above the list, so it can end up covering a figure. Tuck it
     away while scrolling down and bring it straight back on the way up. */
  function bindFabAutoHide() {
    var fab = $("fab"), sc = $("scroll");
    if (!fab || !sc) return;
    var last = 0, idle = null;
    sc.addEventListener("scroll", function () {
      var y = sc.scrollTop;
      if (y > last + 6 && y > 90) fab.classList.add("is-tucked");
      else if (y < last - 6) fab.classList.remove("is-tucked");
      last = y;
      clearTimeout(idle);
      idle = setTimeout(function () { fab.classList.remove("is-tucked"); }, 900);
    }, { passive: true });
  }

  /* ============================================================
     AUTOMATIC BACKUPS
     Device-local settings (each device backs itself up to its own disk), so
     they live outside the synced store. A backup is the whole store as JSON —
     the same file "Restore backup" already reads.
     ============================================================ */
  var BACKUP_KEY = "wpl.backup";
  var BACKUP_EVERY = {
    off: 0,
    daily: 24 * 3600e3,
    weekly: 7 * 24 * 3600e3,
    monthly: 30 * 24 * 3600e3
  };

  function loadBackupCfg() {
    try {
      var c = JSON.parse(localStorage.getItem(BACKUP_KEY) || "null");
      if (c) return Object.assign({ every: "daily", keep: 10, lastAt: 0, lastPath: "" }, c);
    } catch (err) { /* defaults */ }
    return { every: "daily", keep: 10, lastAt: 0, lastPath: "" };
  }
  function saveBackupCfg() {
    try { localStorage.setItem(BACKUP_KEY, JSON.stringify(backupCfg)); } catch (err) { /* ignore */ }
  }
  var backupCfg = loadBackupCfg();

  function backupText() {
    var data={
      app: "Work Payment Log",
      saved: new Date().toISOString(),
      settings: state.settings,
      settingsUpdatedAt: state.settingsUpdatedAt || 0,
      settingsFieldUpdatedAt: state.settingsFieldUpdatedAt || {},
      entries: state.entries,
      history: state.history || []
    };
    data.integrity=backupIntegrity(data);
    return JSON.stringify(data, null, 2);
  }

  function backupIntegrity(data){var text=JSON.stringify({settings:data.settings||{},settingsUpdatedAt:data.settingsUpdatedAt||0,settingsFieldUpdatedAt:data.settingsFieldUpdatedAt||{},entries:data.entries||[],history:data.history||[]}),hash=2166136261;for(var i=0;i<text.length;i++){hash^=text.charCodeAt(i);hash=Math.imul(hash,16777619);}return "fnv1a-"+(hash>>>0).toString(16).padStart(8,"0");}

  function backupName() {
    return "work-payment-log-backup-" + toISO(today()) + ".json";
  }

  function canBackup() {
    return !!((window.WPLDesktop && window.WPLDesktop.autoBackup) ||
              (window.WPLBridge && window.WPLBridge.autoBackup));
  }

  /* Returns a promise of the path written, or "" if it could not be. */
  function writeBackup() {
    var name = backupName(), text = backupText(), keep = backupCfg.keep || 10;
    try{var check=JSON.parse(text),problem=validateBackup(check);if(problem){backupCfg.lastError=problem;saveBackupCfg();return Promise.resolve("");}}catch(err){backupCfg.lastError="The backup could not be checked";saveBackupCfg();return Promise.resolve("");}
    if (window.WPLDesktop && window.WPLDesktop.autoBackup) {
      return window.WPLDesktop.autoBackup(name, text, keep);
    }
    if (window.WPLBridge && window.WPLBridge.autoBackup) {
      try { return Promise.resolve(window.WPLBridge.autoBackup(name, text, keep)); }
      catch (err) { return Promise.resolve(""); }
    }
    return Promise.resolve("");
  }

  function runBackup(manual) {
    if (!canBackup()) {
      if (manual) toast("Backups need the installed app");
      return;
    }
    return writeBackup().then(function (path) {
      if (path) {
        backupCfg.lastAt = Date.now();
        backupCfg.lastPath = path;
        backupCfg.lastError="";
        saveBackupCfg();
        if (manual) toast("Backed up");
      } else if (manual) {
        toast(backupCfg.lastError||"Could not write the backup");
      }
      renderBackup();
    });
  }

  /* Nothing to back up on a brand-new install, and never more than once a
     period — the check is cheap enough to run at every launch. */
  function backupIfDue() {
    if (!backupCfg) return;                 // save() can fire before this is set up
    var span = BACKUP_EVERY[backupCfg.every] || 0;
    if (!span || !canBackup()) return;
    if (!entries().length) return;
    if (Date.now() - (backupCfg.lastAt || 0) < span) return;
    runBackup(false);
  }

  function renderBackup() {
    var sel = $("stBackupEvery");
    if (!sel) return;
    sel.value = backupCfg.every;
    $("stBackupKeep").value = backupCfg.keep;
    $("btnOpenBackups").classList.toggle("hidden", !(window.WPLDesktop && window.WPLDesktop.openBackups));

    var pill = $("backupPill");
    if (pill) {
      pill.textContent = !canBackup() ? "unavailable"
        : backupCfg.every === "off" ? "off" : backupCfg.every;
      pill.className = "pill" + (canBackup() && backupCfg.every !== "off" ? " pill-good" : "");
    }

    var where = window.WPLDesktop ? "Documents / Work Payment Log Backups"
              : "Downloads / Work Payment Log";
    $("backupState").textContent = backupCfg.lastError?backupCfg.lastError:!canBackup()
      ? "Automatic backups work in the installed app."
      : backupCfg.lastAt
        ? "Last backup " + fmtAgo(backupCfg.lastAt) + " · " + (backupCfg.lastPath || where)
        : backupCfg.every === "off"
          ? "Automatic backups are off."
          : "No backup yet. The first one is saved after you log a day.";
  }

  /* ============================================================
     CLOUD ACCOUNT (optional)
     Uses exactly the same merge rules as Wi-Fi sync, so the cloud is just
     another peer: pull, merge, push. Nothing here runs unless an account
     has been set up.
     ============================================================ */
  var cloudBusy = false;
  var cloudQueued = false;
  var cloudLastError = "";

  /* Direct Wi-Fi sync and cloud sync are two routes to the same data. Keep a
     signature of the exact store last uploaded so the UI can treat the cloud
     as a successful peer instead of claiming the Mac or phone is still owed
     changes. Older installs without a signature keep using their last
     successful cloud time until the next local edit or cloud pass. */
  function currentCloudSyncSignature() {
    try { return WPLSync.signature(storeForSync()); }
    catch (err) { return ""; }
  }

  function markCloudNeedsSync() {
    try { localStorage.setItem(CLOUD_SYNC_SIGNATURE_KEY, ""); }
    catch (err) { /* the cloud pass can still run */ }
  }

  function markCloudSynced() {
    var signature = currentCloudSyncSignature();
    if (!signature) return;
    try { localStorage.setItem(CLOUD_SYNC_SIGNATURE_KEY, signature); }
    catch (err) { /* lastAt remains the fallback */ }
  }

  function cloudSyncCoverage() {
    if (!window.WPLCloud || !WPLCloud.signedIn()) return { active: false, synced: false, lastAt: 0 };
    var cfg = WPLCloud.cfg ? WPLCloud.cfg() : {};
    var stored = null;
    try { stored = localStorage.getItem(CLOUD_SYNC_SIGNATURE_KEY); }
    catch (err) { stored = null; }
    var synced = !!cfg.lastAt && (stored === null || (!!stored && stored === currentCloudSyncSignature()));
    return { active: true, synced: synced, lastAt: Number(cfg.lastAt) || 0 };
  }

  /* An edit should reach the other device without waiting for the ten-minute
     tick. Editing a time field fires save() on every keystroke, so this waits
     for the typing to stop rather than syncing per character. */
  var cloudSoonTimer = null;
  function cloudSoon() {
    if (!window.WPLCloud || !WPLCloud.signedIn()) return;
    if (cloudSoonTimer) clearTimeout(cloudSoonTimer);
    cloudSoonTimer = setTimeout(function () {
      cloudSoonTimer = null;
      cloudSync(false);
    }, 2000);
  }

  function flushCloudSoon() {
    if (!cloudSoonTimer) return;
    clearTimeout(cloudSoonTimer);
    cloudSoonTimer = null;
    cloudSync(false);
  }

  function cloudSync(manual) {
    if (!window.WPLCloud || !WPLCloud.signedIn()) {
      if (manual && window.WPLCloud && !WPLCloud.signedIn()) toast("Sign in first");
      return Promise.resolve(false);
    }
    if (typeof navigator !== "undefined" && navigator.onLine === false) {
      cloudLastError = "Waiting for internet";
      if (manual) toast(cloudLastError);
      renderCloud();
      return Promise.resolve(false);
    }
    /* If an edit lands while a cloud request is running, do another pass when
       it finishes. Without this, that edit can wait for the long safety poll
       and make automatic sync look as if it needs the button. */
    if (cloudBusy) {
      cloudQueued = true;
      return Promise.resolve(false);
    }
    cloudBusy = true;
    if (manual) $("cloudState").textContent = "Syncing…";

    return WPLCloud.pull()
      .then(function (remote) {
        if (remote) adoptMerged(remote, manual ? "the cloud" : null);
        return WPLCloud.push(storeForSync());
      })
      .then(function () {
        markCloudSynced();
        cloudLastError = "";
        /* A successful cloud upload covers the same changes that direct
           Mac/phone sync would carry, so it is no longer an attention item. */
        syncFailed = false;
        setSyncStatus("");
        if (manual) toast("Cloud sync done");
        renderCloud();
        renderSync();
        renderDataHealth();
        return true;
      })
      .catch(function (err) {
        var msg = (err && err.message) || "Cloud sync failed";
        cloudLastError = /signed out|token/i.test(msg) ? "Sign in again" :
          (typeof navigator !== "undefined" && navigator.onLine === false ? "Waiting for internet" : msg);
        if (manual) toast(msg);
        $("cloudState").textContent = cloudLastError;
        return false;
      })
      .then(function (ok) {
        cloudBusy = false;
        if (cloudQueued) {
          cloudQueued = false;
          setTimeout(function () { cloudSync(false); }, 0);
        }
        return ok;
      });
  }

  function renderCloud() {
    if (!window.WPLCloud || !$("cloudCard")) return;
    var c = WPLCloud.cfg();
    var setUp = WPLCloud.configured(), inAcct = WPLCloud.signedIn();

    /* The server is built in, so this is no longer a setup step — it stays a
       folded escape hatch for anyone running their own. */
    $("cloudSetup").classList.remove("hidden");
    $("cloudAuth").classList.toggle("hidden", !setUp || inAcct);
    $("cloudAccount").classList.toggle("hidden", !inAcct);

    $("cloudPill").textContent = inAcct ? "automatic" : setUp ? "sign in" : "off";
    $("cloudPill").className = "pill" + (inAcct ? " pill-good" : "");

    $("cloudHelp").innerHTML = WPLCloud.usingBuiltIn()
      ? "Only change this if you run your own Supabase project. The table is in <b>supabase/schema.sql</b>."
      : "Using your own project. Clear both fields to use the built-in server again.";

    if (inAcct) {
      $("cloudWho").textContent = c.email || "your account";
      $("cloudState").textContent = cloudLastError || (c.lastAt
        ? "Saved to cloud " + fmtAgo(c.lastAt)
        : "Waiting for the first cloud sync");
    } else if (setUp) {
      $("cloudUrl").value = c.url;
      $("cloudKey").value = c.key;
    }
  }

  function bindCloud() {
    if (!window.WPLCloud || !$("cloudCard")) return;

    /* Bind through this rather than touching elements directly: a button that
       has been renamed or removed then costs nothing, instead of throwing and
       silently leaving every handler below it unbound. */
    function on(id, handler) {
      var el = $(id);
      if (el) el.onclick = handler;
      else if (window.console) console.warn("no element", id);
    }

    function runCheck(hostId, url, key) {
      var host = $(hostId);
      host.classList.remove("hidden");
      host.innerHTML = '<div class="check-row"><span class="mark">·</span><b>Checking…</b></div>';
      return WPLCloud.test(url, key).then(function (rows) {
        host.innerHTML = rows.map(function (r) {
          return '<div class="check-row ' + (r.ok ? "ok" : "bad") + '">' +
            '<span class="mark">' + (r.ok ? "✓" : "✕") + "</span>" +
            "<div><b>" + esc(r.step) + "</b><span>" + esc(r.note) + "</span></div></div>";
        }).join("");
        return rows;
      });
    }

    on("btnCloudCheck", function () {
      var c = WPLCloud.cfg();
      runCheck("cloudCheck2", c.url, c.key);
    });

    on("btnCloudTest", function () {
      runCheck("cloudCheck", $("cloudUrl").value, $("cloudKey").value).then(function (rows) {
        if (rows.every(function (r) { return r.ok; })) {
          toast("Everything checks out. Save the project and sign in.");
        }
      });
    });

    on("btnCloudSaveProject", function () {
      var url = $("cloudUrl").value.trim(), key = $("cloudKey").value.trim();
      if (!/^https?:\/\//.test(url) || key.length < 20) {
        toast("Paste the project URL and the anon key");
        return;
      }
      WPLCloud.setProject(url, key);
      renderCloud();
      toast("Project saved. Now sign in with Google.");
    });

    /* Google will not run its sign-in page inside an embedded browser, so the
       app hands off to the real one and waits for the worklog:// link back. */
    on("btnCloudGoogle", function () {
      var note = $("cloudAuthNote");
      note.textContent = "Checking…";
      WPLCloud.googleReady().then(function (ready) {
        if (!ready) {
          note.textContent = "Google sign-in is not switched on for this project yet. " +
            "Turn it on under Authentication → Sign In / Providers, then try again.";
          return;
        }
        note.textContent = "Finish signing in in your browser, then come back here.";
        var url = WPLCloud.googleUrl();
        if (window.WPLBridge && window.WPLBridge.openExternal) {
          window.WPLBridge.openExternal(url);
        } else if (window.WPLDesktop && window.WPLDesktop.openExternal) {
          window.WPLDesktop.openExternal(url);
        } else if (/^https?:$/.test(location.protocol)) {
          /* Stay in the installed web app during OAuth. Opening a separate
             Safari tab gives it a different storage context on iPhone. */
          window.location.assign(url);
        } else {
          window.open(url, "_blank");
        }
      });
    });

    on("btnCloudSignOut", function () {
      confirmAsk("Sign out?", "Your days stay on this device. The cloud copy is kept.",
        function () {
          WPLCloud.signOut();
          renderCloud();
          renderSync();
          renderDataHealth();
          toast("Signed out");
        });
    });

    on("btnCloudSync", function () { cloudSync(true); });
    on("btnCloudCheckBackup", function () {
      $("cloudState").textContent = "Checking cloud copy…";
      WPLCloud.pull().then(function(remote){
        if (!remote) { $("cloudState").textContent = "No cloud copy yet"; return; }
        var problem = validateBackup({settings:remote.settings,entries:remote.entries||[]});
        $("cloudState").textContent = problem || ("Cloud copy is healthy · " + (remote.entries || []).filter(function(e){return !e.deleted;}).length + " shifts");
      }).catch(function(err){$("cloudState").textContent=(err&&err.message)||"Could not check the cloud copy";});
    });
    on("btnCloudDownload", function () {
      WPLCloud.pull().then(function(remote){
        if (!remote) { toast("No cloud copy to download"); return; }
        var problem = validateBackup({settings:remote.settings,entries:remote.entries||[]});
        if (problem) { toast(problem); return; }
        saveFile("work-payment-log-cloud-copy.json", JSON.stringify(Object.assign({
          app:"Work Payment Log", saved:new Date().toISOString()
        }, remote), null, 2), "application/json");
      }).catch(function(err){toast((err&&err.message)||"Could not download the cloud copy");});
    });
    on("btnCloudRestore", function () {
      WPLCloud.pull().then(function(remote){
        if (!remote) { toast("No cloud copy to restore"); return; }
        applyImport(JSON.stringify(Object.assign({app:"Work Payment Log",saved:new Date().toISOString()},remote)));
      }).catch(function(err){toast((err&&err.message)||"Could not read the cloud copy");});
    });
    on("btnCloudActivity", function () {
      historyFocusDate = "";
      if ($("hFilter")) $("hFilter").value = "all";
      showView("history");
    });
  }

  /* ============================================================
     FIRST-RUN SETUP
     Five short steps. Nothing here is irreversible — every answer is a
     setting you can change afterwards, and the tour can be reopened.
     ============================================================ */
  var obStep = 0;
  var OB_LAST = 4;

  function fillCurrencies(sel, value) {
    sel.innerHTML = WPL_SEED.currencies.map(function (c) {
      return c.code === "custom"
        ? '<option value="custom">Something else…</option>'
        : '<option value="' + c.code + '">' + c.code + "  " + c.symbol.trim() + "</option>";
    }).join("");
    sel.value = value || "EUR";
  }

  function openOnboarding(fromSettings) {
    obStep = fromSettings ? OB_LAST : 0;
    var st = state.settings;
    var primary = employerFor({ employerId: st.defaultEmployerId });
    fillCurrencies($("obCurrency"), st.currency);
    $("obRate").value = primary.rate || st.rate || "";
    $("obStart").value = primary.start || st.defaultStart;
    $("obFinish").value = primary.finish || st.defaultFinish;
    $("obBreak").value = Number(primary.breakHours == null ? st.breakHours : primary.breakHours);
    $("obWeek").value = st.weekHours || 40;
    $("obCycle").value = st.cycleDays || 28;
    $("obPayday").value = st.firstPayday || toISO(addDays(today(), 7));
    $("onboard").classList.remove("hidden");
    showObStep(obStep);
  }

  function showObStep(n) {
    obStep = n;
    [].forEach.call(document.querySelectorAll(".ob-step"), function (el) {
      el.classList.toggle("hidden", Number(el.dataset.step) !== n);
    });
    $("obDots").innerHTML = "01234".split("").map(function (_, i) {
      return '<i class="' + (i === n ? "on" : "") + '"></i>';
    }).join("");
    $("obBack").classList.toggle("hidden", n === 0);
    $("obNext").textContent = n === 0 ? "Get started" : n === OB_LAST ? "Start logging" : "Next";
    if (n === 2) obShiftHint();
  }

  function obShiftHint() {
    var sm = minutesOf($("obStart").value), fm = minutesOf($("obFinish").value);
    var brk = Math.max(0, parseFloat($("obBreak").value) || 0);
    var rate = Math.max(0, parseFloat($("obRate").value) || 0);
    if (sm === null || fm === null || fm <= sm) { $("obShiftHint").textContent = ""; return; }
    var net = Math.max(0, (fm - sm) / 60 - brk);
    $("obShiftHint").textContent = "That is " + hmPlain(net) + " paid a day" +
      (rate ? ", " + money(xround(net * rate, 2)) : "") + ".";
  }

  /* Each step writes its own answers, so backing out never loses them. */
  function obCommit(n) {
    var st = state.settings;
    var primary = employerFor({ employerId: st.defaultEmployerId });
    if (n === 1) {
      var rate = parseFloat($("obRate").value);
      if (!(rate > 0)) { toast("Enter what you are paid per hour"); return false; }
      st.currency = $("obCurrency").value;
      st.rate = rate;
      primary.rate = rate;
    } else if (n === 2) {
      if (!$("obStart").value || !$("obFinish").value) { toast("Set your usual start and finish"); return false; }
      if (minutesOf($("obFinish").value) <= minutesOf($("obStart").value)) {
        toast("The finish needs to be after the start"); return false;
      }
      st.defaultStart = $("obStart").value;
      st.defaultFinish = $("obFinish").value;
      st.breakHours = Math.max(0, parseFloat($("obBreak").value) || 0);
      st.weekHours = Math.max(1, parseFloat($("obWeek").value) || 40);
      primary.start = st.defaultStart;
      primary.finish = st.defaultFinish;
      primary.breakHours = st.breakHours;
    } else if (n === 3) {
      if (!$("obPayday").value) { toast("Pick your next payday"); return false; }
      st.firstPayday = $("obPayday").value;
      st.cycleDays = Math.max(1, parseInt($("obCycle").value, 10) || 28);
    }
    return true;
  }

  function bindOnboarding() {
    $("obNext").onclick = function () {
      if (!obCommit(obStep)) return;
      if (obStep === OB_LAST) {
        state.settings.onboarded = true;
        state.settingsUpdatedAt = Date.now();
        save();
        $("onboard").classList.add("hidden");
        render();
        toast("You are ready. Tap + to log your first day.");
        return;
      }
      showObStep(obStep + 1);
    };
    $("obBack").onclick = function () { if (obStep > 0) showObStep(obStep - 1); };
    ["obStart", "obFinish", "obBreak", "obRate"].forEach(function (id) {
      $(id).addEventListener("input", obShiftHint);
    });
  }

  /* Called by the Mac shell and by Android once the browser sends the result
     back to worklog://auth. */
  window.WPLAuthCallback = function (url) {
    if (!window.WPLCloud) return;
    var r = WPLCloud.completeOAuth(url);
    if (!r.ok) {
      toast(r.error);
      if ($("cloudAuthNote")) $("cloudAuthNote").textContent = r.error;
      return;
    }
    renderCloud();
    toast("Signed in as " + (r.email || "your account"));
    cloudSync(false);
  };

  /* ---------- leave pay rates ---------- */
  function renderLeaveRates() {
    var host = $("leaveRates");
    if (!host) return;
    host.innerHTML = LEAVE_TYPES.filter(function (t) { return t.key; }).map(function (t) {
      return '<div class="rate-row"><span>' + esc(t.label) + "</span>" +
        '<span><input type="number" min="0" max="100" step="5" inputmode="numeric" ' +
        'data-leave="' + t.key + '" value="' + leavePct(t.key) + '"><span class="pc">%</span></span></div>';
    }).join("");

    [].forEach.call(host.querySelectorAll("input"), function (inp) {
      inp.onchange = function () {
        var v = Math.max(0, Math.min(100, parseFloat(inp.value) || 0));
        inp.value = v;
        var before = Object.assign({}, state.settings);
        state.settings.leaveRates[inp.dataset.leave] = v;
        state.settingsUpdatedAt = Date.now();
        recordHistory("settings", "", before, Object.assign({}, state.settings), "settings");
        save(); render();
        toast(leaveInfo(inp.dataset.leave).label + " now pays " + v + "%");
      };
    });
  }

  /* ---------- boot ---------- */
  applyTheme();
  bind();
  bindOnboarding();
  bindCloud();
  bindPullToRefresh();
  bindVerticalScrollLock();
  bindPageSwipes();
  bindFabAutoHide();
  showView("overview");
  /* The iPhone Action Button shortcut opens this HTTPS deep link. Consume the
     flag immediately so refreshing or reopening Safari does not keep showing
     the editor, then open today's entry with the user's saved defaults. */
  var quickLogRequested = false;
  var quickLogLauncher = false;
  try {
    var launchParams = new URLSearchParams(location.search);
    var quickLogMode = launchParams.get("quicklog");
    var quickLogPath = /\/quicklog\/?$/.test(location.pathname);
    quickLogRequested = quickLogPath || quickLogMode === "1" || quickLogMode === "launcher";
    quickLogLauncher = quickLogPath || quickLogMode === "launcher";
    if (quickLogRequested && !quickLogLauncher) {
      launchParams.delete("quicklog");
      var cleanQuery = launchParams.toString();
      history.replaceState(null, document.title,
        location.pathname + (cleanQuery ? "?" + cleanQuery : "") + location.hash);
    }
  } catch (err) { /* old WebViews simply use the normal launch screen */ }
  /* Supabase returns browser sign-in tokens in the URL fragment. Consume them
     once, then remove them so they are never left in screenshots or history. */
  if (location.hash && /(?:^|[&#])(access_token|error)=/.test(location.hash)) {
    var authReturnUrl = location.href;
    try { history.replaceState(null, document.title, location.pathname + location.search); } catch (err) { /* ignore */ }
    window.WPLAuthCallback(authReturnUrl);
  }
  scheduleWidgetRefresh();
  if (migrationApplied) save();
  syncWorkReminderSchedule(true);
  if (!state.settings.onboarded) openOnboarding(false);
  else if (quickLogRequested) setTimeout(openTodayEditor, 80);
  backupIfDue();
  setInterval(backupIfDue, 3600e3);
  setInterval(function(){renderClockAction();},30000);
  setTimeout(runFriendlyReminders, 1200);

  /* Treat the cloud like an automatic peer. These listeners are installed even
     before sign-in, so an account connected later in the same app session gets
     the full background behaviour immediately. Wi-Fi phone sync is optional:
     cloud sync keeps running when it is off, unpaired, or unreachable. */
  cloudSync(false);
  setInterval(function () { cloudSync(false); }, 60 * 1000);
  document.addEventListener("visibilitychange", function () {
    /* Leaving: send the pending edit now rather than losing the timer when the
       phone suspends the app. Returning: pick up the other device. */
    if (document.hidden) flushCloudSoon();
    else {
      cloudSync(false);
      runFriendlyReminders();
      /* The separate Log Work Day Home Screen launcher may be resumed instead
         of reloaded by iOS. Reopen today's form whenever that launcher comes
         back to the foreground, including from the Action Button. */
      if (quickLogLauncher && state.settings.onboarded) {
        setTimeout(openTodayEditor, 80);
      }
    }
  });
  window.addEventListener("online", function () { cloudSync(false); });
  window.WPLOpenToday = openTodayEditor;
  window.WPLOpenSettings = function () { showView("settings"); };
  window.WPLOpenWidget = function (target) {
    if (target === "today") openTodayEditor();
    else if (target === "week") showView("analytics");
    else if (target === "payday") showView("takehome");
  };
  window.WPLRefreshWidgets = scheduleWidgetRefresh;

  /* Pick up anything the other device left while this one was closed, then keep
     checking quietly: on launch, whenever the app comes back to the foreground,
     and every few minutes while it is open. */
  if (window.WPLDesktop && window.WPLDesktop.appVersion) {
    window.WPLDesktop.appVersion().then(function (v) {
      desktopVersion = v || "";
      if (currentView === "settings") renderSettings();
    }).catch(function () { /* no version, no label */ });
    refreshNativeUpdateInfo();
    if (window.WPLDesktop.onUpdateChanged) {
      window.WPLDesktop.onUpdateChanged(function (info) {
        nativeUpdateInfo = info || null;
        updateChecking = !!(info && info.checking);
        renderUpdate();
      });
    }
  }

  if (window.WPLDesktop && window.WPLDesktop.pull) {
    window.WPLDesktop.pull().then(function (merged) {
      adoptMerged(merged, null);
      pushToHost();            // seeds the shared store on a fresh install
      renderSync();
    }).catch(function () { /* first run, nothing stored yet */ });
  } else if (window.WPLBridge) {
    var autoSync = function () { if (syncCfg.enabled && syncCfg.auto) syncNow(false); };
    /* Android calls this from Activity.onResume because some WebViews do not
       emit visibilitychange when the app returns from the background. */
    window.WPLResume = function (openUpdateSettings, openLogToday, widgetTarget) {
      cloudSync(false);
      autoSync();
      scheduleWidgetRefresh();
      try {
        if (window.WPLBridge.checkForOnlineUpdateAutomatically) {
          window.WPLBridge.checkForOnlineUpdateAutomatically();
        }
      } catch (err) { /* the scheduled updater will retry */ }
      if (openUpdateSettings) {
        showView("settings");
        setTimeout(function () {
          renderUpdate();
          var card = $("updateCard");
          if (card && !card.classList.contains("hidden")) {
            card.scrollIntoView({ block: "center", behavior: reduceMotion ? "auto" : "smooth" });
          }
          toast("Update paused. Install it here when you are ready.");
        }, 120);
      }
      if (openLogToday) setTimeout(openTodayEditor, 120);
      else if (widgetTarget) setTimeout(function () { window.WPLOpenWidget(widgetTarget); }, 120);
    };
    autoSync();
    document.addEventListener("visibilitychange", function () {
      if (document.hidden) closeStream();
      else { autoSync(); openStream(); }
    });
    window.addEventListener("online", autoSync);

    /* App releases are independent from phone-to-Mac sync. Native code also
       schedules background checks; these foreground checks keep the Settings
       card current without making anyone press a button. */
    var checkAppUpdate = function () {
      try {
        if (automaticUpdatesEnabled() && window.WPLBridge.checkForOnlineUpdateAutomatically) {
          window.WPLBridge.checkForOnlineUpdateAutomatically();
        }
      } catch (err) { /* WorkManager will try again */ }
    };
    checkAppUpdate();
    setInterval(checkAppUpdate, 15 * 60 * 1000);
    setInterval(renderUpdate, 15 * 1000);
    document.addEventListener("visibilitychange", function () {
      if (!document.hidden) checkAppUpdate();
    });

    /* While something is still waiting for the Mac, try every 30 seconds so it
       goes over the moment you are back on the same Wi-Fi. Otherwise idle down
       to every three minutes. */
    var ticks = 0;
    setInterval(function () {
      ticks++;
      if (!syncCfg.enabled || !syncCfg.auto) return;
      /* Retry briskly while something is owed or the last try failed — that is
         exactly when the Mac has just come back — and idle down otherwise. */
      if (pendingCount() > 0 || syncFailed || ticks % 6 === 0) syncNow(false);
    }, 30 * 1000);
  }

  /* Android hardware back: close what is open, then step back to Overview,
     and only then let the system leave the app. */
  window.WPLBack = function () {
    /* During first-run setup, back means the previous step — never "quit the
       app half way through being set up". */
    if (!$("onboard").classList.contains("hidden")) {
      if (obStep > 0) showObStep(obStep - 1);
      return true;
    }
    if (!$("confirmWrap").classList.contains("hidden")) {
      $("confirmWrap").classList.add("hidden"); confirmCb = null; return true;
    }
    if (!$("sheet").classList.contains("hidden")) { closeEditor(); return true; }
    if (currentView !== "overview") { showView("overview"); return true; }
    return false;
  };

  window.WPL = { state: state, hoursOf: hoursOf, payOf: payOf, weeks: weeks, periods: periods, totals: grandTotals };
})();
