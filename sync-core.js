/* ============================================================
   Work Payment Log — shared sync rules.

   Loaded three ways, and it must behave identically in all of them:
     • <script> in the web app        → window.WPLSync
     • require() in the Electron host → module.exports
     • the same file copied into both builds

   The merge is deliberately predictable: one record per shift id, the most
   recently edited copy wins, and deletes travel as tombstones so a shift
   removed on one device does not come back from the other. Multiple shifts
   on the same date therefore remain independent.
   ============================================================ */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.WPLSync = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var PORT = 45673;

  function stamp(e) { return typeof e.updatedAt === "number" ? e.updatedAt : 0; }

  function entryKey(e, index) {
    if (e && e.id) return String(e.id);
    /* Compatibility for very old backups. They only supported one shift per
       date; the index keeps malformed duplicate-date backups recoverable. */
    return "legacy:" + String(e && e.date || "") + ":" + String(index || 0);
  }

  /* One entry per id. Later edit wins; on a dead heat the first list wins,
     which keeps the result stable for the device initiating the merge. */
  function mergeEntries(mine, theirs) {
    var byId = {};
    function take(list) {
      (list || []).forEach(function (e, index) {
        if (!e || !e.date) return;
        var key = entryKey(e, index), cur = byId[key];
        if (!cur || stamp(e) > stamp(cur)) byId[key] = e;
      });
    }
    take(mine);
    take(theirs);
    return Object.keys(byId).map(function (k) { return byId[k]; }).sort(function (a, b) {
      return String(a.date).localeCompare(String(b.date)) ||
        String(a.start || "").localeCompare(String(b.start || "")) ||
        String(a.id || "").localeCompare(String(b.id || ""));
    });
  }

  var HISTORY_LIMIT = 300;

  /* History is append-only, so merging is just a union by record id, newest
     last, trimmed to a length that cannot grow unbounded on either device. */
  function mergeHistory(mine, theirs) {
    var byId = {};
    (mine || []).concat(theirs || []).forEach(function (r) {
      if (r && r.id) byId[r.id] = r;
    });
    return Object.keys(byId)
      .map(function (k) { return byId[k]; })
      .sort(function (a, b) { return (a.at || 0) - (b.at || 0); })
      .slice(-HISTORY_LIMIT);
  }

  function mergeStores(mine, theirs) {
    mine = mine || {};
    theirs = theirs || {};
    var mineAt = mine.settingsUpdatedAt || 0;
    var theirsAt = theirs.settingsUpdatedAt || 0;
    var mineSettings = mine.settings || {}, theirSettings = theirs.settings || {};
    var mineFields = mine.settingsFieldUpdatedAt && typeof mine.settingsFieldUpdatedAt === "object"
      ? mine.settingsFieldUpdatedAt : {};
    var theirFields = theirs.settingsFieldUpdatedAt && typeof theirs.settingsFieldUpdatedAt === "object"
      ? theirs.settingsFieldUpdatedAt : {};
    var settingKeys = {}, settings = {}, fieldTimes = {};
    Object.keys(mineSettings).forEach(function(k){settingKeys[k]=true;});
    Object.keys(theirSettings).forEach(function(k){settingKeys[k]=true;});
    Object.keys(settingKeys).forEach(function(key){
      var mineFieldAt = typeof mineFields[key] === "number" ? mineFields[key] : mineAt;
      var theirFieldAt = typeof theirFields[key] === "number" ? theirFields[key] : theirsAt;
      settings[key] = theirFieldAt > mineFieldAt ? theirSettings[key] : mineSettings[key];
      if (typeof settings[key] === "undefined") settings[key] = theirSettings[key];
      fieldTimes[key] = Math.max(mineFieldAt || 0, theirFieldAt || 0);
    });
    return {
      settings: settings,
      settingsUpdatedAt: Math.max(mineAt, theirsAt),
      settingsFieldUpdatedAt: fieldTimes,
      entries: mergeEntries(mine.entries, theirs.entries),
      history: mergeHistory(mine.history, theirs.history)
    };
  }

  /* Cheap change detector, so a sync that changed nothing stays silent. */
  function signature(store) {
    var s = String(store && store.settingsUpdatedAt || 0) + "|";
    (store && store.entries || []).forEach(function (e) {
      s += (e.id || e.date) + ":" + e.date + ":" +
        (e.deleted ? "x" : (e.start || "") + "-" + (e.finish || "")) +
        ":" + Number(e.rateSnapshot || 0) + ":" + Number(e.breakHours || 0) +
        ":" + JSON.stringify(e.breaks || []) +
        ":" + (e.status || "completed") +
        ":" + (e.payMultiplier == null ? 1 : Number(e.payMultiplier)) + ":" + (e.rateType || "standard") +
        ":" + stamp(e) + ";";
    });
    var h = (store && store.history) || [];
    s += "|h" + h.length + ":" + (h.length ? h[h.length - 1].id : "");
    return s;
  }

  /* What actually moves over the wire. */
  function payload(store, device) {
    return {
      app: "work-payment-log",
      protocol: 1,
      deviceId: device && device.id,
      deviceName: device && device.name,
      settings: store.settings,
      settingsUpdatedAt: store.settingsUpdatedAt || 0,
      settingsFieldUpdatedAt: store.settingsFieldUpdatedAt || {},
      entries: store.entries || [],
      history: store.history || []
    };
  }

  function looksValid(body) {
    return !!(body && body.app === "work-payment-log" && Array.isArray(body.entries));
  }

  function newCode() {
    var n = "";
    for (var i = 0; i < 6; i++) n += Math.floor(Math.random() * 10);
    return n;
  }

  return {
    PORT: PORT,
    HISTORY_LIMIT: HISTORY_LIMIT,
    mergeEntries: mergeEntries,
    mergeHistory: mergeHistory,
    mergeStores: mergeStores,
    signature: signature,
    payload: payload,
    looksValid: looksValid,
    newCode: newCode
  };
});
