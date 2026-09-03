/* ============================================================
   Optional cloud account (Supabase)

   Everything here is dormant until a project URL and key are entered in
   Settings. With no account configured the app behaves exactly as before:
   local storage plus Wi-Fi sync, nothing leaving the house.

   Deliberately plain fetch against Supabase's REST API rather than their
   SDK — the SDK does not load cleanly inside an Android WebView from a
   file:// page, and this needs about sixty lines.
   ============================================================ */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.WPLCloud = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var CFG_KEY = "wpl.cloud";
  var CONFIRM_PAGE = "/storage/v1/object/public/site/confirmed.html";

  /* The project this build talks to. Baked in on purpose: the people who use
     the app should never see a project URL or a key — they sign in, and that is
     all. The anon key is designed to be public and shipped inside clients; what
     actually protects the data is the row-level security policy on `stores`,
     which lets each account touch only its own row. */
  var BUILT_IN = {
    url: "https://whxkjfrdoaqsrgbqicvg.supabase.co",
    key: "sb_publishable_4DidnEdtfrbpZ8yFArhrcA_wjCNcfT6"
  };

  /* Projects that no longer exist. A saved project overrides the built-in one,
     so anyone who typed one of these in by hand stays pointed at a dead server
     through every future update until we clear it for them. */
  var RETIRED = ["pgguanevmwbauagudibp"];

  function loadCfg() {
    var c = null;
    try { c = JSON.parse(localStorage.getItem(CFG_KEY) || "null"); } catch (err) { /* fresh */ }
    c = c || { token: "", refresh: "", email: "", userId: "", lastAt: 0 };

    var saved = String(c.url || "");
    for (var i = 0; i < RETIRED.length; i++) {
      if (saved.indexOf(RETIRED[i]) !== -1) {
        c.url = ""; c.key = "";
        c.token = ""; c.refresh = ""; c.userId = ""; c.email = "";
        break;
      }
    }

    /* A saved project overrides the built-in one, so a self-hoster can point
       the same app at their own Supabase. Everyone else gets the built-in. */
    if (!c.url) c.url = BUILT_IN.url;
    if (!c.key) c.key = BUILT_IN.key;
    return c;
  }

  function saveCfg(cfg) {
    try { localStorage.setItem(CFG_KEY, JSON.stringify(cfg)); } catch (err) { /* ignore */ }
  }

  var cfg = loadCfg();

  function configured() { return !!(cfg.url && cfg.key); }
  function signedIn() { return !!(configured() && cfg.token && cfg.userId); }
  function base() { return String(cfg.url || "").replace(/\/+$/, ""); }

  function headers(withAuth) {
    var h = { "Content-Type": "application/json", "apikey": cfg.key };
    if (withAuth && cfg.token) h["Authorization"] = "Bearer " + cfg.token;
    return h;
  }

  function asError(res, body) {
    var msg = (body && (body.error_description || body.msg || body.message || body.error)) || "";
    if (/not confirmed/i.test(msg) || /email_not_confirmed/i.test(msg)) {
      return "That email has not been confirmed yet — see the note below";
    }
    if (res.status === 400 && /invalid login/i.test(msg)) return "That email and password do not match";
    if (res.status === 422 || /already registered/i.test(msg)) return "That email already has an account";
    if (res.status === 401) return "Signed out — sign in again";
    if (/rate limit/i.test(msg)) {
      return "Too many emails have gone out in the last hour — try again later";
    }
    if (/after (\d+) seconds?/i.test(msg)) {
      return "Just sent one — " + msg.match(/after (\d+) seconds?/i)[0].replace("after", "wait");
    }
    return msg || ("Request failed (" + res.status + ")");
  }

  function post(path, body, withAuth) {
    return fetch(base() + path, {
      method: "POST", headers: headers(withAuth), body: JSON.stringify(body)
    }).then(function (res) {
      return res.json().catch(function () { return {}; }).then(function (json) {
        if (!res.ok) throw new Error(asError(res, json));
        return json;
      });
    });
  }

  function adoptSession(json) {
    if (json.access_token) cfg.token = json.access_token;
    if (json.refresh_token) cfg.refresh = json.refresh_token;
    var user = json.user || {};
    if (user.id) cfg.userId = user.id;
    if (user.email) cfg.email = user.email;
    saveCfg(cfg);
  }

  /* Access tokens are short lived. Cloud sync can run for days without the
     settings page ever being opened, so refresh them quietly instead of making
     the user press Sync now (or sign in again) when a background request gets
     a 401. Keep one refresh in flight so a pull and push cannot race each
     other and rotate the refresh token twice. */
  var refreshPromise = null;
  function refresh() {
    if (refreshPromise) return refreshPromise;
    if (!cfg.refresh) return Promise.reject(new Error("Signed out — sign in again"));

    refreshPromise = post("/auth/v1/token?grant_type=refresh_token", {
      refresh_token: cfg.refresh
    }, false).then(function (json) {
      adoptSession(json);
      return true;
    }).catch(function (err) {
      /* Network failures are temporary; keep the session so the automatic
         retry can recover. Only discard credentials when Supabase says the
         refresh token itself is no longer usable. */
      var msg = (err && err.message) || "";
      if (/refresh token|invalid token|expired token/i.test(msg)) signOut();
      throw err;
    }).then(function (value) {
      refreshPromise = null;
      return value;
    }, function (err) {
      refreshPromise = null;
      throw err;
    });

    return refreshPromise;
  }

  /* ---------- account: Google ----------
     Google refuses OAuth inside embedded browsers, so the app never shows the
     sign-in page itself. It opens the system browser, and Supabase sends the
     result back to a worklog:// link that the app is registered to handle. */
  var CALLBACK = "worklog://auth";

  function callbackUrl() {
    /* Native apps return through their registered worklog:// URL. An installed
       web app must return to its own HTTPS page so the OAuth tokens land in
       the same browser storage as the timesheet. */
    try {
      if (typeof location !== "undefined" && /^https?:$/.test(location.protocol)) {
        return new URL("./", location.href).href;
      }
    } catch (err) { /* native or test environment */ }
    return CALLBACK;
  }

  function googleUrl() {
    return base() + "/auth/v1/authorize?provider=google&redirect_to=" +
      encodeURIComponent(callbackUrl());
  }

  /* The tokens arrive in the URL fragment. The user id is inside the token
     itself, so there is no extra round trip to make. */
  function claims(token) {
    try {
      var part = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
      while (part.length % 4) part += "=";
      return JSON.parse(decodeURIComponent(escape(atob(part))));
    } catch (err) {
      return null;
    }
  }

  function completeOAuth(url) {
    var hash = String(url || "");
    var i = hash.indexOf("#");
    if (i === -1) return { ok: false, error: "That sign-in did not return anything" };
    var q = {};
    hash.slice(i + 1).split("&").forEach(function (pair) {
      var j = pair.indexOf("=");
      if (j > 0) q[decodeURIComponent(pair.slice(0, j))] =
        decodeURIComponent(pair.slice(j + 1).replace(/\+/g, " "));
    });

    if (q.error) return { ok: false, error: q.error_description || q.error };
    if (!q.access_token) return { ok: false, error: "No sign-in token came back" };

    var c = claims(q.access_token);
    if (!c || !c.sub) return { ok: false, error: "That sign-in token could not be read" };

    cfg.token = q.access_token;
    cfg.refresh = q.refresh_token || "";
    cfg.userId = c.sub;
    cfg.email = c.email || "";
    saveCfg(cfg);
    return { ok: true, email: cfg.email };
  }

  function signOut() {
    cfg.token = ""; cfg.refresh = ""; cfg.userId = ""; cfg.email = ""; cfg.lastAt = 0;
    saveCfg(cfg);
  }

  /* ---------- the store, one row per user ---------- */
  function pull() {
    return fetch(base() + "/rest/v1/stores?select=data,updated_at", { headers: headers(true) })
      .then(function (res) {
        if (res.status === 401) return refresh().then(pull);
        return res.json().catch(function () { return []; }).then(function (rows) {
          if (!res.ok) throw new Error(asError(res, rows));
          return rows && rows.length ? rows[0].data : null;
        });
      });
  }

  function push(store) {
    var row = { user_id: cfg.userId, data: store, updated_at: new Date().toISOString() };
    return fetch(base() + "/rest/v1/stores", {
      method: "POST",
      headers: Object.assign(headers(true), { "Prefer": "resolution=merge-duplicates,return=minimal" }),
      body: JSON.stringify(row)
    }).then(function (res) {
      if (res.status === 401) return refresh().then(function () { return push(store); });
      if (!res.ok) {
        return res.json().catch(function () { return {}; }).then(function (j) {
          throw new Error(asError(res, j));
        });
      }
      cfg.lastAt = Date.now();
      saveCfg(cfg);
      return true;
    });
  }

  /* ---------- setup check ----------
     Runs each requirement separately so a failure names the actual problem
     instead of "something went wrong": the URL, the key, the table, then the
     row-level security policy. */
  function test(url, key) {
    var u = String(url || "").replace(/\/+$/, "");
    var k = String(key || "").trim();
    var out = [];

    if (!/^https:\/\/.+/.test(u)) {
      return Promise.resolve([{ ok: false, step: "Project URL",
        note: "Should start with https:// and end in .supabase.co" }]);
    }
    if (k.length < 30) {
      return Promise.resolve([{ ok: false, step: "Anon key", note: "That looks too short to be the key" }]);
    }
    if (/service_role/.test(k)) {
      return Promise.resolve([{ ok: false, step: "Anon key",
        note: "That is the service_role key — use the anon public one" }]);
    }

    var h = { "apikey": k, "Content-Type": "application/json" };
    var checks = [];

    return fetch(u + "/rest/v1/stores?select=user_id&limit=1", { headers: h })
      .then(function (res) {
        out.push({ ok: true, step: "Project URL", note: "Reached your project" });

        if (res.status === 401 || res.status === 403) {
          out.push({ ok: false, step: "Anon key", note: "The project did not accept that key" });
          return out;
        }
        out.push({ ok: true, step: "Anon key", note: "Accepted" });

        return res.json().catch(function () { return {}; }).then(function (body) {
          var msg = (body && (body.message || body.hint)) || "";
          if (res.status === 404 || /does not exist|Could not find the table/i.test(msg)) {
            out.push({ ok: false, step: "stores table",
              note: "Not there yet — run the SQL from the README in the SQL Editor" });
            return out;
          }
          out.push({ ok: true, step: "stores table", note: "Found" });

          checks.push(fetch(u + "/auth/v1/settings", { headers: h })
            .then(function (r) { return r.json(); })
            .then(function (st) {
              var on = st && st.external && st.external.google;
              out.push(on
                ? { ok: true, step: "Google sign-in", note: "Enabled on your project" }
                : { ok: false, step: "Google sign-in",
                    note: "Not enabled yet — turn Google on in Authentication → Sign In / " +
                          "Providers, and add worklog://auth under URL Configuration." });
            })
            .catch(function () { /* not fatal */ }));

          return Promise.all(checks).then(function () { return out; });
        });
      })
      .catch(function () {
        out.push({ ok: false, step: "Project URL", note: "Could not reach it — check the address" });
        return out;
      });
  }

  /* Asking the server first means a misconfigured project produces a sentence
     in the app rather than a page of raw JSON in the user's browser. */
  function googleReady() {
    return fetch(base() + "/auth/v1/settings", { headers: { "apikey": cfg.key } })
      .then(function (res) { return res.json(); })
      .then(function (s) { return !!(s && s.external && s.external.google); })
      .catch(function () { return false; });
  }

  return {
    cfg: function () { return cfg; },
    test: test,
    googleReady: googleReady,
    setProject: function (url, key) {
      cfg.url = String(url || "").trim() || BUILT_IN.url;
      cfg.key = String(key || "").trim() || BUILT_IN.key;
      saveCfg(cfg);
    },
    usingBuiltIn: function () { return cfg.url === BUILT_IN.url && cfg.key === BUILT_IN.key; },
    configured: configured,
    signedIn: signedIn,
    googleUrl: googleUrl,
    completeOAuth: completeOAuth,
    signOut: signOut,
    pull: pull,
    push: push
  };
});
