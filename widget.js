(function () {
  "use strict";
  var SUPABASE_URL = "https://whxkjfrdoaqsrgbqicvg.supabase.co";
  var ANON_KEY = "sb_publishable_4DidnEdtfrbpZ8yFArhrcA_wjCNcfT6";
  var token = new URLSearchParams(location.search).get("token") || "";
  var $ = function (id) { return document.getElementById(id); };

  function money(symbol, amount) {
    return (symbol || "€") + Number(amount || 0).toFixed(2);
  }
  function shortDate(raw) {
    if (!raw) return "Next payday not set";
    var parts = String(raw).split("-");
    var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    var month = months[Number(parts[1]) - 1];
    return month ? "Next payday · " + Number(parts[2]) + " " + month : "Next payday · " + raw;
  }
  function fail(message) {
    $("status").textContent = "Needs attention";
    $("message").textContent = message;
    $("message").hidden = false;
  }
  function render(data) {
    if (!data || !data.summary) return fail("Open Work Payment Log, then refresh Live iPhone widgets in Settings.");
    var summary = data.summary;
    var accent = /^#[0-9a-f]{6}$/i.test(data.accent || "") ? data.accent : "#ff9a4d";
    document.documentElement.style.setProperty("--accent", accent);
    if (data.dark === false) {
      document.documentElement.style.setProperty("--bg1", "#fffaf6");
      document.documentElement.style.setProperty("--bg2", "#ffeadd");
      document.documentElement.style.setProperty("--text", "#24150d");
      document.documentElement.style.setProperty("--muted", "#785f4d");
      document.documentElement.style.colorScheme = "light";
    }
    $("takeHome").textContent = money(data.currency, summary.takeHome);
    $("totalPay").textContent = money(data.currency, summary.totalPay);
    $("paydayGross").textContent = money(data.currency, summary.paydayGross);
    $("hours").textContent = Number(summary.totalHours || 0).toFixed(2);
    $("days").textContent = String(Number(summary.daysWorked || 0));
    $("payday").textContent = shortDate(summary.paydayDate);
    $("status").textContent = "Live";
    $("message").hidden = true;
  }
  function load() {
    if (!/^[a-f0-9]{64}$/i.test(token)) return fail("This widget link is incomplete. Copy it again from Work Payment Log Settings.");
    fetch(SUPABASE_URL + "/rest/v1/rpc/get_widget_snapshot", {
      method: "POST",
      headers: { "apikey": ANON_KEY, "Content-Type": "application/json" },
      body: JSON.stringify({ p_token: token }),
      cache: "no-store"
    }).then(function (response) {
      if (!response.ok) throw new Error("Update failed");
      return response.json();
    }).then(render).catch(function () { fail("Could not refresh. Check your internet connection and try again."); });
  }
  load();
  setInterval(load, 5 * 60 * 1000);
})();
