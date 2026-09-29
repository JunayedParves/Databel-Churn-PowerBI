/* Databel churn review: recomputes the Power BI report's measures in the browser. */
(function () {
  "use strict";
  var D = window.DATABEL, C = D.cols, L = D.levels, N = D.n;

  // ---- derived groups (same breaks as the DAX calculated columns) --------
  var AGE_GROUPS = ["Under 30", "30-39", "40-49", "50-64", "65+"];
  var TENURE_GROUPS = ["0-6", "7-12", "13-24", "25-48", "49+"];
  var CALL_GROUPS = ["0", "1", "2", "3", "4", "5+"];
  function ageGroup(a) { return a < 30 ? 0 : a < 40 ? 1 : a < 50 ? 2 : a < 65 ? 3 : 4; }
  function tenureGroup(m) { return m <= 6 ? 0 : m <= 12 ? 1 : m <= 24 ? 2 : m <= 48 ? 3 : 4; }
  function callGroup(c) { return Math.min(c, 5); }
  var ageIdx = C.age.map(ageGroup), tenIdx = C.tenure.map(tenureGroup), callIdx = C.calls.map(callGroup);
  function lv(field, name) { return L[field].indexOf(name); }

  // ---- formatting ----------------------------------------------------------
  function pct(x, d) { return x == null || isNaN(x) ? "–" : (x * 100).toFixed(d == null ? 1 : d) + "%"; }
  function int(x) { return x.toLocaleString("en-US"); }
  function money(x) { return "$" + Math.round(x).toLocaleString("en-US"); }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }

  // Colour rule from the report: orange = problem, blue = benchmark, grey = rest
  function churnColour(r, goodBelow) {
    if (r == null || isNaN(r)) return "var(--mid)";
    return r >= 0.35 ? "var(--hi)" : r < (goodBelow || 0.10) ? "var(--good)" : "var(--mid)";
  }

  // ---- filters ---------------------------------------------------------
  var F = { state: -1, age: -1, contract: -1, gender: -1 };
  function fillSelect(id, options, key) {
    var el = document.getElementById(id);
    el.innerHTML = '<option value="-1">All</option>' + options.map(function (o, i) {
      return '<option value="' + i + '">' + esc(o) + "</option>";
    }).join("");
    el.addEventListener("change", function () { F[key] = +el.value; render(); });
  }
  fillSelect("f-state", L.state, "state");
  fillSelect("f-age", AGE_GROUPS, "age");
  fillSelect("f-contract", L.contract, "contract");
  fillSelect("f-gender", L.gender, "gender");
  document.getElementById("f-reset").addEventListener("click", function () {
    F = { state: -1, age: -1, contract: -1, gender: -1 };
    ["f-state", "f-age", "f-contract", "f-gender"].forEach(function (id) { document.getElementById(id).value = "-1"; });
    render();
  });

  function selectedRows() {
    var out = [];
    for (var i = 0; i < N; i++) {
      if (F.state >= 0 && C.state[i] !== F.state) continue;
      if (F.age >= 0 && ageIdx[i] !== F.age) continue;
      if (F.contract >= 0 && C.contract[i] !== F.contract) continue;
      if (F.gender >= 0 && C.gender[i] !== F.gender) continue;
      out.push(i);
    }
    return out;
  }

  // group rows by key function -> [{key, n, churned, rate}]
  function groupBy(rows, keyFn, keys) {
    var m = {};
    rows.forEach(function (i) {
      var k = keyFn(i);
      if (k == null || k < 0) return;
      var g = m[k] || (m[k] = { key: k, n: 0, churned: 0 });
      g.n++; g.churned += C.churn[i];
    });
    var list = keys ? keys.map(function (k) { return m[k] || { key: k, n: 0, churned: 0 }; }) : Object.keys(m).map(function (k) { return m[k]; });
    list.forEach(function (g) { g.rate = g.n ? g.churned / g.n : null; });
    return list;
  }
  function range(n) { var a = []; for (var i = 0; i < n; i++) a.push(i); return a; }
  function stats(rows) {
    var n = rows.length, ch = 0, rev = 0, lost = 0;
    rows.forEach(function (i) { ch += C.churn[i]; rev += C.charge[i]; if (C.churn[i]) lost += C.charge[i]; });
    return { n: n, churned: ch, rate: n ? ch / n : null, revenue: rev, lost: lost };
  }
  function rateWhere(rows, pred) { return stats(rows.filter(pred)).rate; }
  function tip(title, g, extra) {
    return esc(title) + "|Churn rate " + pct(g.rate) + "|" + int(g.churned) + " of " + int(g.n) + " customers left" + (extra ? "|" + esc(extra) : "");
  }

  // ---- chart builders -------------------------------------------------------
  function hbars(id, items, opts) {
    opts = opts || {};
    var el = document.getElementById(id);
    if (!items.length) { el.innerHTML = '<p class="empty">No customers in this selection.</p>'; return; }
    var max = opts.max || Math.max.apply(null, items.map(function (d) { return d.value || 0; })) || 1;
    el.innerHTML = '<div class="hbars' + (opts.compact ? " compact" : "") + '" role="list">' + items.map(function (d) {
      var w = d.value == null ? 0 : Math.max(0.5, d.value / max * 100);
      return '<div class="hbar" role="listitem" tabindex="0" data-tip="' + d.tip + '" aria-label="' + esc(d.label + ": " + d.text) + '">' +
        '<span class="lbl" title="' + esc(d.label) + '">' + esc(d.label) + "</span>" +
        '<span class="track"><span class="fill" style="display:block;width:' + w + "%;background:" + d.colour + '"></span></span>' +
        '<span class="val">' + d.text + "</span></div>";
    }).join("") + "</div>";
  }

  function columns(id, items, opts) {
    opts = opts || {};
    var el = document.getElementById(id);
    var max = Math.max.apply(null, items.map(function (d) { return d.value || 0; })) || 1;
    el.innerHTML = '<div class="cols' + (opts.short ? " short" : "") + '" role="list">' + items.map(function (d) {
      var h = d.value == null ? 0 : Math.max(0.5, d.value / max * 88);
      return '<div class="col" role="listitem" tabindex="0" data-tip="' + d.tip + '" aria-label="' + esc(d.label + ": " + d.text) + '">' +
        '<span class="val">' + d.text + '</span><span class="bar" style="height:' + h + "%;background:" + d.colour + '"></span></div>';
    }).join("") + '</div><div class="col-labels" aria-hidden="true">' + items.map(function (d) {
      return "<span>" + esc(d.label) + (d.sub ? "<br>" + esc(d.sub) : "") + "</span>";
    }).join("") + "</div>";
  }

  // heat matrix, gradient from --heat-lo to --heat-hi over the visible range (like the Power BI rule)
  function hexToRgb(h) { h = h.trim().replace("#", ""); return [0, 2, 4].map(function (i) { return parseInt(h.substr(i, 2), 16); }); }
  function cssVar(name) { return getComputedStyle(document.documentElement).getPropertyValue(name); }
  function matrix(id, rowLabels, colLabels, cellFn, cls) {
    var lo = hexToRgb(cssVar("--heat-lo")), hi = hexToRgb(cssVar("--heat-hi"));
    var cells = rowLabels.map(function (_, r) { return colLabels.map(function (_, c) { return cellFn(r, c); }); });
    var vals = [].concat.apply([], cells).map(function (g) { return g.rate; }).filter(function (v) { return v != null; });
    var mn = Math.min.apply(null, vals), mx = Math.max.apply(null, vals);
    var html = '<div class="matrix ' + (cls || "") + '" style="grid-template-columns:minmax(80px,auto) repeat(' + colLabels.length + ',minmax(0,1fr))" role="table">';
    html += '<span role="columnheader"></span>' + colLabels.map(function (c) { return '<span class="hd" role="columnheader">' + esc(c) + "</span>"; }).join("");
    cells.forEach(function (row, r) {
      html += '<span class="rh" role="rowheader">' + esc(rowLabels[r]) + "</span>";
      row.forEach(function (g) {
        var t = g.rate == null ? 0 : (mx > mn ? (g.rate - mn) / (mx - mn) : 1);
        var rgb = lo.map(function (v, i) { return Math.round(v + (hi[i] - v) * t); });
        var ink = g.rate != null && g.rate >= 0.35 ? "#ffffff" : "var(--ink)";
        html += '<span class="cell" role="cell" tabindex="0" data-tip="' + g.tip + '" style="background:rgb(' + rgb.join(",") + ");color:" + ink + '">' +
          '<span class="v">' + pct(g.rate) + '</span><span class="n">n = ' + int(g.n) + (g.note ? " · " + esc(g.note) : "") + "</span></span>";
      });
    });
    document.getElementById(id).innerHTML = html + "</div>";
  }

  function kpiTile(label, value, sub, accent, valueColour) {
    return '<div class="kpi" style="--accent:' + accent + ";--value:" + (valueColour || "var(--ink)") + '">' +
      '<div class="k-label">' + esc(label) + '</div><div class="k-value">' + value + '</div><div class="k-sub">' + esc(sub) + "</div></div>";
  }

  // ---- pages ---------------------------------------------------------------
  function renderSummary(rows) {
    var s = stats(rows);
    var group = stats(rows.filter(function (i) { return L.group[C.group[i]] === "Yes"; }));
    var solo = stats(rows.filter(function (i) { return L.group[C.group[i]] === "No"; }));
    var revK = s.revenue >= 1e6 ? "$" + (s.revenue / 1e6).toFixed(1) + "M" : "$" + (s.revenue / 1000).toFixed(1) + "K";
    document.getElementById("kpis").innerHTML =
      kpiTile("Customers", int(s.n), "Total base", "var(--ink)") +
      kpiTile("Churned", int(s.churned), "Customers lost", "var(--hi)") +
      kpiTile("Churn rate", pct(s.rate), "Of all customers", "var(--hi)", "var(--hi)") +
      kpiTile("Monthly revenue lost", money(s.lost), pct(s.revenue ? s.lost / s.revenue : null, 0) + " of " + revK + " total", "var(--hi)", "var(--hi)") +
      kpiTile("Group-plan churn", pct(group.rate), solo.n ? "vs " + pct(solo.rate) + " non-group" : "No non-group customers in selection", "var(--good)", "var(--good)");

    // Share of churners by category, blanks excluded
    var churners = rows.filter(function (i) { return C.churn[i] && C.cat[i] >= 0; });
    var cats = groupBy(churners, function (i) { return C.cat[i]; });
    cats.forEach(function (g) { g.share = churners.length ? g.n / churners.length : 0; });
    cats.sort(function (a, b) { return b.share - a.share; });
    hbars("c-category", cats.map(function (g) {
      var name = L.cat[g.key];
      return { label: name, value: g.share, text: pct(g.share),
        colour: name === "Competitor" ? "var(--hi)" : "var(--mid)",
        tip: esc(name) + "|" + pct(g.share) + " of churned customers|" + int(g.n) + " customers" };
    }));

    var reasons = groupBy(rows.filter(function (i) { return C.churn[i] && C.reason[i] >= 0; }), function (i) { return C.reason[i]; });
    reasons.sort(function (a, b) { return b.n - a.n; });
    hbars("c-reasons", reasons.slice(0, 5).map(function (g) {
      var name = L.reason[g.key];
      var orange = name === "Competitor made better offer" || name === "Competitor had better devices";
      return { label: name, value: g.n, text: int(g.n), colour: orange ? "var(--hi)" : "var(--mid)",
        tip: esc(name) + "|" + int(g.n) + " churned customers" };
    }));
  }

  function renderContract(rows) {
    var m2m = rateWhere(rows, function (i) { return L.contract[C.contract[i]] === "Month-to-Month"; });
    var two = rateWhere(rows, function (i) { return L.contract[C.contract[i]] === "Two Year"; });
    var new6 = rateWhere(rows, function (i) { return tenIdx[i] === 0; });
    var head = (m2m != null && two)
      ? "Month-to-month customers churn " + Math.round(m2m / two) + "× more than two-year customers, and " + pct(new6, 0) + " of new customers leave within six months"
      : "Month-to-month vs two-year churn can't be compared for this selection; " + pct(new6, 0) + " of new customers leave within six months";
    document.getElementById("contract-headline").textContent = head;

    var order = ["Month-to-Month", "One Year", "Two Year"].map(function (n) { return lv("contract", n); });
    columns("c-contract", groupBy(rows, function (i) { return C.contract[i]; }, order).map(function (g) {
      var name = L.contract[g.key];
      return { label: name, sub: "n = " + int(g.n), value: g.rate, text: pct(g.rate), colour: churnColour(g.rate), tip: tip(name, g) };
    }));
    columns("c-tenure", groupBy(rows, function (i) { return tenIdx[i]; }, range(5)).map(function (g) {
      return { label: TENURE_GROUPS[g.key], value: g.rate, text: pct(g.rate), colour: churnColour(g.rate), tip: tip(TENURE_GROUPS[g.key] + " months", g) };
    }));
    var pay = groupBy(rows, function (i) { return C.payment[i]; }).sort(function (a, b) { return (b.rate || 0) - (a.rate || 0); });
    hbars("c-payment", pay.map(function (g) {
      var name = L.payment[g.key];
      return { label: name, value: g.rate, text: pct(g.rate), colour: churnColour(g.rate, 0.15), tip: tip(name, g) };
    }));
  }

  function renderService(rows) {
    columns("c-calls", groupBy(rows, function (i) { return callIdx[i]; }, range(6)).map(function (g) {
      return { label: CALL_GROUPS[g.key], sub: "n = " + int(g.n), value: g.rate, text: pct(g.rate, g.rate === 1 ? 0 : 1),
        colour: churnColour(g.rate), tip: tip(CALL_GROUPS[g.key] + " service calls", g) };
    }));
    var churned = rows.filter(function (i) { return C.churn[i]; });
    var heavy = churned.filter(function (i) { return C.calls[i] >= 3; }).length;
    document.getElementById("calls-note").textContent = churned.length
      ? "3+ calls = " + pct(heavy / churned.length, 0) + " of all churners in this selection."
      : "No churners in this selection.";

    var plan = [lv("intlPlan", "Yes"), lv("intlPlan", "No")], act = [lv("intlActive", "Yes"), lv("intlActive", "No")];
    var notes = [["good fit", "paying for unused plan"], ["paying extra charges", "good fit"]];
    matrix("c-intl", ["Has intl plan", "No intl plan"], ["Calls abroad", "No calls abroad"], function (r, c) {
      var g = stats(rows.filter(function (i) { return C.intlPlan[i] === plan[r] && C.intlActive[i] === act[c]; }));
      g.note = notes[r][c];
      g.tip = tip((r ? "No intl plan" : "Has intl plan") + ", " + (c ? "no calls abroad" : "calls abroad"), g, notes[r][c]);
      return g;
    }, "big");

    [["c-unlimited", "unlimited"], ["c-device", "device"]].forEach(function (p) {
      var gs = groupBy(rows, function (i) { return C[p[1]][i]; }).sort(function (a, b) { return (b.rate || 0) - (a.rate || 0); });
      hbars(p[0], gs.map(function (g) {
        var name = L[p[1]][g.key];
        return { label: name, value: g.rate, text: pct(g.rate), colour: churnColour(g.rate), tip: tip(name, g) };
      }), { compact: true });
    });
  }

  function renderWho(rows) {
    columns("c-age", groupBy(rows, function (i) { return ageIdx[i]; }, range(5)).map(function (g) {
      return { label: AGE_GROUPS[g.key], value: g.rate, text: pct(g.rate), colour: churnColour(g.rate), tip: tip("Age " + AGE_GROUPS[g.key], g) };
    }), { short: true });

    var contracts = ["Month-to-Month", "One Year", "Two Year"];
    matrix("c-senior", contracts, ["Senior", "Not senior"], function (r, c) {
      var ci = lv("contract", contracts[r]), si = lv("senior", c ? "No" : "Yes");
      var g = stats(rows.filter(function (i) { return C.contract[i] === ci && C.senior[i] === si; }));
      g.tip = tip(contracts[r] + ", " + (c ? "not senior" : "senior"), g);
      return g;
    }, "small");

    var grp = stats(rows.filter(function (i) { return L.group[C.group[i]] === "Yes"; }));
    var solo = stats(rows.filter(function (i) { return L.group[C.group[i]] === "No"; }));
    function avg(pred) { var r = rows.filter(pred), t = 0; r.forEach(function (i) { t += C.charge[i]; }); return r.length ? "$" + (t / r.length).toFixed(1) : "–"; }
    document.getElementById("c-group").innerHTML =
      kpiTile("Group plan", pct(grp.rate), "n = " + int(grp.n) + " · avg " + avg(function (i) { return L.group[C.group[i]] === "Yes"; }) + "/mo", "var(--good)", "var(--good)") +
      kpiTile("Individual", pct(solo.rate), "n = " + int(solo.n) + " · avg " + avg(function (i) { return L.group[C.group[i]] === "No"; }) + "/mo", "var(--hi)", "var(--hi)");

    var gen = groupBy(rows, function (i) { return C.gender[i]; }).sort(function (a, b) { return (b.rate || 0) - (a.rate || 0); });
    hbars("c-gender", gen.map(function (g) {
      var name = L.gender[g.key];
      return { label: name, value: g.rate, text: pct(g.rate), colour: "var(--mid)", tip: tip(name, g) };
    }));

    var states = groupBy(rows, function (i) { return C.state[i]; }).filter(function (g) { return g.n > 0; })
      .sort(function (a, b) { return b.rate - a.rate || b.n - a.n; }).slice(0, 10);
    var el = document.getElementById("c-states");
    if (!states.length) { el.innerHTML = '<p class="empty">No customers in this selection.</p>'; return; }
    el.innerHTML = '<table class="table"><thead><tr><th scope="col">State</th><th scope="col" class="num">Customers</th><th scope="col">Churn</th></tr></thead><tbody>' +
      states.map(function (g) {
        var name = L.state[g.key];
        return '<tr tabindex="0" data-tip="' + tip(name, g) + '" class="' + (g.rate >= 0.35 ? "hi" : "") + '"><td>' + esc(name) + '</td><td class="num">' + int(g.n) +
          '</td><td class="bar-cell"><div><i style="width:' + Math.max(2, g.rate * 100) + "%" + (g.rate >= 0.35 ? ";background:var(--hi)" : "") + '"></i><b>' + pct(g.rate) + "</b></div></td></tr>";
      }).join("") + "</tbody></table>";
  }

  function render() {
    var rows = selectedRows();
    var active = F.state >= 0 || F.age >= 0 || F.contract >= 0 || F.gender >= 0;
    document.getElementById("selection").innerHTML = active
      ? "<strong>" + int(rows.length) + "</strong> of " + int(N) + " customers in this selection. Headlines and recommendations describe all customers."
      : "Showing all <strong>" + int(N) + "</strong> customers.";
    renderSummary(rows); renderContract(rows); renderService(rows); renderWho(rows);
  }

  // ---- tabs (WAI-ARIA tabs pattern) ------------------------------------------
  var tabs = Array.prototype.slice.call(document.querySelectorAll('[role="tab"]'));
  function select(tab, focus) {
    tabs.forEach(function (t) {
      var on = t === tab;
      t.setAttribute("aria-selected", on); t.tabIndex = on ? 0 : -1;
      document.getElementById(t.getAttribute("aria-controls")).hidden = !on;
    });
    if (focus) tab.focus();
    try { history.replaceState(null, "", "#" + tab.id.replace("tab-", "")); } catch (e) { /* file:// */ }
  }
  tabs.forEach(function (t, i) {
    t.addEventListener("click", function () { select(t); });
    t.addEventListener("keydown", function (e) {
      var j = e.key === "ArrowRight" ? i + 1 : e.key === "ArrowLeft" ? i - 1 : e.key === "Home" ? 0 : e.key === "End" ? tabs.length - 1 : null;
      if (j == null) return;
      e.preventDefault(); select(tabs[(j + tabs.length) % tabs.length], true);
    });
  });
  var start = document.getElementById("tab-" + location.hash.slice(1));
  if (start) select(start);

  // ---- tooltip ---------------------------------------------------------------
  var tt = document.getElementById("tooltip");
  function showTip(target, x, y) {
    var parts = target.getAttribute("data-tip").split("|");
    tt.innerHTML = "<b>" + parts[0] + "</b>" + parts.slice(1).join("<br>");
    tt.hidden = false;
    var w = tt.offsetWidth, h = tt.offsetHeight;
    tt.style.left = Math.min(window.innerWidth - w - 8, x + 14) + "px";
    tt.style.top = (y + h + 20 > window.innerHeight ? y - h - 12 : y + 16) + "px";
  }
  document.addEventListener("mousemove", function (e) {
    var t = e.target.closest && e.target.closest("[data-tip]");
    if (t) showTip(t, e.clientX, e.clientY); else tt.hidden = true;
  });
  document.addEventListener("focusin", function (e) {
    var t = e.target.closest && e.target.closest("[data-tip]");
    if (!t) { tt.hidden = true; return; }
    var r = t.getBoundingClientRect(); showTip(t, r.left, r.bottom - 10);
  });
  document.addEventListener("keydown", function (e) { if (e.key === "Escape") tt.hidden = true; });

  render();
  if (window.matchMedia) {
    var mq = window.matchMedia("(prefers-color-scheme: dark)");
    if (mq.addEventListener) mq.addEventListener("change", render);  // heat colours read CSS vars
  }
})();
