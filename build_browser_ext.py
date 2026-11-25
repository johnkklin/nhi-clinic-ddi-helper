import os

EXT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "nhi_cloud_lab_ext")

MANIFEST_JSON = r"""{
  "manifest_version": 3,
  "name": "健保雲端 2.0 肝腎檢驗置頂助手 (GOT/GPT/Cr/eGFR)",
  "version": "1.0.0",
  "description": "於健保雲端查詢系統自動擷取最新 GOT(AST)、GPT(ALT)、Creatinine、eGFR 置頂顯示，並同步至本機 Python 警示視窗。",
  "host_permissions": [
    "*://medcloud2.nhi.gov.tw/*",
    "*://medcloud.nhi.gov.tw/*",
    "http://127.0.0.1:18765/*"
  ],
  "background": {
    "service_worker": "background.js"
  },
  "content_scripts": [
    {
      "matches": [
        "*://medcloud2.nhi.gov.tw/*",
        "*://medcloud.nhi.gov.tw/*"
      ],
      "js": ["page_hook.js"],
      "run_at": "document_start",
      "world": "MAIN",
      "all_frames": true
    },
    {
      "matches": [
        "*://medcloud2.nhi.gov.tw/*",
        "*://medcloud.nhi.gov.tw/*"
      ],
      "js": ["content.js"],
      "run_at": "document_idle"
    }
  ]
}
"""

BACKGROUND_JS = r"""// 背景 Service Worker：負責將網頁擷取之肝腎數值轉傳本機 Python (避開 HTTPS Mixed Content / CORS 限制)
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (!msg || msg.type !== "SYNC_TO_PYTHON") return false;

  fetch("http://127.0.0.1:18765/lab_sync", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(msg.payload || {})
  })
    .then(r => r.json())
    .then(data => sendResponse({ ok: true, data }))
    .catch(() => sendResponse({ ok: false }));

  return true; // 非同步回應
});
"""

PAGE_HOOK_JS = r"""// 唯讀監聽健保雲端網頁 API 回傳之 JSON（即使表格有分頁也能一次抓齊全部檢驗）
(function () {
  if (window.__nhiLabHookInstalled) return;
  window.__nhiLabHookInstalled = true;

  function emitJson(url, data) {
    try {
      if (!data || typeof data !== "object") return;
      window.top.postMessage(
        { __nhi_lab_ext__: true, sourceUrl: String(url || ""), payload: data },
        "*"
      );
    } catch (_) {}
  }

  const origFetch = window.fetch;
  if (typeof origFetch === "function") {
    window.fetch = async function (...args) {
      const resp = await origFetch.apply(this, args);
      try {
        const url = typeof args[0] === "string" ? args[0] : (args[0] && args[0].url) || "";
        const ct = (resp.headers && resp.headers.get("content-type")) || "";
        if (ct.includes("json") || url.includes("api") || url.includes("IMUE")) {
          resp.clone().json().then(d => emitJson(url, d)).catch(() => {});
        }
      } catch (_) {}
      return resp;
    };
  }

  const origOpen = XMLHttpRequest.prototype.open;
  const origSend = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function (method, url, ...rest) {
    this.__nhiUrl = url;
    return origOpen.call(this, method, url, ...rest);
  };
  XMLHttpRequest.prototype.send = function (...args) {
    this.addEventListener("load", function () {
      try {
        const txt = this.responseText;
        if (txt && (txt[0] === "{" || txt[0] === "[")) {
          emitJson(this.__nhiUrl, JSON.parse(txt));
        }
      } catch (_) {}
    });
    return origSend.apply(this, args);
  };
})();
"""

CONTENT_JS = r"""// 健保雲端 2.0 置頂肝腎檢驗色條 ＋ 自動翻查與本機 Python 同步
(function () {
  if (window.__nhiLabContentLoaded) return;
  window.__nhiLabContentLoaded = true;

  const AUTO_PEEK_KEY = "nhi_lab_auto_peek_enabled";
  let autoPeekEnabled = localStorage.getItem(AUTO_PEEK_KEY) !== "0";

  let labsState = { got: null, gpt: null, cr: null, egfr: null };
  let hasLabTabState = null; // null=未知, true=有檢驗頁籤, false=灰色無資料
  let pythonConnected = false;
  let lastPatSignature = "";
  let autoPeekDoneForPatient = false;
  let peekInProgress = false;
  let lastSyncHash = "";
  let scanTimer = null;

  // =========================================================================
  // 1. 建立網頁最上方固定置頂色條（下推 body 46px，絕不遮擋原網頁按鈕）
  // =========================================================================
  const bar = document.createElement("div");
  bar.id = "nhi-clinic-lab-topbar";
  bar.style.cssText = [
    "position:fixed",
    "top:0",
    "left:0",
    "right:0",
    "height:44px",
    "background:#17202A",
    "color:#F8F9F9",
    "z-index:2147483645",
    "display:flex",
    "align-items:center",
    "justify-content:space-between",
    "padding:0 12px",
    "box-shadow:0 2px 8px rgba(0,0,0,0.35)",
    "font-family:'Microsoft JhengHei',sans-serif",
    "font-size:13px",
    "box-sizing:border-box"
  ].join(";");

  bar.innerHTML = `
    <div style="display:flex;align-items:center;gap:8px;white-space:nowrap;">
      <span style="font-weight:bold;font-size:14px;color:#85C1E9;">🧪 雲端肝腎快篩</span>
      <span id="nhi-py-dot" title="與本機 Python 視窗連動狀態" style="font-size:11px;padding:2px 6px;border-radius:10px;background:#34495E;color:#BDC3C7;">⚪ 偵測本機連動</span>
    </div>
    <div id="nhi-lab-badges" style="display:flex;align-items:center;gap:8px;overflow-x:auto;padding:2px 4px;">
      <span style="color:#ABB2B9;">⏳ 等待讀取健保雲端檢驗資料...</span>
    </div>
    <div style="display:flex;align-items:center;gap:8px;white-space:nowrap;">
      <label style="display:flex;align-items:center;gap:4px;cursor:pointer;font-size:12px;color:#D5D8DC;" title="按下換卡後，若檢查與檢驗頁籤有資料(黑色)，自動點開抓取數值並切回原頁籤">
        <input type="checkbox" id="nhi-auto-peek-chk" ${autoPeekEnabled ? "checked" : ""} style="cursor:pointer;">
        換卡自動翻查
      </label>
      <button id="nhi-manual-peek-btn" style="background:#1F618D;color:#FFF;border:1px solid #2980B9;border-radius:4px;padding:3px 9px;cursor:pointer;font-size:12px;font-weight:bold;">
        🔍 抓檢驗值
      </button>
    </div>
  `;

  document.documentElement.appendChild(bar);
  if (document.body) {
    document.body.style.paddingTop = "46px";
  }

  document.getElementById("nhi-auto-peek-chk").addEventListener("change", (e) => {
    autoPeekEnabled = !!e.target.checked;
    localStorage.setItem(AUTO_PEEK_KEY, autoPeekEnabled ? "1" : "0");
  });

  document.getElementById("nhi-manual-peek-btn").addEventListener("click", () => {
    triggerLabTabPeek(true);
  });

  // =========================================================================
  // 2. 日期與檢驗數值解析工具（相容民國年/西元年、排除尿液肌酸酐與肺癌 EGFR）
  // =========================================================================
  function parseDateInfo(raw) {
    const s = String(raw || "").trim();
    if (!s) return { sortKey: 0, display: "" };

    let m = s.match(/\b(20\d{2})[\/\-.](\d{1,2})[\/\-.](\d{1,2})\b/);
    if (m) {
      const y = parseInt(m[1], 10), mo = parseInt(m[2], 10), d = parseInt(m[3], 10);
      return { sortKey: y * 10000 + mo * 100 + d, display: `${y - 1911}/${String(mo).padStart(2, "0")}/${String(d).padStart(2, "0")}` };
    }
    m = s.match(/\b(1\d{2}|0?\d{2})[\/\-.](\d{1,2})[\/\-.](\d{1,2})\b/);
    if (m) {
      const ry = parseInt(m[1], 10), mo = parseInt(m[2], 10), d = parseInt(m[3], 10);
      if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
        return { sortKey: (ry + 1911) * 10000 + mo * 100 + d, display: `${ry}/${String(mo).padStart(2, "0")}/${String(d).padStart(2, "0")}` };
      }
    }
    m = s.match(/\b(20\d{2})(\d{2})(\d{2})\b/);
    if (m) {
      const y = parseInt(m[1], 10), mo = parseInt(m[2], 10), d = parseInt(m[3], 10);
      if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
        return { sortKey: y * 10000 + mo * 100 + d, display: `${y - 1911}/${m[2]}/${m[3]}` };
      }
    }
    m = s.match(/\b(1\d{2})(\d{2})(\d{2})\b/);
    if (m) {
      const ry = parseInt(m[1], 10), mo = parseInt(m[2], 10), d = parseInt(m[3], 10);
      if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
        return { sortKey: (ry + 1911) * 10000 + mo * 100 + d, display: `${ry}/${m[2]}/${m[3]}` };
      }
    }
    return { sortKey: 0, display: "" };
  }

  function classifyMarker(codeStr, nameStr) {
    const code = String(codeStr || "").trim().toUpperCase();
    const name = String(nameStr || "").trim();
    const combined = `${code} ${name}`;

    // 1. GOT / AST (09025C)
    if (code === "09025C" || /\b(s[-_]?got|got|ast)\b/i.test(name) || /麩胺酸苯醋酸轉氨|天門冬胺酸轉氨/.test(name)) {
      return "got";
    }
    // 2. GPT / ALT (09026C) — 排除 GGT / r-GT
    if (
      !/r-?gt|ggt|γ-?gt|麩胺醯轉移|09031C/i.test(combined) &&
      (code === "09026C" || /\b(s[-_]?gpt|gpt|alt)\b/i.test(name) || /麩胺酸丙酮酸轉氨|丙胺酸轉氨/.test(name))
    ) {
      return "gpt";
    }
    // 3. eGFR — 排除肺癌 EGFR 基因突變檢測
    if (
      !/mutation|突變|exon|gene|基因|pcr|組織|切片/i.test(combined) &&
      (/\begfr\b|\bmdrd\b|ckd-epi|腎絲球過濾率/i.test(name))
    ) {
      return "egfr";
    }
    // 4. 血清 Creatinine (09015C) — 嚴格排除尿液肌酸酐/UACR/UPCR/CCr
    if (
      !/urine|尿|u-?cre|uacr|upcr|\bacr\b|\bpcr\b|albumin|protein|ccr|clearance|廓清|09016C|12111/i.test(combined) &&
      (code === "09015C" || /\b(crea?|creatinine)\b/i.test(name) || /肌酸酐/.test(name))
    ) {
      return "cr";
    }
    return null;
  }

  function parseLabNumeric(marker, rawVal) {
    let s = String(rawVal || "").trim();
    if (!s) return null;
    // 排除明顯為 7~8 碼純整數日期或醫令碼
    if (/^\d{6,10}$/.test(s)) return null;

    const m = s.match(/^([<>≦≧]=?)?\s*(\d+(?:\.\d+)?)\s*([HL↑↓*]+)?$/i);
    if (!m) return null;

    const prefix = m[1] || "";
    let num = parseFloat(m[2]);
    const flag = (m[3] || "").toUpperCase();
    if (isNaN(num)) return null;

    // 生理合理範圍防呆，避免抓錯欄位
    if ((marker === "got" || marker === "gpt") && (num <= 0 || num > 15000)) return null;
    if (marker === "cr" && (num < 0.05 || num > 40)) return null;
    if (marker === "egfr" && (num < 1 || num > 300)) return null;

    if (marker === "egfr" && (prefix.includes(">") || prefix.includes("≧")) && num >= 60) {
      num = Math.max(num, 60.1);
    }
    const cleanVal = `${prefix}${m[2]}${flag ? " " + flag : ""}`.trim();
    return { value: cleanVal, num, flag };
  }

  function recordCandidate(marker, parsedVal, unit, dateRaw, hosp, refRange) {
    if (!marker || !parsedVal) return false;
    const dt = parseDateInfo(dateRaw);
    const cur = labsState[marker];
    if (!cur || dt.sortKey >= cur.sortKey) {
      labsState[marker] = {
        value: parsedVal.value,
        num: parsedVal.num,
        flag: parsedVal.flag,
        unit: String(unit || "").trim(),
        date: dt.display || String(dateRaw || "").trim(),
        sortKey: dt.sortKey,
        hosp: String(hosp || "").trim(),
        ref: String(refRange || "").trim()
      };
      return true;
    }
    return false;
  }

  // =========================================================================
  // 3. 解析來源 A：API JSON 結構遞迴掃描
  // =========================================================================
  function scanJsonNode(node) {
    let updated = false;
    if (!node || typeof node !== "object") return false;

    if (Array.isArray(node)) {
      for (const item of node) {
        if (scanJsonNode(item)) updated = true;
      }
      return updated;
    }

    const keys = Object.keys(node);
    const getByPattern = (re) => {
      for (const k of keys) {
        if (re.test(k) && node[k] != null && typeof node[k] !== "object") return String(node[k]);
      }
      return "";
    };

    const code = getByPattern(/order_?code|nhi_?code|sOrder|cure_?item|inspect_?code/i);
    const name = getByPattern(/assay_?item|item_?name|order_?name|inspect_?name|chinese_?name|abbr_?name/i);
    const valRaw = getByPattern(/assay_?val|result_?val|inspect_?result|report_?val|test_?val|^value$/i);

    if ((code || name) && valRaw) {
      const marker = classifyMarker(code, name);
      if (marker) {
        const parsed = parseLabNumeric(marker, valRaw);
        const unit = getByPattern(/unit/i);
        const dt = getByPattern(/inspect_?date|collect_?date|report_?date|func_?date|real_?inspect/i);
        const hosp = getByPattern(/hosp_?name|hosp_?abbr/i);
        const ref = getByPattern(/consult_?val|ref_?val|normal_?val/i);
        if (recordCandidate(marker, parsed, unit, dt, hosp, ref)) updated = true;
      }
    }

    for (const k of keys) {
      const v = node[k];
      if (v && typeof v === "object") {
        if (scanJsonNode(v)) updated = true;
      }
    }
    return updated;
  }

  window.addEventListener("message", (ev) => {
    if (!ev.data || !ev.data.__nhi_lab_ext__) return;
    if (scanJsonNode(ev.data.payload)) {
      renderTopBarAndSync();
    }
  });

  // =========================================================================
  // 4. 解析來源 B：畫面 HTML 表格掃描（支援主畫面與同源 iframe）
  // =========================================================================
  function getAllAccessibleDocs() {
    const docs = [document];
    document.querySelectorAll("iframe").forEach((ifr) => {
      try {
        if (ifr.contentDocument) docs.push(ifr.contentDocument);
      } catch (_) {}
    });
    return docs;
  }

  function scanDomTables() {
    let updated = false;
    for (const doc of getAllAccessibleDocs()) {
      const tables = doc.querySelectorAll("table");
      tables.forEach((tbl) => {
        const ths = Array.from(tbl.querySelectorAll("thead th, tr:first-child th")).map(
          th => (th.innerText || "").replace(/\s+/g, "")
        );
        const findCol = (re) => ths.findIndex(t => re.test(t));
        const codeIdx = findCol(/醫令代碼|項目代碼|^代碼$/);
        const nameIdx = findCol(/檢驗項目|細項名稱|檢查項目|醫令名稱|項目名稱/);
        const valIdx = findCol(/檢驗值|檢驗結果|檢查結果|結果值|^結果$|^數值$/);
        const unitIdx = findCol(/單位/);
        const refIdx = findCol(/參考值|正常值|標準值/);
        const dateIdx = findCol(/採檢日|檢驗日|報告日|就醫日|日期/);
        const hospIdx = findCol(/院所|機構/);

        const rows = tbl.querySelectorAll("tbody tr, tr");
        rows.forEach((tr) => {
          const cells = Array.from(tr.querySelectorAll("td")).map(td => (td.innerText || "").trim());
          if (cells.length < 2) return;

          let marker = null;
          let matchedCellPos = -1;

          if (codeIdx >= 0 || nameIdx >= 0) {
            marker = classifyMarker(cells[codeIdx] || "", cells[nameIdx] || "");
            matchedCellPos = nameIdx >= 0 ? nameIdx : codeIdx;
          }
          if (!marker) {
            for (let i = 0; i < cells.length; i++) {
              const m = classifyMarker(cells[i], cells[i]);
              if (m) {
                marker = m;
                matchedCellPos = i;
                break;
              }
            }
          }
          if (!marker) return;

          // 優先抓「結果值欄」，否則往該項目右側鄰近欄尋找數值
          let parsed = null;
          if (valIdx >= 0 && cells[valIdx]) {
            parsed = parseLabNumeric(marker, cells[valIdx]);
          }
          if (!parsed && matchedCellPos >= 0) {
            for (let j = matchedCellPos + 1; j < Math.min(cells.length, matchedCellPos + 4); j++) {
              parsed = parseLabNumeric(marker, cells[j]);
              if (parsed) break;
            }
          }
          if (!parsed) return;

          const unit = unitIdx >= 0 ? cells[unitIdx] : "";
          const ref = refIdx >= 0 ? cells[refIdx] : "";
          const hosp = hospIdx >= 0 ? cells[hospIdx] : "";
          let dtRaw = dateIdx >= 0 ? cells[dateIdx] : "";
          if (!dtRaw) {
            dtRaw = cells.find(c => parseDateInfo(c).sortKey > 0) || "";
          }

          if (recordCandidate(marker, parsed, unit, dtRaw, hosp, ref)) {
            updated = true;
          }
        });
      });
    }
    return updated;
  }

  // =========================================================================
  // 5. 辨識病人基本資料與「檢查與檢驗」頁籤狀態
  // =========================================================================
  function detectPatientHints() {
    const headerTxt = (document.body ? document.body.innerText : "").slice(0, 1200);
    const idMatch = headerTxt.match(/\b([A-Z][1289*○ＯO][\d*○ＯO]{7}\d)\b/);
    let nameHint = "";
    const nameMatch = headerTxt.match(/姓名[：:\s]*([^\s｜|()0-9]{2,5})/);
    if (nameMatch) nameHint = nameMatch[1].trim();
    return {
      patIdHint: idMatch ? idMatch[1] : "",
      patNameHint: nameHint
    };
  }

  function findLabTabElement() {
    for (const doc of getAllAccessibleDocs()) {
      const candidates = doc.querySelectorAll(
        "a, button, li, div[role='tab'], span.el-tabs__item, .nav-link, .tab-item"
      );
      for (const el of candidates) {
        const txt = (el.innerText || "").replace(/\s+/g, "");
        if (/^(檢查與檢驗|檢驗檢查|檢驗\(查\)結果)(\(\d+\)|（\d+）)?$/.test(txt)) {
          return el;
        }
      }
    }
    return null;
  }

  function isTabDisabledOrGray(el) {
    if (!el) return true;
    const txt = (el.innerText || "").replace(/\s+/g, "");
    if (/\(0\)|（0）/.test(txt)) return true;
    if (el.disabled || el.getAttribute("aria-disabled") === "true") return true;
    const cls = `${el.className || ""} ${(el.parentElement && el.parentElement.className) || ""}`;
    if (/disabled|inactive|no-data|empty|gray|grey/i.test(cls)) return true;

    const st = window.getComputedStyle(el);
    if (st.pointerEvents === "none" || st.cursor === "not-allowed" || parseFloat(st.opacity || "1") < 0.55) {
      return true;
    }
    const rgb = st.color.match(/(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
    if (rgb) {
      const r = +rgb[1], g = +rgb[2], b = +rgb[3];
      // 淺灰字判定
      if (r > 165 && g > 165 && b > 165) return true;
    }
    return false;
  }

  function findCurrentActiveTab(labTabEl) {
    if (!labTabEl || !labTabEl.parentElement) return null;
    const container = labTabEl.closest("ul, .el-tabs__nav, .nav, .tabs") || labTabEl.parentElement.parentElement;
    if (!container) return null;
    return container.querySelector(".active, .is-active, [aria-selected='true']");
  }

  function triggerLabTabPeek(forceManual = false) {
    if (peekInProgress) return;
    const labTab = findLabTabElement();
    if (!labTab) return;

    if (isTabDisabledOrGray(labTab)) {
      hasLabTabState = false;
      autoPeekDoneForPatient = true;
      renderTopBarAndSync();
      return;
    }

    hasLabTabState = true;
    if (!forceManual && (!autoPeekEnabled || autoPeekDoneForPatient)) {
      renderTopBarAndSync();
      return;
    }

    autoPeekDoneForPatient = true;
    peekInProgress = true;
    const prevTab = findCurrentActiveTab(labTab);
    const needReturn = !forceManual && prevTab && prevTab !== labTab && !prevTab.contains(labTab);

    labTab.click();

    let attempts = 0;
    const poll = setInterval(() => {
      attempts++;
      scanDomTables();
      const hasAny = !!(labsState.got || labsState.gpt || labsState.cr || labsState.egfr);
      if (hasAny || attempts >= 6) {
        clearInterval(poll);
        if (needReturn) {
          setTimeout(() => {
            try { prevTab.click(); } catch (_) {}
            peekInProgress = false;
            renderTopBarAndSync();
          }, 120);
        } else {
          peekInProgress = false;
          renderTopBarAndSync();
        }
      }
    }, 180);
  }

  // =========================================================================
  // 6. 置頂色條渲染與本機 Python 視窗同步
  // =========================================================================
  function makeBadgeHtml(label, item, levelFn) {
    if (!item) {
      return `<span style="background:#2C3E50;color:#95A5A6;padding:3px 9px;border-radius:5px;font-size:12px;">${label}: 無近期值</span>`;
    }
    const lvl = levelFn(item);
    const palette = {
      CRITICAL: "background:#C0392B;color:#FFF;border:1px solid #FF6B6B;font-weight:bold;",
      HIGH: "background:#D35400;color:#FFF;border:1px solid #F39C12;font-weight:bold;",
      MODERATE: "background:#F1C40F;color:#1C2833;font-weight:bold;",
      NORMAL: "background:#1E8449;color:#FFF;font-weight:bold;"
    };
    const icon = lvl === "CRITICAL" ? "🚨 " : (lvl === "HIGH" || lvl === "MODERATE" ? "⚠️ " : "✅ ");
    const dt = item.date ? ` <small style="opacity:0.9;">(${item.date})</small>` : "";
    const tip = `院所：${item.hosp || "雲端"}\n日期：${item.date || "-"}\n單位：${item.unit || "-"}\n參考值：${item.ref || "-"}`;
    return `<span title="${tip}" style="${palette[lvl] || palette.NORMAL}padding:3px 10px;border-radius:5px;font-size:13px;">${icon}${label}: ${item.value}${dt}</span>`;
  }

  function renderTopBarAndSync() {
    const box = document.getElementById("nhi-lab-badges");
    if (!box) return;

    const anyFound = !!(labsState.got || labsState.gpt || labsState.cr || labsState.egfr);
    if (!anyFound && hasLabTabState === false) {
      box.innerHTML = `<span style="background:#34495E;color:#D5D8DC;padding:3px 12px;border-radius:5px;">⚪ 本次雲端近半年無「檢查與檢驗」紀錄（頁籤為灰色）</span>`;
    } else if (!anyFound && hasLabTabState === true && !autoPeekEnabled) {
      box.innerHTML = `<span style="background:#7D6608;color:#FEF9E7;padding:3px 12px;border-radius:5px;">🔔 偵測到有檢驗紀錄！請點右側「🔍 抓檢驗值」或點開檢驗頁籤</span>`;
    } else {
      const bGot = makeBadgeHtml("GOT(AST)", labsState.got, it =>
        it.num >= 80 ? "CRITICAL" : (it.num > 40 || it.flag ? "HIGH" : "NORMAL")
      );
      const bGpt = makeBadgeHtml("GPT(ALT)", labsState.gpt, it =>
        it.num >= 80 ? "CRITICAL" : (it.num > 40 || it.flag ? "HIGH" : "NORMAL")
      );
      const bCr = makeBadgeHtml("Cr", labsState.cr, it =>
        it.num >= 2.0 ? "CRITICAL" : (it.num > 1.3 || it.flag ? "MODERATE" : "NORMAL")
      );
      const bEgfr = makeBadgeHtml("eGFR", labsState.egfr, it =>
        it.num < 30 ? "CRITICAL" : (it.num < 45 ? "HIGH" : (it.num < 60 ? "MODERATE" : "NORMAL"))
      );
      box.innerHTML = [bGot, bGpt, bCr, bEgfr].join("");
    }

    const hints = detectPatientHints();
    const payload = {
      patIdHint: hints.patIdHint,
      patNameHint: hints.patNameHint,
      hasLabTab: hasLabTabState,
      labs: labsState
    };

    const syncHash = JSON.stringify(payload);
    if (syncHash === lastSyncHash) return;
    lastSyncHash = syncHash;

    chrome.runtime.sendMessage({ type: "SYNC_TO_PYTHON", payload }, (res) => {
      const dot = document.getElementById("nhi-py-dot");
      if (!dot) return;
      pythonConnected = !!(res && res.ok);
      if (pythonConnected) {
        dot.textContent = "🟢 已同步桌面視窗";
        dot.style.background = "#145A32";
        dot.style.color = "#ABEBC6";
      } else {
        dot.textContent = "⚪ 僅網頁顯示";
        dot.style.background = "#34495E";
        dot.style.color = "#BDC3C7";
      }
    });
  }

  function resetForNewCardRead() {
    labsState = { got: null, gpt: null, cr: null, egfr: null };
    hasLabTabState = null;
    autoPeekDoneForPatient = false;
    lastSyncHash = "";
    const box = document.getElementById("nhi-lab-badges");
    if (box) {
      box.innerHTML = `<span style="color:#F7DC6F;font-weight:bold;">⏳ 已按換卡，正在讀取新病人檢驗資料...</span>`;
    }
    setTimeout(runRoutineCheck, 1300);
  }

  function runRoutineCheck() {
    const hints = detectPatientHints();
    const sig = `${hints.patIdHint}|${hints.patNameHint}`;
    if (sig !== "|" && sig !== lastPatSignature) {
      if (lastPatSignature !== "") {
        labsState = { got: null, gpt: null, cr: null, egfr: null };
        hasLabTabState = null;
        autoPeekDoneForPatient = false;
        lastSyncHash = "";
      }
      lastPatSignature = sig;
    }

    scanDomTables();
    const anyFound = !!(labsState.got || labsState.gpt || labsState.cr || labsState.egfr);
    if (anyFound) {
      hasLabTabState = true;
      renderTopBarAndSync();
      return;
    }

    const labTab = findLabTabElement();
    if (labTab) {
      if (isTabDisabledOrGray(labTab)) {
        hasLabTabState = false;
        renderTopBarAndSync();
      } else {
        hasLabTabState = true;
        if (autoPeekEnabled && !autoPeekDoneForPatient) {
          triggerLabTabPeek(false);
        } else {
          renderTopBarAndSync();
        }
      }
    }
  }

  // 監聽綠色「請換卡再按我」按鈕點擊
  document.addEventListener(
    "click",
    (e) => {
      const btn = e.target && e.target.closest("button, a, input[type='button']");
      if (!btn) return;
      const txt = (btn.innerText || btn.value || "").replace(/\s+/g, "");
      if (/換卡|重新讀卡|讀取健保卡/.test(txt)) {
        resetForNewCardRead();
      }
    },
    true
  );

  // 監聽畫面表格渲染變動
  const obs = new MutationObserver(() => {
    if (peekInProgress) return;
    clearTimeout(scanTimer);
    scanTimer = setTimeout(runRoutineCheck, 450);
  });
  if (document.body) {
    obs.observe(document.body, { childList: true, subtree: true });
  }

  setTimeout(runRoutineCheck, 1200);
})();
"""


def main():
    os.makedirs(EXT_DIR, exist_ok=True)
    files = {
        "manifest.json": MANIFEST_JSON.strip() + "\n",
        "background.js": BACKGROUND_JS.strip() + "\n",
        "page_hook.js": PAGE_HOOK_JS.strip() + "\n",
        "content.js": CONTENT_JS.strip() + "\n",
    }
    for fname, content in files.items():
        fpath = os.path.join(EXT_DIR, fname)
        with open(fpath, "w", encoding="utf-8") as f:
            f.write(content)
    print(f"✅ 瀏覽器外掛資料夾已產生：{EXT_DIR}")


if __name__ == "__main__":
    main()