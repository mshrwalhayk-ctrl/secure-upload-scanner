/* =====================================================================
   مساعد المحاسب  —  MaliAI Excel Task Pane
   لوحة جانبية عربية بالكامل تشتغل جوّا إكسل عبر Office.js
   ===================================================================== */
"use strict";

/* ---------------------------------------------------------------------
   1) الإعدادات  (محفوظة على جهاز المستخدم فقط)
   ------------------------------------------------------------------- */
/* المزوّدون الجاهزون.
   الافتراضي OmniRoute: بوابة محلية بتوحّد كل المزوّدين، بتضغط النص،
   بتشتغل بدون مفتاح، وبتحلّ مشكلة CORS لأنها على نفس الجهاز.
   التركيب مرة وحدة:   npm i -g omniroute      ثم:   omniroute            */
const PROVIDERS = {
  omniroute: { label:"OmniRoute — بوابة محلية ★ (بدون مفتاح + ضغط نص)",
               url:"http://localhost:20128/v1", model:"auto", needsKey:false, tools:true,
               hint:"ركّبها مرة وحدة من cmd:  npm i -g omniroute   ثم شغّلها:  omniroute" },
  openrouter:{ label:"OpenRouter (نماذج مجانية ★)",
               url:"https://openrouter.ai/api/v1", model:"nvidia/nemotron-3-ultra-550b-a55b:free",
               needsKey:true, tools:true,
               hint:"مفتاح مجاني من openrouter.ai ← Keys. اختر نموذج عليه ★." },
  groq:      { label:"Groq (سريع مجاني)",
               url:"https://api.groq.com/openai/v1", model:"llama-3.3-70b-versatile",
               needsKey:true, tools:true,
               hint:"مفتاح مجاني من console.groq.com ← API Keys." },
  nvidia:    { label:"NVIDIA NIM (أقوى النماذج المجانية)",
               url:"https://integrate.api.nvidia.com/v1", model:"nvidia/nemotron-3-super-120b-a12b",
               needsKey:true, tools:true,
               hint:"⚠ NVIDIA ما بتسمح بالاتصال المباشر من المتصفح (CORS). استخدمها عبر OmniRoute." },
  custom:    { label:"مزوّد مخصّص (رابط خاص)",
               url:"", model:"", needsKey:true, tools:true,
               hint:"حط العنوان الأساسي بدون /chat/completions." }
};

const PRESET = {
  provider: "omniroute",
  url:   PROVIDERS.omniroute.url,
  model: PROVIDERS.omniroute.model,
  key:   "",
  steps: 14,
  confirm: false,
  nativeTools: true
};

const CFG = {
  read(){
    let s = {};
    try { s = JSON.parse(localStorage.getItem("maliai.cfg") || "{}"); } catch(e){}
    return Object.assign({}, PRESET, s);
  },
  write(o){
    try { localStorage.setItem("maliai.cfg", JSON.stringify(o)); } catch(e){}
  },
  ok(){
    const c = this.read();
    const p = PROVIDERS[c.provider] || PROVIDERS.custom;
    return !!(c.url && c.model && (c.key || !p.needsKey));
  }
};

/* ---------------------------------------------------------------------
   2) أدوات مساعدة عامة
   ------------------------------------------------------------------- */
const $ = id => document.getElementById(id);

function esc(s){ const d = document.createElement("div"); d.textContent = s; return d.innerHTML; }

/* تقريب تجميلي لـ 15 خانة معنوية — نفس اللي بيعمله إكسل عند العرض */
function cosmetic(d){
  if (!isFinite(d) || d === 0) return d;
  const a = Math.abs(d);
  if (a >= 1e15 || a <= 1e-14) return d;
  const mag = Math.floor(Math.log10(a));
  const f = Math.pow(10, 14 - mag);
  return Math.round(d * f) / f;
}

function numText(v){
  if (typeof v !== "number" || !isFinite(v)) return String(v);
  let s = String(cosmetic(v));
  if (s.includes("e") || s.includes("E")) return s;
  return s;
}

/* جمع تعويضي (Neumaier) — بيمنع تراكم خطأ الفاصلة العائمة في الأرقام المالية */
function compensatedSum(nums){
  let sum = 0, c = 0;
  for (const x of nums){
    const t = sum + x;
    c += (Math.abs(sum) >= Math.abs(x)) ? (sum - t) + x : (x - t) + sum;
    sum = t;
  }
  return sum + c;
}

/* تحويل الأرقام العربية والرموز قبل التقييم */
function normalizeExpr(s){
  let out = "";
  for (const ch of String(s)){
    const c = ch.codePointAt(0);
    if (c >= 0x660 && c <= 0x669) out += String(c - 0x660);
    else if (c >= 0x6F0 && c <= 0x6F9) out += String(c - 0x6F0);
    else if (c === 0x66B) out += ".";
    else if (c === 0x66C) continue;
    else if (ch === "×") out += "*";
    else if (ch === "÷") out += "/";
    else if (ch === "−") out += "-";
    else if (c === 0x200F || c === 0x200E || c === 0x61C) continue;
    else out += ch;
  }
  out = out.trim();
  if (out.startsWith("=")) out = out.slice(1);
  return out;
}

function cleanAddr(a){
  if (!a) return a;
  a = String(a).trim();
  const i = a.indexOf("!");
  if (i >= 0) a = a.slice(i + 1);
  return a.replace(/\$/g, "");
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

/* ---------------------------------------------------------------------
   3) طبقة الاتصال بالنموذج (صيغة OpenAI) مع إعادة محاولة تلقائية
   ------------------------------------------------------------------- */
async function aiChat(system, history, tools){
  const c = CFG.read();
  const body = {
    model: c.model,
    temperature: 0.2,
    max_tokens: 4096,
    messages: [{ role:"system", content: system }, ...history]
  };
  if (tools && tools.length){ body.tools = tools; body.tool_choice = "auto"; }

  let lastErr = "";
  for (let attempt = 1; attempt <= 3; attempt++){
    try {
      const hdrs = { "Content-Type":"application/json" };
      if (c.key) hdrs["Authorization"] = "Bearer " + c.key;
      const r = await fetch(c.url.replace(/\/+$/,"") + "/chat/completions", {
        method: "POST", headers: hdrs, body: JSON.stringify(body)
      });
      const txt = await r.text();

      if (r.status === 401 || r.status === 403)
        return { ok:false, err:"المفتاح مرفوض أو منتهي (" + r.status + "). تأكد منه في الإعدادات." };
      if (r.status === 404)
        return { ok:false, err:"اسم النموذج أو الرابط غلط (404). تأكد منهم في الإعدادات." };
      if (r.status === 400 || r.status === 422)
        return { ok:false, err:"NOTOOLS", detail:txt.slice(0,200) };

      let j = null;
      try { j = JSON.parse(txt); } catch(e){}

      const m = j && j.choices && j.choices[0] && j.choices[0].message;
      if (m && (m.content || (m.tool_calls && m.tool_calls.length)))
        return { ok:true, text: m.content || "", toolCalls: m.tool_calls || [] };

      if (j && j.error && j.error.message){
        lastErr = "الخادم: " + j.error.message;
        if (r.status !== 429 && r.status < 500) return { ok:false, err: lastErr };
      } else {
        lastErr = "رد غير مفهوم من الخادم (" + r.status + "): " + txt.slice(0,180);
      }
    } catch(e){
      const cfg = CFG.read();
      lastErr = (cfg.provider === "omniroute")
        ? "OmniRoute مش شغّالة على الجهاز.\nافتح cmd واكتب:  omniroute\nوإذا أول مرة:  npm i -g omniroute"
        : "تعذّر الاتصال بالخدمة (غالباً المزوّد ما بيسمح بالاتصال من المتصفح — CORS).\nالحل: استخدم OmniRoute كبوابة محلية.";
    }
    if (attempt < 3) await sleep(1800);
  }
  return { ok:false, err: lastErr || "فشل الاتصال بعد 3 محاولات." };
}

async function listModels(){
  const c = CFG.read();
  const r = await fetch(c.url.replace(/\/+$/,"") + "/models", {
    headers: { "Authorization":"Bearer " + c.key }
  });
  const j = await r.json();
  return (j.data || []).map(m => m.id).filter(Boolean);
}

/* =====================================================================
   4) أدوات إكسل  —  اللي بيقدر النموذج يشغّلها على الملف
   ===================================================================== */

const MAX_ROWS = 300, MAX_COLS = 40;
let UNDO = [];                       // مكدّس التراجع

/* --- إيجاد ورقة بالاسم (بدون حساسية لحالة الأحرف) ------------------- */
async function resolveSheet(ctx, name){
  const sheets = ctx.workbook.worksheets;
  sheets.load("items/name");
  await ctx.sync();
  if (!name || !String(name).trim()){
    const ws = sheets.getActiveWorksheet();
    ws.load("name");
    await ctx.sync();
    return ws;
  }
  const target = String(name).trim().toLowerCase();
  const hit = sheets.items.find(s => s.name.trim().toLowerCase() === target);
  if (!hit){
    throw new Error('ما في ورقة اسمها "' + name + '". الأوراق الموجودة: ' +
                    sheets.items.map(s => s.name).join(" ، "));
  }
  return sheets.getItem(hit.name);
}

async function getRangeOf(ctx, ws, addr){
  if (!addr || !String(addr).trim()){
    const sel = ctx.workbook.getSelectedRange();
    sel.load("address");
    await ctx.sync();
    return sel;
  }
  return ws.getRange(cleanAddr(addr));
}

/* --- حفظ حالة نطاق قبل التعديل (للتراجع) ---------------------------- */
async function pushUndo(ctx, ws, rng, desc){
  try{
    rng.load(["address","formulas","numberFormat","rowCount","columnCount"]);
    ws.load("name");
    await ctx.sync();
    if (rng.rowCount * rng.columnCount > 60000) return;
    UNDO.push({
      sheet: ws.name,
      address: cleanAddr(rng.address),
      formulas: rng.formulas,
      numberFormat: rng.numberFormat,
      desc
    });
    if (UNDO.length > 20) UNDO.shift();
  }catch(e){}
}

async function undoLast(){
  if (!UNDO.length) return "ما في شي للتراجع عنه.";
  const u = UNDO.pop();
  try{
    await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(u.sheet);
      const rng = ws.getRange(u.address);
      rng.formulas = u.formulas;
      try { rng.numberFormat = u.numberFormat; } catch(e){}
      await ctx.sync();
    });
    return "تم التراجع عن: " + u.desc + "  (باقي " + UNDO.length + " خطوة)";
  }catch(e){
    return "تعذّر التراجع: " + e.message;
  }
}

/* --- تقييم تعبير بمحرّك إكسل نفسه عبر خلية مؤقتة على نفس الورقة ------ */
async function evalOnSheet(ctx, ws, expr){
  const ur = ws.getUsedRange(true);
  ur.load(["rowIndex","columnIndex","rowCount","columnCount"]);
  await ctx.sync();

  let r = 0, c = 0;
  if (!ur.isNullObject){
    r = (ur.rowIndex || 0) + (ur.rowCount || 0) + 3;
    c = (ur.columnIndex || 0) + (ur.columnCount || 0) + 2;
  } else { r = 3; c = 3; }
  if (r > 1048500) r = 1048500;
  if (c > 16300) c = 16300;

  const cell = ws.getCell(r, c);
  cell.load(["formulas"]);
  await ctx.sync();
  const backup = cell.formulas;

  cell.formulas = [["=" + expr]];
  await ctx.sync();

  cell.load(["values","text","valueTypes"]);
  await ctx.sync();

  const v = cell.values[0][0];
  const t = cell.text[0][0];
  const vt = cell.valueTypes[0][0];

  cell.formulas = backup;
  await ctx.sync();

  return { value: v, text: t, type: vt };
}

/* ===================== الأدوات ===================== */

async function T_list_sheets(){
  return Excel.run(async ctx => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    const wb = ctx.workbook; wb.load("name");
    await ctx.sync();

    const infos = sheets.items.map(s => {
      const ur = ctx.workbook.worksheets.getItem(s.name).getUsedRange(true);
      ur.load(["address","isNullObject"]);
      return { name: s.name, ur };
    });
    await ctx.sync();

    let out = "الملف: " + wb.name + "\n";
    for (const i of infos){
      out += "- " + i.name + "  (النطاق المستخدم: " +
             (i.ur.isNullObject ? "فاضية" : cleanAddr(i.ur.address)) + ")\n";
    }
    return out;
  });
}

async function T_read_range(a){
  return Excel.run(async ctx => {
    const ws = await resolveSheet(ctx, a.sheet);
    const rng = await getRangeOf(ctx, ws, a.address);
    rng.load(["address","text","formulas","rowCount","columnCount"]);
    ws.load("name");
    await ctx.sync();

    const nR = Math.min(rng.rowCount, MAX_ROWS);
    const nC = Math.min(rng.columnCount, MAX_COLS);
    const cut = (nR < rng.rowCount) || (nC < rng.columnCount);

    let out = "الورقة: " + ws.name + "   النطاق: " + cleanAddr(rng.address) +
              "  (" + rng.rowCount + " صف × " + rng.columnCount + " عمود)\n";

    for (let i = 0; i < nR; i++){
      const cells = [];
      for (let j = 0; j < nC; j++){
        let s = String(rng.text[i][j]);
        const f = rng.formulas[i][j];
        if (typeof f === "string" && f.startsWith("=")) s += " {" + f + "}";
        cells.push(s);
      }
      out += cells.join("\t") + "\n";
      if (out.length > 22000) { return out + "\n(تم اقتطاع الناتج)"; }
    }
    if (cut) out += "\n(معروض " + nR + "×" + nC + " من أصل " + rng.rowCount + "×" + rng.columnCount + ")";
    return out;
  });
}

async function T_calc(a){
  const expr = normalizeExpr(a.expression || "");
  if (!expr) return "خطأ: الحقل expression مطلوب.";
  return Excel.run(async ctx => {
    const ws = await resolveSheet(ctx, a.sheet);
    const r = await evalOnSheet(ctx, ws, expr);
    if (r.type === "Error") return "خطأ في التعبير: " + r.text;
    const v = (typeof r.value === "number") ? numText(r.value) : String(r.value);
    return "النتيجة الدقيقة لـ (" + expr + ") = " + v;
  });
}

async function T_precise_sum(a){
  return Excel.run(async ctx => {
    const ws = await resolveSheet(ctx, a.sheet);
    const rng = await getRangeOf(ctx, ws, a.address);
    rng.load(["address","values","valueTypes"]);
    ws.load("name");
    await ctx.sync();

    const nums = [];
    let textNums = 0;
    for (let i = 0; i < rng.values.length; i++)
      for (let j = 0; j < rng.values[i].length; j++){
        const v = rng.values[i][j], t = rng.valueTypes[i][j];
        if (t === "Double" || t === "Integer") nums.push(Number(v));
        else if (t === "String" && v !== "" && !isNaN(Number(String(v).replace(/,/g,"")))) textNums++;
      }

    let out = "المجموع الدقيق لـ " + ws.name + "!" + cleanAddr(rng.address) +
              " = " + numText(compensatedSum(nums)) +
              "   (خلايا رقمية: " + nums.length + ")";
    if (textNums) out += "\n⚠ في " + textNums + " خلية رقمها مخزّن كنص — ما دخلت بالمجموع.";
    return out;
  });
}

async function T_write_values(a){
  if (!a.address) return "خطأ: الحقل address مطلوب.";
  if (!Array.isArray(a.values) || !a.values.length) return "خطأ: values لازم تكون مصفوفة صفوف.";

  const rows = a.values.map(r => Array.isArray(r) ? r : [r]);
  const nC = Math.max(...rows.map(r => r.length));
  const grid = rows.map(r => {
    const out = r.slice();
    while (out.length < nC) out.push("");
    return out.map(v => (v === null || v === undefined) ? "" : v);
  });

  return Excel.run(async ctx => {
    const ws = await resolveSheet(ctx, a.sheet);
    const start = ws.getRange(cleanAddr(a.address)).getCell(0,0);
    const target = start.getResizedRange(grid.length - 1, nC - 1);
    await pushUndo(ctx, ws, target, "كتابة قيم");
    target.values = grid;
    target.load("address");
    ws.load("name");
    await ctx.sync();
    return "تم كتابة " + grid.length + " صف × " + nC + " عمود في " +
           ws.name + "!" + cleanAddr(target.address) + ".";
  });
}

function colIndex(name){
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/* إزاحة المراجع النسبية في معادلة A1 (بديل احتياطي لـ copyFrom) */
function shiftFormula(f, dr, dc){
  if (dr === 0 && dc === 0) return f;
  return f.replace(
    /("[^"]*")|(?<![A-Za-z0-9_$])(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(])/g,
    (m, str, cAbs, col, rAbs, row) => {
      if (str) return str;
      let c = colIndex(col);
      let r = parseInt(row, 10);
      if (!cAbs) c += dc;
      if (!rAbs) r += dr;
      if (c < 0) c = 0;
      if (r < 1) r = 1;
      return (cAbs || "") + colName(c) + (rAbs || "") + r;
    });
}

async function T_write_formula(a){
  if (!a.address) return "خطأ: الحقل address مطلوب.";
  let f = String(a.formula || "");
  if (!f) return "خطأ: الحقل formula مطلوب.";
  if (!f.startsWith("=")) f = "=" + f;

  return Excel.run(async ctx => {
    const ws = await resolveSheet(ctx, a.sheet);
    const rng = ws.getRange(cleanAddr(a.address));
    await pushUndo(ctx, ws, rng, "كتابة معادلة");

    rng.load(["rowCount","columnCount","address"]);
    ws.load("name");
    await ctx.sync();

    // نكتب المعادلة في أول خلية، ثم ننسخها على باقي النطاق
    // حتى تتزحزح المراجع النسبية لكل صف — زي ما بيعمل إكسل بالسحب.
    const top = rng.getCell(0, 0);
    top.formulas = [[f]];
    await ctx.sync();

    const many = (rng.rowCount * rng.columnCount) > 1;
    if (many){
      let canCopy = false;
      try {
        canCopy = Office.context.requirements.isSetSupported("ExcelApi", "1.9");
      } catch(e){ canCopy = false; }

      if (canCopy){
        rng.copyFrom(top, Excel.RangeCopyType.formulas);
        await ctx.sync();
      } else {
        const grid = [];
        for (let i = 0; i < rng.rowCount; i++){
          const row = [];
          for (let j = 0; j < rng.columnCount; j++) row.push(shiftFormula(f, i, j));
          grid.push(row);
        }
        rng.formulas = grid;
        await ctx.sync();
      }
    }

    const first = rng.getCell(0, 0);
    const last  = rng.getCell(rng.rowCount - 1, rng.columnCount - 1);
    first.load(["text","valueTypes"]);
    last.load(["text","formulas"]);
    await ctx.sync();

    const shown = (first.valueTypes[0][0] === "Error")
      ? "⚠ الخلية رجّعت خطأ: " + first.text[0][0]
      : "أول نتيجة: " + first.text[0][0];

    let tail = "";
    if (many) tail = "   آخر خلية: " + last.formulas[0][0] + " = " + last.text[0][0];

    return "تم وضع المعادلة " + f + " في " + ws.name + "!" + cleanAddr(rng.address) +
           ". " + shown + tail;
  });
}

async function T_format_range(a){
  return Excel.run(async ctx => {
    const ws = await resolveSheet(ctx, a.sheet);
    const rng = await getRangeOf(ctx, ws, a.address);
    await pushUndo(ctx, ws, rng, "تنسيق");
    rng.load(["address","rowCount","columnCount"]);
    ws.load("name");
    await ctx.sync();

    if (a.number_format){
      const grid = [];
      for (let i = 0; i < rng.rowCount; i++)
        grid.push(new Array(rng.columnCount).fill(a.number_format));
      rng.numberFormat = grid;
    }
    if (a.bold !== undefined)   rng.format.font.bold = !!a.bold;
    if (a.italic !== undefined) rng.format.font.italic = !!a.italic;
    if (a.font_size)  rng.format.font.size = Number(a.font_size);
    if (a.font_color) rng.format.font.color = String(a.font_color);
    if (a.fill){
      if (String(a.fill).toLowerCase() === "none") rng.format.fill.clear();
      else rng.format.fill.color = String(a.fill);
    }
    if (a.align){
      const m = { right:"Right", left:"Left", center:"Center" };
      if (m[String(a.align).toLowerCase()]) rng.format.horizontalAlignment = m[String(a.align).toLowerCase()];
    }
    if (a.wrap !== undefined) rng.format.wrapText = !!a.wrap;
    if (a.borders){
      for (const b of ["EdgeTop","EdgeBottom","EdgeLeft","EdgeRight","InsideVertical","InsideHorizontal"]){
        try { const e = rng.format.borders.getItem(b); e.style = "Continuous"; e.color = "#AAAAAA"; } catch(err){}
      }
    }
    await ctx.sync();
    if (a.autofit){ try { rng.format.autofitColumns(); await ctx.sync(); } catch(e){} }

    return "تم تنسيق " + ws.name + "!" + cleanAddr(rng.address) + ".";
  });
}

async function T_add_sheet(a){
  return Excel.run(async ctx => {
    const name = String(a.name || "ورقة جديدة").slice(0,31);
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    if (sheets.items.some(s => s.name.toLowerCase() === name.toLowerCase()))
      return 'الورقة "' + name + '" موجودة أصلاً، رح نستخدمها.';
    const ws = sheets.add(name);
    ws.load("name");
    await ctx.sync();
    return 'تم إنشاء ورقة جديدة اسمها "' + ws.name + '".';
  });
}

async function T_find_text(a){
  const needle = String(a.text || "");
  if (!needle) return "خطأ: الحقل text مطلوب.";
  return Excel.run(async ctx => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();

    const names = a.sheet ? [String(a.sheet)] : sheets.items.map(s => s.name);
    let out = "", hits = 0;

    for (const n of names){
      const ws = await resolveSheet(ctx, n);
      const ur = ws.getUsedRange(true);
      ur.load(["address","text","rowIndex","columnIndex","isNullObject"]);
      await ctx.sync();
      if (ur.isNullObject) continue;

      for (let i = 0; i < ur.text.length && hits < 60; i++)
        for (let j = 0; j < ur.text[i].length && hits < 60; j++){
          if (String(ur.text[i][j]).toLowerCase().includes(needle.toLowerCase())){
            hits++;
            const addr = colName(ur.columnIndex + j) + (ur.rowIndex + i + 1);
            out += n + "!" + addr + "  =  " + String(ur.text[i][j]).slice(0,80) + "\n";
          }
        }
      if (hits >= 60) break;
    }
    return hits ? ("عدد النتائج: " + hits + "\n" + out) : ('ما لقيت "' + needle + '".');
  });
}

function colName(idx){
  let s = "";
  idx = Number(idx);
  while (idx >= 0){ s = String.fromCharCode(65 + (idx % 26)) + s; idx = Math.floor(idx / 26) - 1; }
  return s;
}

async function T_sort_range(a){
  return Excel.run(async ctx => {
    const ws = await resolveSheet(ctx, a.sheet);
    const rng = await getRangeOf(ctx, ws, a.address);
    await pushUndo(ctx, ws, rng, "فرز");
    rng.load("address");
    ws.load("name");
    await ctx.sync();

    let key = 0;
    if (a.key_column !== undefined && a.key_column !== null && a.key_column !== ""){
      key = isNaN(Number(a.key_column)) ? 0 : (Number(a.key_column) - 1);
      if (key < 0) key = 0;
    }
    const desc = String(a.order || "asc").toLowerCase() === "desc";
    const hasHeader = (a.has_header === undefined) ? true : !!a.has_header;

    rng.sort.apply([{ key: key, ascending: !desc }], false, hasHeader, "Rows");
    await ctx.sync();
    return "تم فرز " + ws.name + "!" + cleanAddr(rng.address) + " حسب العمود رقم " +
           (key + 1) + " (" + (desc ? "تنازلي" : "تصاعدي") + ").";
  });
}

async function T_remove_duplicates(a){
  return Excel.run(async ctx => {
    const ws = await resolveSheet(ctx, a.sheet);
    const rng = await getRangeOf(ctx, ws, a.address);
    await pushUndo(ctx, ws, rng, "حذف مكرر");
    rng.load(["address","rowCount","columnCount"]);
    ws.load("name");
    await ctx.sync();

    let cols;
    if (a.columns){
      cols = String(a.columns).split(",").map(x => Number(x.trim()) - 1).filter(n => n >= 0);
    } else {
      cols = Array.from({ length: rng.columnCount }, (_, i) => i);
    }
    const hasHeader = (a.has_header === undefined) ? true : !!a.has_header;

    const res = rng.removeDuplicates(cols, hasHeader);
    res.load(["removed","uniqueRemaining"]);
    await ctx.sync();
    return "تم حذف " + res.removed + " صف مكرر. الباقي " + res.uniqueRemaining + " صف فريد.";
  });
}

async function T_create_chart(a){
  return Excel.run(async ctx => {
    const ws = await resolveSheet(ctx, a.sheet);
    const rng = await getRangeOf(ctx, ws, a.data_address || a.address);
    rng.load("address");
    ws.load("name");
    await ctx.sync();

    const map = { line:"Line", pie:"Pie", bar:"BarClustered", scatter:"XYScatterLines",
                  area:"Area", column:"ColumnClustered" };
    const type = map[String(a.chart_type || "column").toLowerCase()] || "ColumnClustered";

    const ch = ws.charts.add(type, rng, "Auto");
    if (a.title){ ch.title.text = String(a.title); }
    ch.setPosition(a.anchor ? cleanAddr(a.anchor) : undefined, undefined);
    await ctx.sync();
    return "تم إنشاء رسم بياني (" + type + ") في الورقة " + ws.name +
           " من البيانات " + cleanAddr(rng.address) + ".";
  });
}

async function T_stats_summary(a){
  return Excel.run(async ctx => {
    const ws = await resolveSheet(ctx, a.sheet);
    const rng = await getRangeOf(ctx, ws, a.address);
    rng.load(["address","values","valueTypes","text"]);
    ws.load("name");
    await ctx.sync();

    const nums = [];
    let blanks = 0, texts = 0, textNums = 0, errs = 0;

    for (let i = 0; i < rng.values.length; i++)
      for (let j = 0; j < rng.values[i].length; j++){
        const t = rng.valueTypes[i][j], v = rng.values[i][j];
        if (t === "Error") errs++;
        else if (t === "Empty" || v === "") blanks++;
        else if (t === "Double" || t === "Integer") nums.push(Number(v));
        else if (t === "String"){
          texts++;
          if (!isNaN(Number(String(v).replace(/[,\s]/g,"")))) textNums++;
        }
      }

    let out = "ملخص " + ws.name + "!" + cleanAddr(rng.address) + "\n";
    out += "خلايا رقمية: " + nums.length + "\n";
    out += "المجموع الدقيق: " + numText(compensatedSum(nums)) + "\n";

    if (nums.length){
      const sorted = nums.slice().sort((x,y) => x - y);
      const mean = compensatedSum(nums) / nums.length;
      const mid = Math.floor(sorted.length / 2);
      const median = sorted.length % 2 ? sorted[mid] : (sorted[mid-1] + sorted[mid]) / 2;
      const varc = nums.length > 1
        ? compensatedSum(nums.map(x => (x - mean) * (x - mean))) / (nums.length - 1) : 0;
      out += "المتوسط: " + numText(mean) + "\n";
      out += "الوسيط: " + numText(median) + "\n";
      out += "أصغر: " + numText(sorted[0]) + "    أكبر: " + numText(sorted[sorted.length-1]) + "\n";
      out += "الانحراف المعياري: " + numText(Math.sqrt(varc)) + "\n";
    }
    out += "خلايا فاضية: " + blanks + "\nخلايا نصية: " + texts;
    if (textNums) out += "\n⚠ منها " + textNums + " رقم مخزّن كنص — ما بتدخل بالمجاميع! لازم تتصلّح.";
    out += "\nخلايا فيها خطأ: " + errs;
    return out;
  });
}

async function T_check_column(a){
  return Excel.run(async ctx => {
    const ws = await resolveSheet(ctx, a.sheet);
    const rng = await getRangeOf(ctx, ws, a.address);
    rng.load(["address","values","valueTypes","text","rowIndex","columnIndex"]);
    ws.load("name");
    await ctx.sync();

    const seen = new Map();
    let issues = 0, lines = "";

    for (let i = 0; i < rng.values.length; i++)
      for (let j = 0; j < rng.values[i].length; j++){
        const t = rng.valueTypes[i][j], v = rng.values[i][j];
        const addr = colName(rng.columnIndex + j) + (rng.rowIndex + i + 1);
        let note = "";

        if (t === "Error") note = "قيمة خطأ: " + rng.text[i][j];
        else if (t === "String" && v !== ""){
          const s = String(v);
          if (!isNaN(Number(s.replace(/[,\s]/g,"")))) note = 'رقم مخزّن كنص: "' + s + '"';
          else if (s !== s.trim() || s.includes("  ")) note = "مسافات زائدة في النص";
        }

        if (note){
          issues++;
          if (issues <= 40) lines += "  " + addr + " — " + note + "\n";
        }
        if (t !== "Empty" && v !== "" && t !== "Error"){
          const k = String(v);
          seen.set(k, (seen.get(k) || 0) + 1);
        }
      }

    let dups = 0;
    for (const n of seen.values()) if (n > 1) dups++;

    let head = "تدقيق " + ws.name + "!" + cleanAddr(rng.address) + "\n";
    head += "عدد المشاكل: " + issues + "    قيم مكررة: " + dups + "\n";
    if (issues > 40) head += "(معروض أول 40)\n";
    if (!issues && !dups) head += "ما في مشاكل ظاهرة.\n";
    return head + lines;
  });
}

/* --- البحث عن هدف: طريقة القاطع (secant) لأنه Office.js ما فيه GoalSeek --- */
async function T_goal_seek(a){
  const goal = Number(String(a.target_value).replace(/,/g,""));
  if (!isFinite(goal)) return "خطأ: target_value لازم يكون رقم.";
  if (!a.target_cell || !a.changing_cell) return "خطأ: لازم target_cell و changing_cell.";

  return Excel.run(async ctx => {
    const ws = await resolveSheet(ctx, a.sheet);
    const tgt = ws.getRange(cleanAddr(a.target_cell)).getCell(0,0);
    const chg = ws.getRange(cleanAddr(a.changing_cell)).getCell(0,0);

    tgt.load(["formulas","values"]);
    chg.load(["formulas","values"]);
    ws.load("name");
    await ctx.sync();

    const tf = tgt.formulas[0][0];
    if (typeof tf !== "string" || !tf.startsWith("="))
      return "خطأ: خلية الهدف لازم تكون فيها معادلة، مش رقم ثابت.";

    const original = chg.values[0][0];
    await pushUndo(ctx, ws, chg, "بحث عن هدف");

    const evalAt = async (x) => {
      chg.values = [[x]];
      await ctx.sync();
      tgt.load("values");
      await ctx.sync();
      const v = Number(tgt.values[0][0]);
      return isFinite(v) ? v : NaN;
    };

    let x0 = Number(original);
    if (!isFinite(x0) || x0 === 0) x0 = 1;
    let x1 = x0 * 1.1 + 1;

    let f0 = await evalAt(x0) - goal;
    let f1 = await evalAt(x1) - goal;
    let sol = NaN;

    for (let i = 0; i < 60; i++){
      if (!isFinite(f0) || !isFinite(f1)) break;
      if (Math.abs(f1) < 1e-9){ sol = x1; break; }
      const den = (f1 - f0);
      if (Math.abs(den) < 1e-15) break;
      const x2 = x1 - f1 * (x1 - x0) / den;
      if (!isFinite(x2)) break;
      x0 = x1; f0 = f1;
      x1 = x2; f1 = await evalAt(x1) - goal;
    }
    if (!isFinite(sol) && isFinite(f1) && Math.abs(f1) < 1e-6) sol = x1;

    if (!isFinite(sol)){
      chg.values = [[original]];
      await ctx.sync();
      return "ما قدرت ألاقي حل. تأكد إنه خلية الهدف بتعتمد فعلاً على الخلية المتغيّرة، " +
             "وإنه العلاقة بينهم مستمرة.";
    }

    chg.values = [[sol]];
    await ctx.sync();
    tgt.load("text");
    await ctx.sync();

    return "لقيت الحل: عشان " + cleanAddr(a.target_cell) + " تصير " + goal +
           " لازم " + cleanAddr(a.changing_cell) + " تكون " + numText(sol) +
           "  (كانت " + numText(Number(original)) + "). قيمة الهدف الحالية: " + tgt.text[0][0];
  });
}


/* ---------------------------------------------------------------------
   مخطط الأدوات للاستدعاء الأصلي (Function Calling)
   النموذج ما بيكتب JSON — المزوّد بيرجّع الاستدعاء منظّم ومضمون.
   ------------------------------------------------------------------- */
const SH = { type:"string", description:"اسم الورقة. اتركه فارغاً للورقة النشطة." };
const fn = (name, description, properties, required) => ({
  type:"function",
  function:{ name, description,
    parameters:{ type:"object", properties, required: required || [], additionalProperties:false } }
});

const TOOL_SCHEMA = [
  fn("list_sheets","استعراض كل أوراق الملف ونطاقاتها المستخدمة.",{}),
  fn("read_range","قراءة قيم ومعادلات نطاق. اتركه فارغاً لقراءة التحديد الحالي.",
     { sheet:SH, address:{type:"string",description:"مثل A1:F50"} }),
  fn("find_text","البحث عن نص وإرجاع عناوين الخلايا.",
     { text:{type:"string"}, sheet:SH }, ["text"]),
  fn("calc","حساب أي تعبير أو دالة إكسل بمحرّك إكسل نفسه. استخدمها لكل عملية حسابية بدون استثناء.",
     { expression:{type:"string",description:'بدون علامة يساوي، مثل SUMIFS(C2:C500,A2:A500,"مبيعات")'}, sheet:SH },
     ["expression"]),
  fn("precise_sum","جمع نطاق بدقة عالية مع تنبيه عن الأرقام المخزّنة كنص.",
     { sheet:SH, address:{type:"string"} }, ["address"]),
  fn("stats_summary","ملخص إحصائي: العدد، المجموع، المتوسط، الوسيط، الانحراف المعياري، وكم خلية رقمها مخزّن كنص.",
     { sheet:SH, address:{type:"string"} }, ["address"]),
  fn("check_column","تدقيق عمود: أرقام كنص، مسافات زائدة، قيم خطأ، مكررات.",
     { sheet:SH, address:{type:"string"} }, ["address"]),
  fn("goal_seek","البحث عن هدف: إيجاد القيمة اللازمة في خلية عشان معادلة بخلية تانية توصل لرقم معيّن.",
     { sheet:SH, target_cell:{type:"string",description:"خلية المعادلة مثل B10"},
       target_value:{type:"number"}, changing_cell:{type:"string",description:"الخلية المتغيّرة مثل B2"} },
     ["target_cell","target_value","changing_cell"]),
  fn("write_values","كتابة قيم. address هي الخلية الأولى فقط.",
     { sheet:SH, address:{type:"string",description:"الخلية الأولى مثل H1"},
       values:{type:"array",description:"مصفوفة صفوف",items:{type:"array",items:{}}} },
     ["address","values"]),
  fn("write_formula","كتابة معادلة على نطاق كامل مرة وحدة. المراجع النسبية بتتزحزح لكل صف تلقائياً.",
     { sheet:SH, address:{type:"string",description:"النطاق الكامل مثل E2:E420"},
       formula:{type:"string",description:"بأسماء إنجليزية مثل =C2*0.16"} },
     ["address","formula"]),
  fn("format_range","تنسيق نطاق: صيغة أرقام، ألوان، خط، محاذاة، حدود، ضبط عرض.",
     { sheet:SH, address:{type:"string"},
       number_format:{type:"string",description:"مثل #,##0.00 أو 0.0%"},
       bold:{type:"boolean"}, italic:{type:"boolean"}, font_size:{type:"number"},
       font_color:{type:"string",description:"#RRGGBB"}, fill:{type:"string",description:"#RRGGBB أو none"},
       align:{type:"string",description:"right | left | center"},
       wrap:{type:"boolean"}, borders:{type:"boolean"}, autofit:{type:"boolean"} },
     ["address"]),
  fn("sort_range","فرز نطاق حسب عمود.",
     { sheet:SH, address:{type:"string"}, key_column:{type:"number",description:"رقم العمود داخل النطاق من 1"},
       order:{type:"string",description:"asc | desc"}, has_header:{type:"boolean"} }, ["address"]),
  fn("remove_duplicates","حذف الصفوف المكررة.",
     { sheet:SH, address:{type:"string"}, columns:{type:"string",description:'مثل "1,2"'},
       has_header:{type:"boolean"} }, ["address"]),
  fn("add_sheet","إنشاء ورقة جديدة.", { name:{type:"string"} }, ["name"]),
  fn("trial_balance","فحص ميزان المراجعة: هل المدين = الدائن؟ وإذا في فرق بيقترح سببه.",
     { sheet:SH, debit_address:{type:"string"}, credit_address:{type:"string"} },
     ["debit_address","credit_address"]),
  fn("benford_analysis","تحليل بنفورد لكشف الأرقام المفبركة أو المُدخلة يدوياً. بدّه 50 قيمة على الأقل.",
     { sheet:SH, address:{type:"string"} }, ["address"]),
  fn("find_gaps","فحص تسلسل أرقام الفواتير/الشيكات/السندات: الأرقام الناقصة والمكررة.",
     { sheet:SH, address:{type:"string"} }, ["address"]),
  fn("reconcile","مطابقة عمودين من المبالغ (كشف بنك مع دفتر) وتحديد غير المتطابق في كل جهة.",
     { sheet:SH, address_a:{type:"string"}, address_b:{type:"string"},
       tolerance:{type:"number",description:"هامش المطابقة، افتراضي 0.01"} },
     ["address_a","address_b"]),
  fn("aging_analysis","تحليل أعمار الذمم: فئات 30/60/90/180/أكثر مع النسب والتنبيهات.",
     { sheet:SH, date_address:{type:"string"}, amount_address:{type:"string"},
       as_of:{type:"string",description:"تاريخ الاحتساب، افتراضي اليوم"} },
     ["date_address","amount_address"]),
  fn("duplicate_payments","كشف المدفوعات المكررة: نفس الجهة بنفس المبلغ أكثر من مرة.",
     { sheet:SH, key_address:{type:"string"}, amount_address:{type:"string"} },
     ["key_address","amount_address"]),
  fn("create_chart","إنشاء رسم بياني من نطاق.",
     { sheet:SH, data_address:{type:"string"},
       chart_type:{type:"string",description:"column|line|pie|bar|scatter|area"},
       title:{type:"string"}, anchor:{type:"string"} }, ["data_address"])
];


/* =====================================================================
   أدوات المدقق المالي المحترف
   تقنيات تدقيق فعلية — مش بس كتابة معادلات
   ===================================================================== */

async function readNums(ctx, sheet, address){
  const ws = await resolveSheet(ctx, sheet);
  const rng = await getRangeOf(ctx, ws, address);
  rng.load(["address","values","valueTypes","text","rowIndex","columnIndex"]);
  ws.load("name");
  await ctx.sync();
  return { ws, rng };
}

function flatNums(rng){
  const out = [];
  for (let i=0;i<rng.values.length;i++)
    for (let j=0;j<rng.values[i].length;j++){
      const t = rng.valueTypes[i][j];
      if (t === "Double" || t === "Integer") out.push(Number(rng.values[i][j]));
    }
  return out;
}

/* --- 1) ميزان المراجعة --- */
async function T_trial_balance(a){
  return Excel.run(async ctx => {
    const A = await readNums(ctx, a.sheet, a.debit_address);
    const B = await readNums(ctx, a.sheet, a.credit_address);
    const d = flatNums(A.rng), c = flatNums(B.rng);
    const sd = compensatedSum(d), sc = compensatedSum(c);
    const diff = sd - sc;

    let s = "ميزان المراجعة — " + A.ws.name + "\n";
    s += "مجموع المدين : " + numText(sd) + "   (" + d.length + " قيد)\n";
    s += "مجموع الدائن : " + numText(sc) + "   (" + c.length + " قيد)\n";
    s += "الفرق        : " + numText(diff) + "\n\n";

    if (Math.abs(diff) < 0.005) return s + "✔ الميزان مطابق.";

    s += "✖ الميزان غير مطابق بفرق " + numText(Math.abs(diff)) + "\n";
    s += "الفحوصات المقترحة بالترتيب:\n";
    if (Math.abs(Math.round(diff * 100)) % 900 === 0 || Math.abs(diff) % 9 === 0)
      s += "  • الفرق بيقبل القسمة على 9 ← غالباً خطأ تبديل أرقام (transposition)\n";
    s += "  • دوّر على قيد بقيمة " + numText(Math.abs(diff)/2) + " متسجّل بالجهة الغلط\n";
    s += "  • دوّر على قيد بقيمة " + numText(Math.abs(diff)) + " مفقود من جهة\n";
    s += "  • تأكد ما في أرقام مخزّنة كنص (شغّل stats_summary على العمودين)";
    return s;
  });
}

/* --- 2) تحليل بنفورد --- */
async function T_benford_analysis(a){
  return Excel.run(async ctx => {
    const { ws, rng } = await readNums(ctx, a.sheet, a.address);
    const nums = flatNums(rng).filter(n => n !== 0);
    const cnt = new Array(10).fill(0);
    let total = 0;
    for (const n of nums){
      let v = Math.abs(n);
      while (v < 1) v *= 10;
      while (v >= 10) v /= 10;
      const d = Math.floor(v);
      if (d >= 1 && d <= 9){ cnt[d]++; total++; }
    }
    if (total < 50)
      return "تحليل بنفورد بدّه 50 رقم على الأقل (المتاح: " + total + ").\n" +
             "التحليل على عيّنة صغيرة بيعطي إنذارات كاذبة.";

    let s = "تحليل بنفورد — " + ws.name + "!" + cleanAddr(rng.address) + "\n";
    s += "عدد القيم المفحوصة: " + total + "\n\nرقم | الفعلي % | المتوقع % | الانحراف\n";
    let chi = 0, maxDev = 0, maxDigit = 1;
    for (let i=1;i<=9;i++){
      const act = cnt[i]/total*100;
      const exp = Math.log10(1 + 1/i)*100;
      const dev = act - exp;
      chi += ((act-exp)*(act-exp))/exp;
      if (Math.abs(dev) > Math.abs(maxDev)){ maxDev = dev; maxDigit = i; }
      s += " " + i + "  |  " + act.toFixed(1) + "  |  " + exp.toFixed(1) + "  |  " +
           (dev>=0?"+":"") + dev.toFixed(1) + (Math.abs(dev)>5 ? "   ⚠" : "") + "\n";
    }
    s += "\n";
    if (chi < 15) s += "✔ التوزيع طبيعي ومتوافق مع قانون بنفورد. ما في مؤشر تلاعب.";
    else if (chi < 30) s += "⚠ في انحراف متوسط (أعلى انحراف عند الرقم " + maxDigit + ").\n" +
      "ممكن يكون طبيعي إذا البيانات فيها حدود سعرية ثابتة أو أرقام مدوّرة.";
    else s += "⚠⚠ انحراف كبير عن التوزيع الطبيعي (أعلى انحراف عند الرقم " + maxDigit + ").\n" +
      "هذا مؤشر يستحق فحص يدوي: ممكن أرقام مُدخلة يدوياً، حدود موافقة متحايَل عليها،\n" +
      "أو قيم مكررة. مش دليل تلاعب لحاله — بس نقطة بداية للفحص.";
    return s;
  });
}

/* --- 3) فجوات التسلسل --- */
async function T_find_gaps(a){
  return Excel.run(async ctx => {
    const { ws, rng } = await readNums(ctx, a.sheet, a.address);
    const seen = new Map();
    let mn = Infinity, mx = -Infinity, n = 0;
    for (const row of rng.text)
      for (const cell of row){
        const digits = String(cell).replace(/[^0-9]/g,"");
        if (!digits || digits.length > 15) continue;
        const v = Number(digits);
        seen.set(v, (seen.get(v)||0)+1);
        n++; if (v<mn) mn=v; if (v>mx) mx=v;
      }
    if (!n) return "ما لقيت أرقام تسلسلية في النطاق.";
    if (mx - mn > 100000)
      return "المدى واسع جداً (" + mn + " إلى " + mx + "). غالباً هذا مش عمود تسلسلي.";

    const gaps = [], dups = [];
    for (let i=mn;i<=mx;i++){
      if (!seen.has(i)) gaps.push(i);
      else if (seen.get(i) > 1) dups.push(i + "×" + seen.get(i));
    }
    let s = "فحص التسلسل — " + ws.name + "!" + cleanAddr(rng.address) + "\n";
    s += "المدى: من " + mn + " إلى " + mx + "   الموجود: " + n + "   المتوقع: " + (mx-mn+1) + "\n\n";
    if (!gaps.length) s += "✔ ما في أرقام ناقصة — التسلسل كامل.\n";
    else {
      s += "✖ ناقص " + gaps.length + " رقم:\n   " + gaps.slice(0,60).join(" ، ") +
           (gaps.length>60 ? "\n   ... (معروض أول 60)" : "") + "\n";
      s += "   الفجوات ممكن تعني: مستندات ملغاة، أو محذوفة، أو ما انسجلت.\n";
    }
    if (dups.length)
      s += "\n⚠ في " + dups.length + " رقم مكرر:\n   " + dups.slice(0,30).join(" ، ") +
           "\n   الترقيم المكرر خطر — ممكن فاتورة مسجّلة مرتين.";
    return s;
  });
}

/* --- 4) مطابقة عمودين --- */
async function T_reconcile(a){
  return Excel.run(async ctx => {
    const tol = Number(a.tolerance) > 0 ? Number(a.tolerance) : 0.01;
    const A = await readNums(ctx, a.sheet, a.address_a);
    const B = await readNums(ctx, a.sheet, a.address_b);

    const key = v => Math.round(v/tol).toString();
    const mapB = new Map();
    const cellsB = [];
    for (let i=0;i<B.rng.values.length;i++)
      for (let j=0;j<B.rng.values[i].length;j++){
        const t = B.rng.valueTypes[i][j];
        if (t!=="Double" && t!=="Integer") continue;
        const v = Number(B.rng.values[i][j]);
        const addr = colName(B.rng.columnIndex+j) + (B.rng.rowIndex+i+1);
        cellsB.push({v, addr});
        const k = key(v);
        if (!mapB.has(k)) mapB.set(k, []);
        mapB.get(k).push(addr);
      }

    const unA = []; let nA=0, matched=0, sumA=0;
    for (let i=0;i<A.rng.values.length;i++)
      for (let j=0;j<A.rng.values[i].length;j++){
        const t = A.rng.valueTypes[i][j];
        if (t!=="Double" && t!=="Integer") continue;
        const v = Number(A.rng.values[i][j]); nA++; sumA+=v;
        const k = key(v), lst = mapB.get(k);
        const addr = colName(A.rng.columnIndex+j) + (A.rng.rowIndex+i+1);
        if (lst && lst.length){ lst.shift(); matched++; }
        else unA.push("   " + addr + " = " + numText(v));
      }

    const sumB = compensatedSum(cellsB.map(c=>c.v));
    const unB = [];
    for (const [k,lst] of mapB) for (const addr of lst)
      unB.push("   " + addr + " = " + numText(Number(k)*tol));

    let s = "مطابقة " + cleanAddr(A.rng.address) + " مع " + cleanAddr(B.rng.address) + "\n";
    s += "الجهة أ: " + nA + " قيمة، مجموعها " + numText(sumA) + "\n";
    s += "الجهة ب: " + cellsB.length + " قيمة، مجموعها " + numText(sumB) + "\n";
    s += "فرق المجموع: " + numText(sumA - sumB) + "\n";
    s += "تطابق: " + matched + " قيمة   (بهامش " + tol + ")\n\n";
    if (!unA.length && !unB.length) return s + "✔ كل القيم متطابقة.";
    if (unA.length) s += "✖ موجود في (أ) وما له مقابل في (ب) — " + unA.length + ":\n" + unA.slice(0,50).join("\n") + "\n";
    if (unB.length) s += "✖ موجود في (ب) وما له مقابل في (أ) — " + unB.length + ":\n" + unB.slice(0,50).join("\n") + "\n";
    s += "\nملاحظة: قيمتين غير متطابقتين مجموعهما = فرق المجموع ← غالباً نفس البند بمبلغ مختلف.";
    return s;
  });
}

/* --- 5) أعمار الديون --- */
async function T_aging_analysis(a){
  return Excel.run(async ctx => {
    const D = await readNums(ctx, a.sheet, a.date_address);
    const M = await readNums(ctx, a.sheet, a.amount_address);
    const asOf = (a.as_of && !isNaN(Date.parse(a.as_of))) ? new Date(a.as_of) : new Date();

    const dates = [].concat(...D.rng.values);
    const amts  = [].concat(...M.rng.values);
    const types = [].concat(...M.rng.valueTypes);

    const b = [0,0,0,0,0], c = [0,0,0,0,0];
    let notDated = 0, total = 0;
    const rows = Math.min(dates.length, amts.length);

    for (let i=0;i<rows;i++){
      const t = types[i];
      if (t!=="Double" && t!=="Integer") continue;
      const amt = Number(amts[i]);
      let dt = null;
      if (typeof dates[i] === "number" && dates[i] > 1)
        dt = new Date(Date.UTC(1899,11,30) + dates[i]*86400000);
      else if (dates[i] && !isNaN(Date.parse(dates[i]))) dt = new Date(dates[i]);
      if (!dt){ notDated++; continue; }
      const days = Math.floor((asOf - dt)/86400000);
      const idx = days<=30?0 : days<=60?1 : days<=90?2 : days<=180?3 : 4;
      b[idx]+=amt; c[idx]++; total+=amt;
    }

    const nm = ["أقل من 30 يوم","31 - 60","61 - 90","91 - 180","أكثر من 180"];
    let s = "تحليل أعمار الديون — كما في " + asOf.toLocaleDateString("en-GB") + "\n";
    s += "الفئة | العدد | المبلغ | النسبة\n";
    for (let i=0;i<5;i++){
      const pct = total ? b[i]/total*100 : 0;
      s += nm[i] + " | " + c[i] + " | " + numText(b[i]) + " | " + pct.toFixed(1) + "%\n";
    }
    s += "الإجمالي: " + numText(total) + "\n";
    if (notDated) s += "\n⚠ " + notDated + " صف مبلغه موجود بس تاريخه ناقص أو غير صالح.\n";
    const over = b[3]+b[4];
    if (total && over/total > 0.25)
      s += "\n⚠ " + Math.round(over/total*100) + "٪ من الذمم أعمارها فوق 90 يوم — " +
           "نسبة مرتفعة تستدعي مخصص ديون مشكوك فيها.";
    return s;
  });
}

/* --- 6) المدفوعات المكررة --- */
async function T_duplicate_payments(a){
  return Excel.run(async ctx => {
    const K = await readNums(ctx, a.sheet, a.key_address);
    const M = await readNums(ctx, a.sheet, a.amount_address);
    const keys = [].concat(...K.rng.text);
    const amts = [].concat(...M.rng.values);
    const types = [].concat(...M.rng.valueTypes);
    const baseRow = K.rng.rowIndex + 1;

    const seen = new Map();
    const hits = []; let dupAmt = 0;
    const rows = Math.min(keys.length, amts.length);

    for (let i=0;i<rows;i++){
      const kv = String(keys[i]).trim();
      const t = types[i];
      if (!kv || (t!=="Double" && t!=="Integer")) continue;
      const av = Number(amts[i]);
      const k = kv.toLowerCase() + "#" + av.toFixed(2);
      if (seen.has(k)){
        hits.push("   صف " + (baseRow+i) + " يطابق صف " + seen.get(k) + "  —  " + kv + "  بمبلغ " + numText(av));
        dupAmt += av;
      } else seen.set(k, baseRow+i);
    }

    let s = "فحص المدفوعات المكررة — " + K.ws.name + "\nالصفوف المفحوصة: " + rows + "\n\n";
    if (!hits.length) return s + "✔ ما في مدفوعات مكررة (نفس الجهة + نفس المبلغ).";
    s += "⚠ لقيت " + hits.length + " حالة تكرار محتملة، مجموعها " + numText(dupAmt) + ":\n";
    s += hits.slice(0,40).join("\n") + (hits.length>40 ? "\n   ... (معروض أول 40)" : "") + "\n";
    s += "\nهاي مؤشرات مش أحكام — في حالات تكرار مشروع (دفعات شهرية ثابتة). " +
         "راجع رقم المستند والتاريخ لكل حالة.";
    return s;
  });
}

/* --- موزّع الأدوات --- */
const TOOLS = {
  list_sheets: T_list_sheets, read_range: T_read_range, calc: T_calc,
  precise_sum: T_precise_sum, write_values: T_write_values, write_formula: T_write_formula,
  format_range: T_format_range, add_sheet: T_add_sheet, find_text: T_find_text,
  sort_range: T_sort_range, remove_duplicates: T_remove_duplicates,
  create_chart: T_create_chart, stats_summary: T_stats_summary,
  check_column: T_check_column, goal_seek: T_goal_seek,
  trial_balance: T_trial_balance, benford_analysis: T_benford_analysis,
  find_gaps: T_find_gaps, reconcile: T_reconcile,
  aging_analysis: T_aging_analysis, duplicate_payments: T_duplicate_payments
};

const TOOL_LABEL = {
  list_sheets:"استعراض الأوراق", read_range:"قراءة البيانات", calc:"حساب دقيق",
  precise_sum:"جمع دقيق", write_values:"كتابة قيم", write_formula:"كتابة معادلة",
  format_range:"تنسيق", add_sheet:"إضافة ورقة", find_text:"بحث", sort_range:"فرز",
  remove_duplicates:"حذف المكرر", create_chart:"رسم بياني", stats_summary:"ملخص إحصائي",
  check_column:"تدقيق عمود", goal_seek:"بحث عن هدف",
  trial_balance:"ميزان المراجعة", benford_analysis:"تحليل بنفورد",
  find_gaps:"فحص التسلسل", reconcile:"مطابقة",
  aging_analysis:"أعمار الديون", duplicate_payments:"المدفوعات المكررة"
};

const MUTATING = new Set(["write_values","write_formula","format_range","sort_range",
                          "remove_duplicates","add_sheet","create_chart","goal_seek"]);

async function runTool(name, args){
  const fn = TOOLS[name];
  if (!fn) return "خطأ: أداة غير معروفة (" + name + ").";
  try { return await fn(args || {}); }
  catch(e){
    const m = (e && e.message) ? e.message : String(e);
    return "خطأ أثناء تنفيذ " + name + ": " + m;
  }
}

/* =====================================================================
   5) الوكيل: التعليمات، استخراج الأدوات، والحلقة
   ===================================================================== */


const NATIVE_PROMPT = `أنت "مساعد المحاسب" — مساعد ذكي مدمج داخل مايكروسوفت إكسل، بتشتغل مع محاسب مالي عربي.
شخصيتك: خبير محاسبة ومالية ورياضيات، دقيق جداً بالأرقام، عملي، وبتحكي عربي بسيط ومباشر.

== قواعد أساسية ==
1. كل ردودك النهائية بالعربي. أسماء دوال إكسل ومراجع الخلايا تبقى إنجليزي كما هي.
2. ممنوع تحسب أي عملية حسابية بعقلك. أي جمع/طرح/ضرب/قسمة/نسبة/فائدة → استخدم أداة calc.
3. ما تفترض شكل البيانات. اقرأ الشيت أولاً بـ read_range أو list_sheets قبل ما تعدّل.
4. المعادلات بأسماء إنجليزية وبفاصلة عادية: =SUMIFS(C:C,A:A,"x")
5. لما تخلّص، اشرح للمحاسب باختصار شو عملت ووين، وإذا في معادلة اشرح منطقها بسطر.
6. إذا الطلب ناقص أو غامض بشكل بيأثر على النتيجة، اسأل سؤال واحد واضح بدل ما تخمّن.

== الأدوات ==
عندك أدوات جاهزة بتشتغل على الشيت مباشرة — استدعِها بالطريقة الرسمية (tool calling).
• استدعِ أداة وحدة بكل خطوة وانتظر نتيجتها قبل اللي بعدها.
• ممنوع تقول "تمّ" أو "أضفت" أو "حسبت" قبل ما توصلك نتيجة الأداة فعلياً.
• نتيجة read_range ممكن تكون مقتطعة — اعتمد على عدد الصفوف المذكور في أول سطر،
  مش على عدد الصفوف المعروضة، لما تحدد نطاق المعادلة.

== دليل المدقق المحترف ==
أنت مش بس بتكتب معادلات — أنت بتفكّر زي مدقق عنده 10 سنين خبرة.
لما يعطيك المحاسب بيانات، اسأل حالك: وين ممكن يكون الخطأ؟ واستخدم الأداة المناسبة:
• "الميزان ما بيوزن" → trial_balance. بيقلك الفرق ويقترح سببه
  (فرق بيقبل القسمة على 9 = خطأ تبديل أرقام، ونص الفرق = قيد بالجهة الغلط).
• "في أرقام مشبوهة" أو تدقيق مصاريف → benford_analysis.
• فواتير أو شيكات أو سندات → find_gaps.
• "طابق كشف البنك مع الدفتر" → reconcile.
• ذمم مدينة → aging_analysis.
• مصاريف أو مدفوعات موردين → duplicate_payments.
• قبل أي مجموع → stats_summary لكشف الأرقام المخزّنة كنص.

عادات المدقق المحترف:
- ما تقبل رقم بدون ما تتحقق من مصدره.
- لما تلاقي مشكلة، قل حجمها بالأرقام (كم صف، كم مبلغ، كم نسبة) مش بس "في مشكلة".
- فرّق بين "مؤشر يستحق الفحص" و"خطأ مؤكد". ما تتهم بدون دليل.
- رتّب النتائج حسب الأثر المالي، الأكبر أولاً.
- إذا الطلب عام مثل "دقّق الشيت"، شغّل أكتر من أداة فحص وبعدين لخّص.

== أسلوب الشغل الصحيح ==
- اقرأ → احسب بـ calc → اكتب → نسّق → لخّص.
- المعادلة تُكتب على النطاق كامل مرة وحدة (مثل E2:E420) مش خلية خلية.
- بعد أي write_formula، اقرأ النتيجة اللي رجعتلك وتأكد إنها منطقية (مش #REF! ولا #VALUE!).
- قبل ما تعتمد على مجموع عمود، شغّل stats_summary — إذا في أرقام مخزّنة كنص
  المجموع بيطلع ناقص وأنت ما بتحس. نبّه المحاسب واقترح تصليحها.
- في أسئلة "شو لازم يكون X عشان Y يصير كذا" استخدم goal_seek مش التخمين.
- في الشغل المحاسبي: انتبه للتقريب، للنصوص اللي شكلها أرقام، وللصفوف الفاضية جوّا البيانات.
- ملاحظة: هذه النسخة ما بتقدر تشغّل ماكرو VBA. استخدم الأدوات الموجودة فقط.`;

function systemPrompt(native){
  if (native) return NATIVE_PROMPT;
  return `أنت "مساعد المحاسب" — مساعد ذكي مدمج داخل مايكروسوفت إكسل، بتشتغل مع محاسب مالي عربي.
شخصيتك: خبير محاسبة ومالية ورياضيات، دقيق جداً بالأرقام، عملي، وبتحكي عربي بسيط ومباشر.

== قواعد أساسية ==
1. كل ردودك النهائية بالعربي. أسماء دوال إكسل ومراجع الخلايا تبقى إنجليزي كما هي.
2. ممنوع تحسب أي عملية حسابية بعقلك. أي جمع/طرح/ضرب/قسمة/نسبة/فائدة → استخدم أداة calc.
3. ما تفترض شكل البيانات. اقرأ الشيت أولاً بـ read_range أو list_sheets قبل ما تعدّل.
4. المعادلات بأسماء إنجليزية وبفاصلة عادية: =SUMIFS(C:C,A:A,"x")
5. لما تخلّص، اشرح للمحاسب باختصار شو عملت ووين، وإذا في معادلة اشرح منطقها بسطر.
6. إذا الطلب ناقص أو غامض بشكل بيأثر على النتيجة، اسأل سؤال واحد واضح بدل ما تخمّن.

== كيف تستدعي أداة (اقرأ هذا القسم بعناية) ==
لما تحتاج تشوف أو تعدّل الشيت، ارجع بهذا الشكل بالضبط ولا شي غيره:
<<<TOOL>>>
{"tool":"read_range","args":{"sheet":"اليومية","address":"A1:D420"}}
<<<END>>>

قواعد صارمة ما فيها استثناء:
• المحتوى بين العلامتين لازم يكون JSON صحيح فيه مفتاحين بالضبط: "tool" و "args".
• "args" لازم يكون كائن بأسماء الحقول، مش قائمة قيم:
    صح  : {"tool":"read_range","args":{"sheet":"اليومية","address":"A1:D420"}}
    غلط : read_range {"اليومية", "A1:D420"}
    غلط : {"tool":"read_range","args":["اليومية","A1:D420"]}
• أداة واحدة فقط في الرد. ممنوع كتلتين أو أكثر.
• ممنوع أي كلام قبل <<<TOOL>>> أو بعد <<<END>>>.
• بعد ما ترسل الأداة، توقّف وانتظر. نتيجتها رح توصلك في الرسالة الجاية، وبعدها ترسل التالية.
• ممنوع تقول "تمّ" أو "أضفت" أو "حسبت" قبل ما توصلك نتيجة الأداة فعلياً.
  نتيجة الأداة هي الدليل الوحيد إنه الشي انعمل. بدونها ما صار شي على الملف.

مثال كامل على مهمة من عدة خطوات:
  أنت   : <<<TOOL>>>{"tool":"read_range","args":{"sheet":"اليومية","address":"A1:D5"}}<<<END>>>
  النظام: نتيجة الأداة [read_range]: ... البيانات ...
  أنت   : <<<TOOL>>>{"tool":"write_values","args":{"sheet":"اليومية","address":"E1","values":[["ضريبة 16%"]]}}<<<END>>>
  النظام: نتيجة الأداة [write_values]: تم كتابة 1 صف × 1 عمود في اليومية!E1.
  أنت   : <<<TOOL>>>{"tool":"write_formula","args":{"sheet":"اليومية","address":"E2:E420","formula":"=C2*0.16"}}<<<END>>>
  النظام: نتيجة الأداة [write_formula]: تم وضع المعادلة ... أول نتيجة: 200
  أنت   : خلصت. ضفت عمود الضريبة في العمود E والمجموع في E421.

لما تخلص كل الخطوات فعلاً، ارجع بنص عربي عادي بدون <<<TOOL>>> — وهذا ردك النهائي.

== الأدوات المتاحة ==
read_range   {sheet, address}           قراءة نطاق. address فاضي = التحديد الحالي.
list_sheets  {}                          أسماء الأوراق ونطاقاتها المستخدمة.
calc         {expression, sheet}         حساب دقيق بمحرّك إكسل. بيقبل أي دالة إكسل ومراجع خلايا.
                                         مثال: {"expression":"SUMIFS(C2:C500,A2:A500,\\"مبيعات\\")"}
precise_sum  {sheet, address}            جمع تعويضي دقيق + بينبّه على الأرقام المخزّنة كنص.
stats_summary{sheet, address}            العدد، المجموع، المتوسط، الوسيط، الانحراف المعياري،
                                         وكم خلية فيها رقم مخزّن كنص (سبب رقم 1 لاختلاف المجاميع).
check_column {sheet, address}            تدقيق: أرقام كنص، مسافات زائدة، قيم خطأ، مكررات.
write_values {sheet, address, values}    values مصفوفة صفوف: [["البيان","المبلغ"],["إيجار",1200]]
                                         address = الخلية الأولى فقط (مثل "H1").
write_formula{sheet, address, formula}   على النطاق كامل مرة وحدة: address "E2:E420".
format_range {sheet, address, number_format, bold, italic, font_size, font_color,
              fill, align, wrap, borders, autofit}   الألوان "#RRGGBB".
add_sheet    {name}
find_text    {text, sheet}
sort_range   {sheet, address, key_column, order:asc|desc, has_header}
remove_duplicates {sheet, address, columns:"1,2", has_header}
create_chart {sheet, data_address, chart_type:column|line|pie|bar|scatter|area, title, anchor}
goal_seek    {sheet, target_cell, target_value, changing_cell}
             مثال: شو لازم يكون سعر البيع (B2) عشان صافي الربح (B10) يصير 25000.
             شرط: خلية الهدف لازم تكون معادلة بتعتمد على الخلية المتغيّرة.

== دليل المدقق المحترف ==
أنت مش بس بتكتب معادلات — أنت بتفكّر زي مدقق عنده 10 سنين خبرة.
لما يعطيك المحاسب بيانات، اسأل حالك: وين ممكن يكون الخطأ؟ واستخدم الأداة المناسبة:
• "الميزان ما بيوزن" → trial_balance. بيقلك الفرق ويقترح سببه
  (فرق بيقبل القسمة على 9 = خطأ تبديل أرقام، ونص الفرق = قيد بالجهة الغلط).
• "في أرقام مشبوهة" أو تدقيق مصاريف → benford_analysis.
• فواتير أو شيكات أو سندات → find_gaps.
• "طابق كشف البنك مع الدفتر" → reconcile.
• ذمم مدينة → aging_analysis.
• مصاريف أو مدفوعات موردين → duplicate_payments.
• قبل أي مجموع → stats_summary لكشف الأرقام المخزّنة كنص.

عادات المدقق المحترف:
- ما تقبل رقم بدون ما تتحقق من مصدره.
- لما تلاقي مشكلة، قل حجمها بالأرقام (كم صف، كم مبلغ، كم نسبة) مش بس "في مشكلة".
- فرّق بين "مؤشر يستحق الفحص" و"خطأ مؤكد". ما تتهم بدون دليل.
- رتّب النتائج حسب الأثر المالي، الأكبر أولاً.
- إذا الطلب عام مثل "دقّق الشيت"، شغّل أكتر من أداة فحص وبعدين لخّص.

== أسلوب الشغل الصحيح ==
- اقرأ → احسب بـ calc → اكتب → نسّق → لخّص.
- المعادلة تُكتب على النطاق كامل مرة وحدة، مش خلية خلية.
- بعد أي write_formula، اقرأ النتيجة اللي رجعتلك وتأكد إنها منطقية (مش #REF! ولا #VALUE!).
- قبل ما تعتمد على مجموع عمود، شغّل stats_summary — إذا في أرقام مخزّنة كنص المجموع بيطلع ناقص.
- في أسئلة "شو لازم يكون X عشان Y يصير كذا" استخدم goal_seek مش التخمين.
- في الشغل المحاسبي: انتبه للتقريب، للنصوص اللي شكلها أرقام، وللصفوف الفاضية جوّا البيانات.
- ملاحظة: هذه النسخة ما بتقدر تشغّل ماكرو VBA. استخدم الأدوات الموجودة فقط.`;
}

/* --- استخراج متسامح لاستدعاء الأداة --- */
const ACT = { NONE:0, OK:1, BAD:2 };

function extractBody(s){
  const m = /<<<\s*TOOL/i.exec(s);
  if (!m) return "";
  let q = m.index + m[0].length;
  while (q < s.length && ">\t\r\n ".includes(s[q])) q++;
  const e = s.toLowerCase().indexOf("<<<end", q);
  return s.slice(q, e < 0 ? s.length : e);
}

function countMarkers(s){ return (s.match(/<<<\s*TOOL/gi) || []).length; }

function looksLikeToolAttempt(s){
  if (/<<<\s*TOOL/i.test(s)) return true;
  for (const n of Object.keys(TOOLS)){
    const p = s.indexOf(n);
    if (p >= 0){
      const b = s.indexOf("{", p);
      if (b >= 0 && b - p <= 40) return true;
    }
  }
  return false;
}

function parseAction(txt){
  let body = extractBody(txt);
  if (!body.trim()){
    const fence = /```(?:json)?\s*([\s\S]*?)```/.exec(txt);
    if (fence && fence[1].includes('"tool"')) body = fence[1];
  }
  if (!body.trim())
    return looksLikeToolAttempt(txt) ? { v:ACT.BAD, raw: txt.slice(0,300) } : { v:ACT.NONE };

  const raw = body.trim().slice(0,300);
  const i = body.indexOf("{"), j = body.lastIndexOf("}");
  if (i < 0 || j <= i) return { v:ACT.BAD, raw };

  try{
    const o = JSON.parse(body.slice(i, j + 1));
    if (!o || !o.tool) return { v:ACT.BAD, raw };
    if (o.args && (Array.isArray(o.args) || typeof o.args !== "object")) return { v:ACT.BAD, raw };
    return { v:ACT.OK, tool: String(o.tool).trim(), args: o.args || {} };
  }catch(e){ return { v:ACT.BAD, raw }; }
}

const badFormatMsg = raw =>
`خطأ: صيغة استدعاء الأداة غلط، فما انعمل ولا شي على الملف.

اللي بعتّه:
${raw}

الصيغة الصحيحة، كتلة وحدة بس، JSON صحيح، بمفتاحين "tool" و "args"،
و "args" كائن بأسماء الحقول مش قائمة قيم:

<<<TOOL>>>
{"tool":"read_range","args":{"sheet":"Sheet1","address":"A1:D50"}}
<<<END>>>

أعد إرسال أول أداة بس بالصيغة الصحيحة، بدون أي كلام قبلها أو بعدها، وانتظر نتيجتها.`;

/* --- سياق الملف اللي بينشاف للنموذج --- */
async function buildContext(){
  try{
    return await Excel.run(async ctx => {
      const wb = ctx.workbook; wb.load("name");
      const sheets = ctx.workbook.worksheets; sheets.load("items/name");
      const ws = sheets.getActiveWorksheet(); ws.load("name");
      const sel = ctx.workbook.getSelectedRange(); sel.load(["address","text","formulas","cellCount"]);
      const ur = ws.getUsedRange(true);
      ur.load(["address","rowCount","columnCount","isNullObject"]);
      await ctx.sync();

      let out = "[سياق الملف الحالي]\n";
      out += "الملف: " + wb.name + "\n";
      out += "الأوراق: " + sheets.items.map(s => s.name).join(" ، ") + "\n";
      out += "الورقة النشطة: " + ws.name;

      if (!ur.isNullObject){
        out += "   النطاق المستخدم: " + cleanAddr(ur.address) +
               "  (" + ur.rowCount + " صف × " + ur.columnCount + " عمود)\n";
        try{
          const prev = ws.getUsedRange(true).getAbsoluteResizedRange(
            Math.min(ur.rowCount, 4), Math.min(ur.columnCount, 12));
          prev.load("text");
          await ctx.sync();
          out += "معاينة أول صفوف:\n";
          for (const row of prev.text)
            out += "  " + row.map(c => String(c).slice(0,22)).join(" | ") + "\n";
        }catch(e){}
      } else { out += "   (الورقة فاضية)\n"; }

      out += "التحديد الحالي: " + cleanAddr(sel.address);
      if (sel.cellCount === 1){
        out += "\n  محتوى الخلية: " + String(sel.text[0][0]).slice(0,200);
        const f = sel.formulas[0][0];
        if (typeof f === "string" && f.startsWith("=")) out += "\n  المعادلة: " + f;
      }
      return out + "\n";
    });
  }catch(e){ return "[تعذّر قراءة سياق الملف: " + (e.message || e) + "]"; }
}

/* --- الحلقة --- */
let HIST = [];
let BUSY = false;

async function agentAsk(userText){
  const cfg = CFG.read();
  const maxSteps = Math.max(3, Math.min(40, Number(cfg.steps) || 14));
  const prov = PROVIDERS[cfg.provider] || PROVIDERS.custom;

  let native = (cfg.nativeTools !== false) && prov.tools !== false;

  const ctxBlock = await buildContext();
  HIST.push({ role:"user", content: ctxBlock + "\n\nطلب المستخدم:\n" + userText });
  while (HIST.length > 40) HIST.shift();

  let bad = 0;

  for (let step = 1; step <= maxSteps; step++){
    setStatus("أفكّر…  (خطوة " + step + " من " + maxSteps + ")");

    const rep = await aiChat(systemPrompt(native), HIST, native ? TOOL_SCHEMA : null);

    if (!rep.ok){
      if (rep.err === "NOTOOLS" && native){
        native = false;
        addStep("النموذج ما بيدعم الأدوات الأصلية", "بحوّل للوضع النصي", "bad");
        continue;
      }
      return { error: rep.err === "NOTOOLS" ? ("الطلب مرفوض من المزوّد: " + (rep.detail||"")) : rep.err };
    }

    /* ============ الوضع الأصلي: Function Calling ============ */
    if (native){
      const calls = rep.toolCalls || [];

      if (!calls.length) return { text: (rep.text || "").trim() };

      HIST.push({ role:"assistant", content: rep.text || null, tool_calls: calls });

      for (const c of calls){
        const name = c.function && c.function.name;
        const label = TOOL_LABEL[name] || name;

        let args = null, parseErr = "";
        try { args = JSON.parse((c.function && c.function.arguments) || "{}"); }
        catch(e){ parseErr = String((c.function && c.function.arguments) || "").slice(0,150); }

        if (cfg.confirm && MUTATING.has(name) && args){
          const okGo = confirm("المساعد بدّه ينفّذ: " + label + "\n\n" +
                               JSON.stringify(args, null, 1).slice(0,700) + "\n\nبتوافق؟");
          if (!okGo){
            HIST.push({ role:"tool", tool_call_id:c.id,
              content:"المستخدم رفض تنفيذ هذه الخطوة. اسأله شو بدّه بالضبط أو اقترح طريقة تانية." });
            continue;
          }
        }

        const el = addStep(label, JSON.stringify(args || {}).slice(0,120), "run");
        setStatus("أنفّذ: " + label + " …");

        let res;
        if (args === null){
          res = "خطأ: تعذّر قراءة معاملات الأداة.\nاللي وصلني: " + parseErr +
                "\nأعد إرسال الاستدعاء بمعاملات صحيحة.";
        } else {
          res = await runTool(name, args);
        }
        if (res.length > 12000) res = res.slice(0,12000) + "\n…(مقتطع)";

        const failed = res.startsWith("خطأ") || res.includes("⚠ الخلية رجّعت خطأ");
        markStep(el, failed ? "bad" : "ok");

        HIST.push({ role:"tool", tool_call_id:c.id, content:res });
      }
      continue;
    }

    /* ============ الوضع النصي الاحتياطي ============ */
    HIST.push({ role:"assistant", content: rep.text });
    const act = parseAction(rep.text);

    if (act.v === ACT.OK){
      bad = 0;
      const label = TOOL_LABEL[act.tool] || act.tool;

      if (cfg.confirm && MUTATING.has(act.tool)){
        const okGo = confirm("المساعد بدّه ينفّذ: " + label + "\n\n" +
                             JSON.stringify(act.args, null, 1).slice(0,700) + "\n\nبتوافق؟");
        if (!okGo){
          HIST.push({ role:"user", content:"المستخدم رفض تنفيذ " + act.tool + "." });
          continue;
        }
      }

      const el = addStep(label, JSON.stringify(act.args).slice(0,120), "run");
      setStatus("أنفّذ: " + label + " …");
      let res = await runTool(act.tool, act.args);
      if (res.length > 12000) res = res.slice(0,12000) + "\n…(مقتطع)";
      markStep(el, (res.startsWith("خطأ") ? "bad" : "ok"));

      let follow = "نتيجة الأداة [" + act.tool + "]:\n" + res;
      if (countMarkers(rep.text) > 1){
        follow += "\n\nتنبيه: بعتّ أكتر من كتلة أداة. نفّذت الأولى فقط — ابعت الباقي وحدة وحدة.";
      }
      follow += "\n\nالآن ابعت الخطوة التالية: كتلة أداة واحدة فقط، أو ردك النهائي بالعربي إذا خلصت.";
      HIST.push({ role:"user", content: follow });

    } else if (act.v === ACT.BAD){
      bad++;
      if (bad >= 3){
        return { text: "⚠ النموذج مش ملتزم بصيغة استدعاء الأدوات، فما انعمل شي على الملف.\n\n" +
                       "آخر محاولة منه:\n" + act.raw + "\n\n" +
                       "الحل الأفضل: بدّل المزوّد لـ OmniRoute من الإعدادات — بتدعم استدعاء " +
                       "الأدوات الأصلي وما بتصير هالمشكلة نهائياً." };
      }
      addStep("تصحيح صيغة الاستدعاء", "محاولة " + bad + " من 3", "bad");
      HIST.push({ role:"user", content: badFormatMsg(act.raw) });

    } else {
      return { text: rep.text.trim() };
    }
  }

  return { text: "وصلت للحد الأقصى من الخطوات (" + maxSteps + ") ولسا ما خلّصت.\n" +
                 "جرّب تقسّم الطلب لخطوات أصغر، أو زيد عدد الخطوات من الإعدادات." };
}

/* =====================================================================
   6) الواجهة
   ===================================================================== */

const logEl = () => $("log");

function scrollDown(){ const l = logEl(); l.scrollTop = l.scrollHeight; }

function addMsg(text, cls, who){
  const d = document.createElement("div");
  d.className = "msg " + (cls || "bot");
  if (who) d.innerHTML = '<span class="who">' + esc(who) + "</span>" + esc(text);
  else d.textContent = text;
  logEl().appendChild(d);
  scrollDown();
  return d;
}

function addStep(label, detail, state){
  const d = document.createElement("div");
  d.className = "step " + (state === "run" ? "run" : state === "bad" ? "bad" : "");
  d.innerHTML = '<span class="dot"></span><span>' + esc(label) +
                '</span><code>' + esc(detail || "") + "</code>";
  logEl().appendChild(d);
  scrollDown();
  return d;
}

function markStep(el, state){
  if (!el) return;
  el.classList.remove("run");
  if (state === "bad") el.classList.add("bad");
}

function setStatus(s){ $("status").textContent = s || ""; }

function refreshUndo(){
  const b = $("btnUndo");
  b.disabled = UNDO.length === 0;
  b.textContent = UNDO.length ? "تراجع (" + UNDO.length + ")" : "تراجع";
}

function setBusy(b){
  BUSY = b;
  $("send").disabled = b;
  $("send").textContent = b ? "…" : "إرسال";
  $("input").disabled = b;
  $("btnNew").disabled = b;
  document.querySelectorAll("#quick button").forEach(x => x.disabled = b);
  if (!b) refreshUndo();
}

function welcome(){
  const d = document.createElement("div");
  d.className = "msg bot welcome";
  d.innerHTML =
    "<b>أهلاً 👋</b> اكتب طلبك بالعربي وأنا بنفّذه على الملف المفتوح." +
    "<ul>" +
    "<li>اعملي عمود ضريبة ١٦٪ جنب المبلغ</li>" +
    "<li>طابق عمود D مع H وقلي وين الفرق</li>" +
    "<li>ليش هاي المعادلة بترجع <bdi>#VALUE!</bdi> ؟</li>" +
    "<li>شو لازم يكون سعر البيع عشان الربح يصير ٢٥٠٠٠</li>" +
    "</ul>";
  logEl().appendChild(d);
}

async function send(text){
  if (BUSY) return;
  const q = (text !== undefined ? text : $("input").value).trim();
  if (!q) return;

  if (!CFG.ok()){
    openCfg();
    $("cfgStatus").innerHTML = '<span class="bad">لازم تحط المفتاح أولاً، بعدها اضغط حفظ.</span>';
    return;
  }

  addMsg(q, "user", "أنت");
  $("input").value = "";
  setBusy(true);

  let out;
  try { out = await agentAsk(q); }
  catch(e){ out = { error: (e && e.message) ? e.message : String(e) }; }

  setStatus("");
  if (out.error) addMsg("⚠ " + out.error, "err", "خطأ");
  else addMsg(out.text, "bot", "المساعد");

  setBusy(false);
  $("input").focus();
}

/* --- الإعدادات --- */
function fillProviders(){
  const sel = $("cfgProvider");
  if (sel.options.length) return;
  for (const k of Object.keys(PROVIDERS)){
    const o = document.createElement("option");
    o.value = k; o.textContent = PROVIDERS[k].label;
    sel.appendChild(o);
  }
  sel.onchange = () => {
    const p = PROVIDERS[sel.value];
    $("cfgProvHint").textContent = p.hint || "";
    if (sel.value !== "custom"){ $("cfgUrl").value = p.url; $("cfgModel").value = p.model; }
    $("cfgKey").disabled = !p.needsKey;
    $("cfgKey").placeholder = p.needsKey ? "sk-…" : "مش مطلوب لهذا المزوّد";
    $("cfgUrl").readOnly = (sel.value !== "custom");
  };
}

function openCfg(){
  const c = CFG.read();
  fillProviders();
  $("cfgProvider").value = c.provider || "omniroute";
  $("cfgProvider").onchange();
  $("cfgUrl").value   = c.url;
  $("cfgKey").value   = c.key;
  $("cfgModel").value = c.model;
  $("cfgSteps").value = c.steps;
  $("cfgConfirm").checked = !!c.confirm;
  $("cfgStatus").textContent = "—";
  $("settings").classList.add("open");
}

function closeCfg(){ $("settings").classList.remove("open"); }

function readCfgFields(){
  return {
    provider: $("cfgProvider").value,
    nativeTools: true,
    url:   $("cfgUrl").value.trim(),
    key:   $("cfgKey").value.trim(),
    model: $("cfgModel").value.trim(),
    steps: Math.max(3, Math.min(40, Number($("cfgSteps").value) || 14)),
    confirm: $("cfgConfirm").checked
  };
}

function updateChip(){
  const c = CFG.read();
  $("modelChip").textContent = c.model || "—";
  $("modelChip").title = c.model || "";
}

/* --- الربط --- */
Office.onReady(info => {
  if (info.host !== Office.HostType.Excel){
    document.body.innerHTML =
      '<div style="padding:20px">هذه الإضافة بتشتغل داخل مايكروسوفت إكسل فقط.</div>';
    return;
  }

  updateChip();
  welcome();
  refreshUndo();
  if (!CFG.ok()) setStatus("⚙ افتح الإعدادات لاختيار المزوّد.");

  $("send").onclick = () => send();
  $("input").addEventListener("keydown", e => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)){ e.preventDefault(); send(); }
  });

  $("btnUndo").onclick = async () => {
    const m = await undoLast();
    addStep("تراجع", m, UNDO.length >= 0 ? "ok" : "bad");
    refreshUndo();
  };

  $("btnNew").onclick = () => {
    if (BUSY) return;
    HIST = []; UNDO = [];
    logEl().innerHTML = "";
    welcome();
    refreshUndo();
    setStatus("");
  };

  document.querySelectorAll("#quick button").forEach(b => {
    b.onclick = () => {
      const q = b.dataset.q;
      if (q === "audit")
        send("دقّق الورقة النشطة كمراجع مالي: شغّل stats_summary و check_column على أعمدة الأرقام، " +
             "ودوّر على أخطاء المعادلات وأرقام مخزّنة كنص وصفوف فاضية جوّا البيانات ومجاميع ما بتطابق تفاصيلها. " +
             "بالآخر اعطيني تقرير مختصر بالمشاكل مرتبة حسب الخطورة مع مكان كل مشكلة.");
      else if (q === "explain")
        send("اشرحلي بالتفصيل شو بتعمل الخلايا المحددة وشو منطق المعادلة، وإذا في فيها خطأ أو خطر قوللي.");
      else if (q === "stats")
        send("اعطيني ملخص إحصائي للنطاق المحدد، ونبّهني إذا في أرقام مخزّنة كنص.");
      else if (q === "audit2")
        send("شغّل فحص تدقيق شامل على الورقة: ابدأ بـ list_sheets و read_range لتفهم الأعمدة، " +
             "بعدين شغّل الأدوات المناسبة حسب نوع البيانات (trial_balance إذا في مدين ودائن، " +
             "benford_analysis على عمود المبالغ، find_gaps إذا في أرقام تسلسلية، " +
             "duplicate_payments إذا في موردين ومبالغ، stats_summary على كل عمود أرقام). " +
             "بالآخر اعطيني تقرير مرتب حسب الأثر المالي، الأكبر أولاً.");
      else { $("input").value = "بدي معادلة تعمل: "; $("input").focus(); }
    };
  });

  $("btnCfg").onclick = openCfg;
  $("btnCloseCfg").onclick = closeCfg;

  $("btnSaveCfg").onclick = () => {
    CFG.write(readCfgFields());
    updateChip();
    closeCfg();
    setStatus("تم حفظ الإعدادات.");
  };

  $("btnTest").onclick = async () => {
    CFG.write(readCfgFields());
    updateChip();
    const s = $("cfgStatus");
    s.textContent = "… بجرّب الاتصال";
    const r = await aiChat("أنت مساعد. رد بكلمة واحدة.", [{ role:"user", content:"رد بكلمة: تمام" }]);
    s.innerHTML = r.ok
      ? '<span class="ok">✔ الاتصال شغّال والنموذج بيرد.</span>'
      : '<span class="bad">✖ ' + esc(r.err) + "</span>";
  };

  $("btnModels").onclick = async () => {
    CFG.write(readCfgFields());
    const s = $("cfgStatus");
    s.textContent = "… بجيب قائمة النماذج";
    try{
      const ids = await listModels();
      if (!ids.length){ s.innerHTML = '<span class="bad">ما رجعت أي نماذج.</span>'; return; }
      const sel = $("cfgModelList");
      sel.innerHTML = "";
      for (const id of ids){
        const o = document.createElement("option");
        o.value = id; o.textContent = id;
        sel.appendChild(o);
      }
      sel.value = $("cfgModel").value || ids[0];
      sel.style.display = "block";
      sel.onchange = () => { $("cfgModel").value = sel.value; };
      if (!$("cfgModel").value) $("cfgModel").value = ids[0];
      s.innerHTML = '<span class="ok">✔ ' + ids.length + " نموذج متاح — اختار من القائمة.</span>";
    }catch(e){
      s.innerHTML = '<span class="bad">✖ ' + esc(e.message || String(e)) + "</span>";
    }
  };
});
