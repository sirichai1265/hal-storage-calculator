/* ไฟล์ STORAGE TEMPLATE ของ 1 เที่ยว แบบเดียวกับที่ทีมทำใน Excel (ตัวอย่าง 10-14-KUS-2616N-UTCT.xlsx)
   - 1.Loading / 2.Bookinglist / 3.GateMove = วางทั้งไฟล์ระบบของเที่ยวนั้น (รวมแถวสรุปท้ายไฟล์)
   - CS = เลขตู้ในคอลัมน์ A + ลากสูตรจากแถว 3 ของ template ลงมาทุกตู้ พร้อมค่าแคช (เปิดใน Protected View ก็เห็นตัวเลข)
   - BKK / LCH / Charge ไม่แตะ · เปิดไฟล์ที่ชีท CS
   ใช้ ExcelJS เขียน แล้วใช้ JSZip นำ Conditional Formatting เดิมของ template กลับมา (ExcelJS ไม่รองรับ duplicateValues)
   set = {rows:[...]} จาก voyageSets() (แต่ละแถวมี _r LOAD, _b BKG, _g GATE) · sources = SOURCES ของแอพ */
(function (root) {
  const XL_EPOCH = 25569;                                    // day number (days since 1970-01-01) -> Excel serial
  const serial = d => (d == null ? 0 : d + XL_EPOCH);
  const z = v => (v === undefined || v === null || v === "" ? 0 : v);   // VLOOKUP ของช่องว่าง = 0
  const blank = v => (v === undefined || v === null || v === "" ? null : v);
  const clone = o => JSON.parse(JSON.stringify(o || {}));

  // แปลงสูตรแถว 3 ไปแถว r (อ้างอิงแถวแบบสัมพัทธ์เท่านั้น เช่น C3 → C10 แต่ $J$1 คงเดิม)
  // ข้ามข้อความในเครื่องหมายคำพูด เช่น B3="G3" ต้องได้ B10="G3" (G3 คือกลุ่ม terminal ไม่ใช่ช่องเซลล์)
  const translate = (f, r) => f.split(/("(?:[^"]|"")*")/).map((part, i) => (i % 2 ? part :
    part.replace(/(^|[^A-Za-z0-9_$!'.])(\$?[A-Z]{1,3})3(?![0-9])/g, (m, pre, col) => pre + col + r))).join("");
  const colName = n => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

  function cached(r) {
    const b = r._b || [], ld = r._r || [], g = r._g || [];
    return {
      B: r.group, C: r.loc, D: r.bkno, E: z(b[26]), F: z(b[27]), G: z(b[4]), H: z(b[14]), I: z(b[18]), J: z(b[25]),
      K: r.tpsz, L: z(b[84]), M: r.dg ? "DG" : "", N: r.rf ? "RF" : "", O: z(ld[2]), P: z(ld[3]), Q: z(b[11]), R: r.size,
      T: z(ld[1]), U: serial(r.etd) - serial(r.orig), W: z(g[3]), X: r.stay, Y: r.ft, Z: r.days, AA: r.charge,
      AD: r.custFt, AE: serial(r.etd) - serial(r.allow) + 1, AF: r.over, AG: r.custAmt, AI: r.charge - (r.collect || 0),
      AK: z(b[39]), AL: z(b[100]), AM: z(b[62]), AN: z(b[51]), AO: z(b[102]),
    };
  }

  // ข้อมูลดิบของเที่ยวนี้จาก source: ไฟล์ที่มีแต่ข้อมูลเที่ยวนี้ → วางทั้งไฟล์ (รวมแถวท้ายไฟล์) · ไฟล์หลายเที่ยวปนกัน → เฉพาะแถวของเที่ยวนี้
  const KEYRE = { loading: /^[A-Z]{4}\d{7}$/, gate: /^[A-Z]{4}\d{7}$/, booking: /^HASLS\d/ };
  function rawFor(s, kind, keys, vkey) {
    const rows = s[kind] || [];
    const data = rows.filter(r => KEYRE[kind].test(String(r[0] == null ? "" : r[0]).trim()));
    // GATE ต้องตรงทั้งเลขตู้และ booking (ตู้ที่เคยเข้า booking เก่า สูตร VLOOKUP จะหยิบแถวเก่า)
    // ตู้ที่ไม่มีแถวไหน booking ตรงเลย (เช่น booking เปลี่ยน) → เก็บแถวของตู้นั้นไว้ (ตรงกับ storage_calc.py)
    const cn = kind === "gate" ? new Set([...keys].map(k => k.split("|")[0])) : null;
    const hasBk = kind === "gate" ? new Set(data.filter(r => r[12] && keys.has(`${r[0]}|${r[12]}`)).map(r => r[0])) : null;
    const mine = kind === "booking" && vkey ? (r => r[12] === vkey[0] && r[13] === vkey[1] && r[14] === vkey[2])
      : kind === "gate" ? (r => keys.has(`${r[0]}|${r[12] == null ? "" : r[12]}`) || (cn.has(String(r[0])) && !hasBk.has(r[0]))) : (r => keys.has(r[0]));
    if (!data.some(mine)) return [];
    if (data.every(mine)) return (s.raw && s.raw[kind]) || rows;
    return data.filter(r => mine(r) && (kind !== "loading" || r[10] === "F"));
  }
  // ตู้ซ้ำใน GATE → ใช้แถวที่ GateDate ล่าสุด (ตรงกับ pick_gate/gate_first ใน storage_calc.py และ pickGate ใน storage_app.html)
  function gDay(v) {
    if (v == null || v === "") return "";
    if (typeof v === "number") return new Date(Math.round((Math.floor(v) - 25569) * 864e5)).toISOString().slice(0, 10);
    if (v instanceof Date) return new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate())).toISOString().slice(0, 10);
    const t = String(v).trim();
    let m = t.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/); if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
    m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
    return "";
  }
  const gNewer = (a, b) => { const x = gDay(a[3]) + "|" + String(a[4] == null ? "" : a[4]), y = gDay(b[3]) + "|" + String(b[4] == null ? "" : b[4]); return x > y; };
  function pickGate(gl, bkno) {
    let cand = gl.filter(x => x[12] === bkno); if (!cand.length) cand = gl;
    let best = null; cand.forEach(x => { if (!best || gNewer(x, best)) best = x; }); return best;
  }
  // สูตร VLOOKUP หยิบแถวแรกของตู้ → ย้ายแถวที่ใช้คำนวณขึ้นก่อน (สลับเฉพาะตำแหน่งแถวข้อมูลของตู้นั้น แถวท้ายไฟล์ไม่ขยับ)
  function gateFirst(gate, bkOf) {
    const pos = {};
    gate.forEach((r, i) => { if (KEYRE.gate.test(String(r && r[0] == null ? "" : r[0]).trim())) (pos[r[0]] ||= []).push(i); });
    const out = gate.slice();
    Object.keys(pos).forEach(c => {
      const idx = pos[c]; if (idx.length < 2 || !(c in bkOf)) return;
      const gi = bkOf[c], best = pickGate(gi.dgl ? idx.map(i => gate[i]).filter(x => x[1] === gi.lw) : idx.map(i => gate[i]), gi.bk); if (!best) return;
      [best, ...idx.map(i => gate[i]).filter(r => r !== best)].forEach((r, k) => { out[idx[k]] = r; });
    });
    return out;
  }
  function pick(sources, kind, keys, fallback, vkey) {
    const out = [];
    (sources || []).forEach(s => out.push(...rawFor(s, kind, keys, vkey)));
    return out.length ? out : fallback;
  }

  function putRaw(ws, data, width) {
    const proto = Array.from({ length: width }, (_, j) => clone(ws.getRow(2).getCell(j + 1).style));
    const h = ws.getRow(2).height;
    data.forEach((row, i) => {
      const xr = ws.getRow(2 + i);
      for (let j = 0; j < width; j++) {
        const c = xr.getCell(j + 1);
        c.value = blank(row ? row[j] : null);
        if (i > 0) c.style = clone(proto[j]);
      }
      if (h) xr.height = h;
    });
  }

  // sheet name -> part path, from workbook.xml + rels
  async function sheetParts(zip) {
    const wbx = await zip.file("xl/workbook.xml").async("string");
    const rel = await zip.file("xl/_rels/workbook.xml.rels").async("string");
    const target = {};
    for (const m of rel.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = (m[0].match(/Id="([^"]+)"/) || [])[1], t = (m[0].match(/Target="([^"]+)"/) || [])[1];
      if (id && t) target[id] = t.replace(/^\/?xl\//, "").replace(/^\//, "");
    }
    const out = {};
    for (const m of wbx.matchAll(/<sheet\b[^>]*>/g)) {
      const name = (m[0].match(/name="([^"]+)"/) || [])[1], rid = (m[0].match(/r:id="([^"]+)"/) || [])[1];
      if (name && target[rid]) out[name.replace(/&amp;/g, "&")] = "xl/" + target[rid];
    }
    return out;
  }
  // ExcelJS ทิ้งกฎที่ไม่รองรับไว้เป็นแท็กว่าง <conditionalFormatting .../> (Excel ถือว่าไฟล์เสีย) → ลบทั้งแบบว่างและแบบเต็ม
  const CF_RE = /<conditionalFormatting\b[^>]*\/>|<conditionalFormatting\b[\s\S]*?<\/conditionalFormatting>/g;
  const CF_BEFORE = /<(dataValidations|hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|drawing|legacyDrawing|tableParts|extLst)\b|<\/worksheet>/;

  // ExcelJS เขียนลูกของ <sheetPr> สลับลำดับ (pageSetUpPr ก่อน outlinePr) → Excel ถือว่าชีทเสีย
  // ลำดับตามมาตรฐาน: tabColor, outlinePr, pageSetUpPr
  function fixSheetPr(xml) {
    return xml.replace(/<sheetPr\b([^>]*)>([\s\S]*?)<\/sheetPr>/, (m, attrs, body) => {
      const pick = tag => (body.match(new RegExp(`<${tag}\\b[^>]*\\/>|<${tag}\\b[^>]*>[\\s\\S]*?<\\/${tag}>`)) || [""])[0];
      return `<sheetPr${attrs}>${pick("tabColor")}${pick("outlinePr")}${pick("pageSetUpPr")}</sheetPr>`;
    });
  }

  async function restoreConditionalFormatting(JSZip, templateBuf, outBuf) {
    const zt = await JSZip.loadAsync(templateBuf), zo = await JSZip.loadAsync(outBuf);
    const pt = await sheetParts(zt), po = await sheetParts(zo);
    for (const name of Object.keys(po)) zo.file(po[name], fixSheetPr(await zo.file(po[name]).async("string")));
    for (const name of Object.keys(pt)) {
      if (!po[name]) continue;
      const tx = await zt.file(pt[name]).async("string");
      let ox = await zo.file(po[name]).async("string");
      const blocks = tx.match(CF_RE) || [];
      ox = ox.replace(CF_RE, "");
      if (blocks.length) { const m = ox.match(CF_BEFORE); ox = ox.slice(0, m.index) + blocks.join("") + ox.slice(m.index); }
      zo.file(po[name], ox);
    }
    const ts = await zt.file("xl/styles.xml").async("string");
    let os = await zo.file("xl/styles.xml").async("string");
    const DX = /<dxfs\b[\s\S]*?<\/dxfs>|<dxfs\b[^>]*\/>/;
    const tdxfs = (ts.match(DX) || [""])[0];
    os = DX.test(os) ? os.replace(DX, tdxfs) : os.replace(/<tableStyles\b|<\/styleSheet>/, s => tdxfs + s);
    zo.file("xl/styles.xml", os);
    return zo.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }

  async function fillStorageTemplate(ExcelJS, JSZip, templateBuf, set, sources) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(templateBuf);
    const rows = [...set.rows].sort((a, b) => String(a.cntr).localeCompare(String(b.cntr)));
    const cntrs = new Set(rows.map(r => r.cntr)), bks = new Set(rows.map(r => r.bkno)), pairs = new Set(rows.flatMap(r => [`${r.cntr}|${r.bkno}`, `${r.cntr}|`]));   // + แถวเติมของระบบ (มีแต่เลขตู้)
    const vkey = rows[0] && rows[0].vsl != null ? [rows[0].vsl, rows[0].voyNo, rows[0].pol] : null;
    const uniq = list => [...new Set(list.filter(x => x && x.length))];

    putRaw(wb.getWorksheet("1.Loading"), pick(sources, "loading", cntrs, uniq(rows.map(r => r._r))), 29);
    putRaw(wb.getWorksheet("2.Bookinglist"), pick(sources, "booking", bks, uniq(rows.map(r => r._b)), vkey), 103);
    const bkOf = {}; rows.forEach(r => { bkOf[r.cntr] = { bk: r.bkno, lw: r.lwharf, dgl: !!r.dg && r.pol === "THLCH" }; });
    putRaw(wb.getWorksheet("3.GateMove"), gateFirst(pick(sources, "gate", pairs, uniq(rows.map(r => r._g))), bkOf), 37);

    const ws = wb.getWorksheet("CS");
    const p = ws.getRow(3), h = p.height;
    const proto = Array.from({ length: 41 }, (_, j) => { const c = p.getCell(j + 1); return { v: c.value, style: clone(c.style) }; });
    const manual = { S: r => r.orig, V: r => r.allow, AB: r => r.spft || null, AC: r => r.tml || null, AH: r => r.collect || null };
    rows.forEach((r, i) => {
      const rr = 3 + i, xr = ws.getRow(rr), cv = cached(r);
      const ftFixed = r.ftOverride != null;                 // F/T ที่แก้เอง → ใส่เป็นค่า
      const locFixed = !!(r._r && r.loc && r._r[13] !== r.loc);   // Location ที่แก้ใน CORRECTIONS.csv → ใส่เป็นค่า
      proto.forEach((pc, j) => {
        const col = colName(j + 1), cell = xr.getCell(j + 1);
        if (i > 0) cell.style = clone(pc.style);
        const f = pc.v && typeof pc.v === "object" && pc.v.formula;
        if (j === 0) cell.value = r.cntr;
        else if (col === "C" && locFixed) cell.value = r.loc;
        else if (col === "Y" && ftFixed) cell.value = r.ft;
        else if (col === "W" && r.dinFixed) cell.value = new Date(r.etd * 864e5);   // DG LCH วันตู้เข้า = ETD (ไม่มีแถว GATE ให้สูตรหา) → ใส่เป็นค่า
        else if (f) cell.value = { formula: translate(f, rr), result: cv[col] !== undefined ? cv[col] : undefined };
        else if (manual[col]) {
          const v = manual[col](r);
          cell.value = (col === "S" || col === "V") ? (v == null ? null : new Date(v * 864e5)) : blank(v);
        } else if (col === "AJ") cell.value = (ftFixed || locFixed || r.dinFixed) ? (r.fixNote || `F/T แก้ในแอพ (เดิม ${r.ftDefault})`) : null;
        else cell.value = null;
      });
      if (h) xr.height = h;
    });
    const tot = k => rows.reduce((t, r) => t + (k(r) || 0), 0);
    [["AA", r => r.charge], ["AG", r => r.custAmt], ["AH", r => r.collect], ["AI", r => r.charge - (r.collect || 0)]].forEach(([col, k]) => {
      const c = ws.getCell(col + "1");
      if (c.value && c.value.formula) c.value = { formula: c.value.formula, result: tot(k) };
    });

    wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });
    const ci = wb.worksheets.findIndex(w => w.name === "CS");
    if (wb.views && wb.views.length) wb.views[0].activeTab = ci; else wb.views = [{ activeTab: ci }];

    const out = await wb.xlsx.writeBuffer();
    return restoreConditionalFormatting(JSZip, templateBuf, out);
  }

  root.fillStorageTemplate = fillStorageTemplate;
  if (typeof module !== "undefined") module.exports = { fillStorageTemplate, _translate: translate };
})(typeof window !== "undefined" ? window : globalThis);
