/* เติมข้อมูล 1 เที่ยวลงไฟล์ STORAGE TEMPLATE.XLSX
   - ใช้ ExcelJS เขียนค่า (คงรูปแบบ/สี/สูตรของ template)
   - แล้วใช้ JSZip นำ Conditional Formatting เดิมของ template กลับมา (ExcelJS ไม่รองรับ "duplicateValues" จึงตัดทิ้ง)
   ใช้ในแอพ (แท็บ "คำนวณจากไฟล์เรือ") — set = {rows:[...]} จาก voyageSets(); แต่ละแถวมี _r (LOAD), _b (BKG), _g (GATE)
   templateBuf = ArrayBuffer/Uint8Array ของ STORAGE TEMPLATE.XLSX */
(function (root) {
  const DAYMS = 864e5;
  const asDate = d => (d == null ? null : new Date(d * DAYMS));          // day number (UTC) -> Date
  const val = v => (v === undefined || v === "" ? null : v);
  const FMT = { date: "dd/mm/yyyy", money: "#,##0", int: "0" };

  // ExcelJS shares one style object between cells: copy it before changing the number format
  function put(cell, v, fmt) {
    cell.value = val(v);
    if (fmt) cell.style = Object.assign({}, cell.style, { numFmt: fmt });
  }
  function putRows(ws, startRow, rows, width) {
    rows.forEach((r, i) => {
      const xr = ws.getRow(startRow + i);
      for (let c = 0; c < width; c++) { const v = val(r ? r[c] : null); if (v !== null) xr.getCell(c + 1).value = v; }
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
  const CF_RE = /<conditionalFormatting\b[\s\S]*?<\/conditionalFormatting>/g;
  const CF_BEFORE = /<(dataValidations|hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|drawing|legacyDrawing|tableParts|extLst)\b|<\/worksheet>/;

  async function restoreConditionalFormatting(JSZip, templateBuf, outBuf) {
    const zt = await JSZip.loadAsync(templateBuf), zo = await JSZip.loadAsync(outBuf);
    const pt = await sheetParts(zt), po = await sheetParts(zo);
    for (const name of Object.keys(pt)) {
      if (!po[name]) continue;
      const tx = await zt.file(pt[name]).async("string");
      let ox = await zo.file(po[name]).async("string");
      const blocks = tx.match(CF_RE) || [];
      ox = ox.replace(CF_RE, "");
      if (blocks.length) { const m = ox.match(CF_BEFORE); ox = ox.slice(0, m.index) + blocks.join("") + ox.slice(m.index); }
      zo.file(po[name], ox);
    }
    // the restored rules point at the template's dxf list
    const ts = await zt.file("xl/styles.xml").async("string");
    let os = await zo.file("xl/styles.xml").async("string");
    const tdxfs = (ts.match(/<dxfs\b[\s\S]*?<\/dxfs>|<dxfs\b[^>]*\/>/) || [""])[0];
    if (/<dxfs\b[\s\S]*?<\/dxfs>|<dxfs\b[^>]*\/>/.test(os)) os = os.replace(/<dxfs\b[\s\S]*?<\/dxfs>|<dxfs\b[^>]*\/>/, tdxfs);
    else os = os.replace(/<tableStyles\b|<\/styleSheet>/, s => tdxfs + s);
    zo.file("xl/styles.xml", os);
    return zo.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }

  async function fillStorageTemplate(ExcelJS, JSZip, templateBuf, set) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(templateBuf);
    const rows = [...set.rows].sort((a, b) => String(a.cntr).localeCompare(String(b.cntr)));
    const uniq = list => [...new Set(list.filter(Boolean))];

    // 1-3: raw data (same as pasting the system exports into the template)
    putRows(wb.getWorksheet("1.Loading"), 2, uniq(rows.map(r => r._r)), 29);
    putRows(wb.getWorksheet("2.Bookinglist"), 2, uniq(rows.map(r => (r._b && r._b.length ? r._b : null))), 103);
    putRows(wb.getWorksheet("3.GateMove"), 2, uniq(rows.map(r => r._g)), 37);

    const remark = r => (r.ftOverride != null ? `F/T แก้ในแอพ (เดิม ${r.ftDefault})` : (r.ftNote || null));

    // CS: 41 columns A..AO from row 3
    const wsC = wb.getWorksheet("CS");
    rows.forEach((r, i) => {
      const b = r._b || [], ld = r._r || [];
      const delay = (r.orig != null && r.etd != null) ? r.etd - r.orig : null;
      const early = (r.allow != null && r.etd != null) ? r.etd - r.allow + 1 : null;
      const cells = [
        [r.cntr], [r.group], [r.loc], [r.bkno], [b[26] ?? r.shipper], [b[27]], [b[4]], [b[14]], [b[18]], [b[25]], [r.tpsz], [b[84]],
        [r.dg ? "DG" : null], [r.rf ? "RF" : null], [ld[2]], [ld[3]], [r.svc], [r.size],
        [asDate(r.orig), FMT.date], [asDate(r.etd), FMT.date], [delay, FMT.int], [asDate(r.allow), FMT.date], [asDate(r.din), FMT.date],
        [r.stay, FMT.int], [r.ft, FMT.int], [r.days, FMT.int], [r.charge, FMT.money], [r.spft || null, FMT.int], [r.tml || null],
        [r.custFt, FMT.int], [early, FMT.int], [r.over, FMT.int], [r.custAmt, FMT.money], [r.collect || null, FMT.money],
        [r.charge - (r.collect || 0), FMT.money], [remark(r)], [b[39]], [b[100]], [b[62]], [b[51]], [b[102]]];
      const xr = wsC.getRow(3 + i);
      cells.forEach(([v, f], c) => put(xr.getCell(c + 1), v, f));
    });

    // BKK / LCH: values by header meaning (the template's VLOOKUPs read one column off)
    const area = rows[0] ? rows[0].area : "BKK";
    const wsS = wb.getWorksheet(area === "LCH" ? "LCH" : "BKK");
    rows.forEach((r, i) => {
      const b = r._b || [], ld = r._r || [];
      const cells = [
        [r.cntr], [r.loc], [r.bkno], [b[26] ?? r.shipper], [b[14] ?? ld[7]], [b[18] ?? ld[8]], [r.tpsz], [ld[2]], [ld[3]],
        [asDate(r.orig), FMT.date], [asDate(r.etd), FMT.date], [(r.orig != null && r.etd != null) ? r.etd - r.orig : null, FMT.int],
        [asDate(r.din), FMT.date], [asDate(r.etd), FMT.date], [r.stay, FMT.int], [r.ft, FMT.int], [r.days, FMT.int],
        [r.charge, FMT.money], [r.custFt, FMT.int], [r.custAmt, FMT.money], [r.charge - (r.collect || 0), FMT.money],
        [remark(r)], [b[39]]];
      const xr = wsS.getRow(4 + i);
      cells.forEach(([v, f], c) => put(xr.getCell(c + 1), v, f));
    });

    const out = await wb.xlsx.writeBuffer();
    return restoreConditionalFormatting(JSZip, templateBuf, out);
  }

  root.fillStorageTemplate = fillStorageTemplate;
  if (typeof module !== "undefined") module.exports = { fillStorageTemplate };
})(typeof window !== "undefined" ? window : globalThis);
