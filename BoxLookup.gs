var MENU_ITEMS = MENU_ITEMS || [];
MENU_ITEMS.push({ label: 'Run Box Lookup', fn: 'runBoxLookup' });

const SOURCE_SS_ID = '1aZzyH7fpMtAxY1UyxRsBBLAWRS1cWpQIN37TvxNm00I';

function runBoxLookup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dSheet = ss.getSheetByName('D-1');
  const src = SpreadsheetApp.openById(SOURCE_SS_ID);

  const psData   = src.getSheetByName('P&S').getDataRange().getValues();
  const pbdData  = src.getSheetByName('PBD').getDataRange().getValues();

  const qcNewSheet = src.getSheetByName('QC NEW');
  const qcNewData  = qcNewSheet ? qcNewSheet.getDataRange().getValues() : [];

  const coldSheet = src.getSheetByName('Cold-NEW');
  const coldData  = coldSheet ? coldSheet.getDataRange().getValues() : [];

  const hvcSheet = src.getSheetByName('HVC');
  const hvcData  = hvcSheet ? hvcSheet.getDataRange().getValues() : [];

  // Build P&S Map (Col B: ID, Col C: Floor, Col D: Station, Col A: Timestamp)
  const psMap = {};
  for (let i = 1; i < psData.length; i++) {
    const id = String(psData[i][1]).trim();
    if (!id || psMap[id]) continue;
    const floor = String(psData[i][2]).trim();
    const station = String(psData[i][3]).trim();
    psMap[id] = { ts: psData[i][0], loc: station ? floor + ' - ' + station : floor };
  }

  // Build PBD Map (Col B: ID, Col C: Location, Col A: Timestamp)
  const pbdMap = {};
  for (let i = 1; i < pbdData.length; i++) {
    const id = String(pbdData[i][1]).trim();
    if (!id || pbdMap[id]) continue;
    pbdMap[id] = { ts: pbdData[i][0], loc: String(pbdData[i][2]).trim() };
  }

  // Build QC NEW Map (same layout as PBD -> Col B: ID, Col C: Location, Col A: Timestamp)
  const qcNewMap = {};
  for (let i = 1; i < qcNewData.length; i++) {
    const id = String(qcNewData[i][1]).trim();
    if (!id || qcNewMap[id]) continue;
    qcNewMap[id] = { ts: qcNewData[i][0], loc: String(qcNewData[i][2]).trim() };
  }

  // Build Cold-NEW Map (Col F: ID, Col E: Table No, Col B: Timestamp)
  const coldMap = {};
  for (let i = 1; i < coldData.length; i++) {
    const id = String(coldData[i][5]).trim();
    if (!id || coldMap[id]) continue;
    coldMap[id] = { ts: coldData[i][1], loc: 'Cold- ' + String(coldData[i][4]).trim() };
  }

  // Build HVC Map (Col C: ID, Hardcoded Location: 'HVC', Col A: Timestamp)
  const hvcMap = {};
  for (let i = 1; i < hvcData.length; i++) {
    const id = String(hvcData[i][2]).trim();
    if (!id || hvcMap[id]) continue;
    hvcMap[id] = { ts: hvcData[i][0], loc: 'HVC' };
  }

  const lastRow = dSheet.getLastRow();
  if (lastRow < 2) return;
  const extIds = dSheet.getRange(2, 4, lastRow - 1, 1).getValues();

  const tz = ss.getSpreadsheetTimeZone();
  const fmt = v => v instanceof Date ? Utilities.formatDate(v, tz, 'M/d/yyyy HH:mm:ss') : String(v);

  // Priority: [P&S / PBD / QC NEW] (any combo found, appended) -> Cold-NEW -> HVC -> Not Found
  const out = extIds.map(([raw]) => {
    const id = String(raw).trim();
    const a = psMap[id], b = pbdMap[id], q = qcNewMap[id], c = coldMap[id], d = hvcMap[id];

    const topTier = [a, b, q].filter(Boolean);
    if (topTier.length) {
      return [
        topTier.map(x => x.loc).join(' / '),
        topTier.map(x => fmt(x.ts)).join(' / ')
      ];
    }
    if (c) return [c.loc, fmt(c.ts)];
    if (d) return [d.loc, fmt(d.ts)];

    return ['#N/A', 'Not Found'];
  });

  dSheet.getRange(1, 25, 1, 2).setValues([['Location', 'Timestamp']]);
  dSheet.getRange(2, 25, out.length, 2).setValues(out);
}
