const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const TEMPLATE_PATH = path.join(__dirname, '..', 'template.xlsx');

async function fillPlan(data, wb) {
  const ws = wb.getWorksheet(1);
  const ui       = data.userInfo   || {};
  const goals    = data.goals      || {};
  const services = data.weeklyServices || {};
  const created  = data.createdAt  || '';

  const w = (cell, value, size = 8) => {
    const c = ws.getCell(cell);
    c.value = value;
    c.font = { ...(c.font || {}), size, name: 'メイリオ' };
  };
  const wa = (cell, value, size = 8) => {
    const c = ws.getCell(cell);
    c.value = value;
    c.font = { size, name: 'メイリオ' };
    c.alignment = { wrapText: true, vertical: 'top' };
  };

  // 作成日
  let dt;
  try { dt = created ? new Date(created) : new Date(); } catch { dt = new Date(); }
  w('H1', dt.getFullYear(), 9); w('K1', dt.getMonth()+1, 9); w('M1', dt.getDate(), 9);
  w('H2', ui.svcresp || '', 9);

  // 利用者情報
  w('E3', ui.name || '', 10);
  w('N3', ui.gender === '男' ? '■男　　□ 女' : '□男　　■ 女', 9);
  const birth = ui.birth || '';
  w('V3', birth.includes('昭和') ? '□明治　□大正　■昭和' : '□明治　□大正　□昭和', 8);
  const bm = birth.match(/(\d+)年\s*(\d+)月\s*(\d+)日/);
  if (bm) { w('AB3', +bm[1]); w('AD3', +bm[2]); w('AF3', +bm[3]); }
  w('AG3', ui.addr || '', 8);

  // 緊急連絡先
  w('E4',  ui.emg1_name || '', 9);
  w('L4',  '続柄：' + (ui.emg1_rel || ''), 8);
  w('S4',  ui.emg1_tel  || '', 8);
  w('AD4', ui.emg2_name || '', 9);
  w('AK4', '続柄：' + (ui.emg2_rel || ''), 8);
  w('AR4', ui.emg2_tel  || '', 8);

  // 事業所
  w('E5',  ui.office      || '', 9);
  w('O5',  ui.office_tel  || '', 8);
  w('V5',  ui.office_addr || '', 8);
  w('AM5', ui.manager     || '', 9);
  w('E6',  ui.svcresp     || '', 9);
  w('T6',  ui.helper      || '', 9);
  w('L7',  (ui.cm_name||'') + '　' + (ui.cm_office||''), 8);
  w('Z7',  ui.cm_tel    || '', 8);
  w('AH7', ui.doctor    || '', 8);
  w('AR7', ui.doctor_tel|| '', 8);

  // 計画期間
  const period = ui.period || '';
  const pm = [...period.matchAll(/(\d{4})年(\d{1,2})月(\d{1,2})日/g)];
  if (pm.length >= 2) {
    w('H9',+pm[0][1]); w('J9',+pm[0][2]); w('L9',+pm[0][3]);
    w('O9',+pm[1][1]); w('Q9',+pm[1][2]); w('S9',+pm[1][3]);
  }

  // 週間サービス
  const dayStartCol = { 月:3, 火:8, 水:13, 木:18, 金:23, 土:28, 日:33 };
  const timeRows = {6:12,7:14,8:16,9:18,10:20,11:22,12:24,13:26,14:28,15:30,16:32,17:34,18:36,19:38,20:40,21:42,22:44,23:46};
  let bodyRow = timeRows[9];
  let lifeRow = timeRows[14];

  const colLetter = n => {
    let s = '';
    while (n > 0) { s = String.fromCharCode(64 + (n % 26 || 26)) + s; n = Math.floor((n-1)/26); }
    return s;
  };

  for (const [svcName, svcInfo] of Object.entries(services)) {
    const svcType = svcInfo.type || '';
    const svcDays = svcInfo.days || [];
    const row  = svcType === '身体介護' ? bodyRow : lifeRow;
    const icon = svcType === '身体介護' ? '▶' : '◆';
    const label = `${icon}${svcName}`;

    for (const day of svcDays) {
      const sc = dayStartCol[day];
      if (!sc) continue;
      const addr = `${colLetter(sc)}${row}`;
      const c = ws.getCell(addr);
      c.value = c.value ? `${c.value}\n${label}` : label;
      c.font = { size: 7, name: 'メイリオ' };
      c.alignment = { wrapText: true, vertical: 'top' };
    }

    if (svcType === '身体介護') bodyRow = Math.min(bodyRow + 2, lifeRow - 2);
    else lifeRow = Math.min(lifeRow + 2, 46);
  }

  // 援助目標・留意事項
  const notesList = Array.isArray(goals.notes) ? goals.notes : [String(goals.notes||'')];
  const notesStr = notesList.map(n => `・${n}`).join('\n');
  const goalText = `【長期目標】\n${goals.long||''}\n\n【短期目標①】\n${goals.short1||''}\n\n【短期目標②】\n${goals.short2||''}\n\n【留意事項】\n${notesStr}`;
  wa('AM10', goalText, 8);
  ws.getRow(10).height = 200;
  for (let r = 11; r < 48; r++) {
    const row = ws.getRow(r);
    if (!row.height || row.height < 25) row.height = 25;
  }
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method Not Allowed' });

  try {
    const data = req.body;

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(TEMPLATE_PATH);
    await fillPlan(data, wb);

    const buf = await wb.xlsx.writeBuffer();

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', "attachment; filename*=UTF-8''%E8%A8%AA%E5%95%8F%E4%BB%8B%E8%AD%B7%E8%A8%88%E7%94%BB%E6%9B%B8.xlsx");
    res.status(200).send(Buffer.from(buf));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
};
