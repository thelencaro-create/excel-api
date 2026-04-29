import ExcelJS from 'exceljs';

export const config = { maxDuration: 30 };

export default async function handler(req, res) {
  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  res.setHeader('Access-Control-Allow-Origin', '*');

  try {
    let body = req.body;
    if (typeof body === 'string') {
      try { body = JSON.parse(body); } catch(e) {}
    }
    if (!body || (Array.isArray(body) && body.length === 0)) {
      return res.status(400).json({ error: 'Leerer Body' });
    }

    // Metadaten für Dateiname
    const meta       = body._meta ?? {};
    const zielgruppe = (meta.zielgruppe || 'suche').replace(/[^a-zA-Z0-9äöüÄÖÜß\-]/g, '_');
    const stadt      = (meta.stadt      || '').replace(/[^a-zA-Z0-9äöüÄÖÜß\-]/g, '_');
    const umkreis    = meta.umkreis     || '';
    const rows       = Array.isArray(body.data) ? body.data : (Array.isArray(body) ? body : [body]);

    const headers = [
      'Unternehmen', 'Straße', 'Hausnummer', 'PLZ', 'Stadt',
      'Telefonnummer', 'Website', 'E-Mail', 'Fax'
    ];

    const wb = new ExcelJS.Workbook();
    wb.creator = 'Research Snoop';
    wb.created = new Date();

    const ws = wb.addWorksheet('Ergebnisse', {
      views: [{ state: 'frozen', ySplit: 1 }]
    });

    const colWidths = [36, 24, 12, 8, 14, 18, 38, 36, 22];
    ws.columns = headers.map((h, i) => ({
      header: h,
      key: h,
      width: colWidths[i]
    }));

    // Header-Styling
    const headerRow = ws.getRow(1);
    headerRow.height = 28;
    headerRow.eachCell(cell => {
      cell.font      = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border    = {
        top:    { style: 'thin', color: { argb: 'FFAAAAAA' } },
        bottom: { style: 'thin', color: { argb: 'FFAAAAAA' } },
        left:   { style: 'thin', color: { argb: 'FFAAAAAA' } },
        right:  { style: 'thin', color: { argb: 'FFAAAAAA' } }
      };
    });

    // Datenzeilen
    rows.forEach((item, idx) => {
      const rowData = headers.map(h => item[h] ?? '');
      const row     = ws.addRow(rowData);
      row.height    = 18;

      const fillColor = idx % 2 === 0 ? 'FFD6E4F0' : 'FFFFFFFF';
      row.eachCell({ includeEmpty: true }, cell => {
        cell.font      = { name: 'Arial', size: 10 };
        cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
        cell.alignment = { vertical: 'middle' };
        cell.border    = {
          top:    { style: 'thin', color: { argb: 'FFAAAAAA' } },
          bottom: { style: 'thin', color: { argb: 'FFAAAAAA' } },
          left:   { style: 'thin', color: { argb: 'FFAAAAAA' } },
          right:  { style: 'thin', color: { argb: 'FFAAAAAA' } }
        };
      });
    });

    // Autofilter
    ws.autoFilter = {
      from: { row: 1, column: 1 },
      to:   { row: 1, column: headers.length }
    };

    // Dateiname: Datum_Zielgruppe_Stadt_Umkreis.xlsx
    const date      = new Date().toISOString().slice(0, 10);
    const stadtStr  = stadt   ? `_${stadt}`    : '';
    const umkreisStr = umkreis ? `_${umkreis}km` : '';
    const fileName  = `${date}_${zielgruppe}${stadtStr}${umkreisStr}.xlsx`;

    const buffer = await wb.xlsx.writeBuffer();

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.send(Buffer.from(buffer));

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
}
