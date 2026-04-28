import ExcelJS from 'exceljs';

export const config = { maxDuration: 30 };

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const rows = Array.isArray(req.body) ? req.body : [req.body];

    const headers = [
      'Unternehmen', 'Straße', 'Hausnummer', 'PLZ', 'Stadt',
      'Telefonnummer', 'Website', 'E-Mail', 'Fax'
    ];

    const wb = new ExcelJS.Workbook();
    wb.creator = 'Research Snoop';
    wb.created = new Date();

    const ws = wb.addWorksheet('Apotheken', {
      views: [{ state: 'frozen', ySplit: 1 }]
    });

    const colWidths = [36, 24, 12, 8, 14, 18, 38, 36, 22];
    ws.columns = headers.map((h, i) => ({
      header: h,
      key: h,
      width: colWidths[i]
    }));

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

    rows.forEach((item, idx) => {
      const rowData = headers.map(h => item[h] ?? '');
      const row = ws.addRow(rowData);
      row.height = 18;

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

    ws.autoFilter = {
      from: { row: 1, column: 1 },
      to:   { row: 1, column: headers.length }
    };

    const buffer = await wb.xlsx.writeBuffer();
    const date   = new Date().toISOString().slice(0, 10);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${date}_apotheken.xlsx"`);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.send(Buffer.from(buffer));

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
}
