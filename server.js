const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Upload Endpoint
app.post('/api/convert', upload.single('excel'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('Please upload an Excel file.');
    }

    // 1. Excel File Read
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

    // 2. Response Header Setup for PDF Download
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="converted.pdf"');

    // 3. Create PDF
    const doc = new PDFDocument({ margin: 30 });
    doc.pipe(res);

    doc.fontSize(16).text('Excel Data Report', { align: 'center' });
    doc.moveDown();

    // 4. Write Excel Rows to PDF
    sheetData.forEach((row) => {
      const rowText = row.join('  |  ');
      doc.fontSize(10).text(rowText);
      doc.moveDown(0.5);
    });

    doc.end();
  } catch (error) {
    res.status(500).send('Error processing file: ' + error.message);
  }
});

// Root route for testing
app.get('/', (req, res) => {
  res.send(`
    <h2>Upload Excel File to Convert to PDF</h2>
    <form action="/api/convert" method="post" enctype="multipart/form-data">
      <input type="file" name="excel" accept=".xlsx, .xls" required />
      <button type="submit">Convert to PDF</button>
    </form>
  `);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));

module.exports = app;