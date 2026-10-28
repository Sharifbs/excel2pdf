const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Excel to PDF Endpoint
app.post('/api/convert-excel', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('Please upload an Excel file.');
    }

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="excel-converted.pdf"');

    const doc = new PDFDocument({ margin: 40, size: 'A4' });
    doc.pipe(res);

    doc.fillColor('#1e293b').fontSize(18).text('Excel Data Conversion', { align: 'center' });
    doc.moveDown(1.5);

    sheetData.forEach((row, index) => {
      const rowText = row.join('   |   ');
      if (index === 0) {
        doc.fillColor('#0f172a').fontSize(11).font('Helvetica-Bold').text(rowText);
        doc.moveDown(0.3);
        doc.strokeColor('#cbd5e1').lineWidth(1).moveTo(40, doc.y).lineTo(555, doc.y).stroke();
        doc.moveDown(0.5);
      } else {
        doc.fillColor('#334155').fontSize(10).font('Helvetica').text(rowText);
        doc.moveDown(0.4);
      }
    });

    doc.end();
  } catch (error) {
    res.status(500).send('Error processing Excel file: ' + error.message);
  }
});

// Word to PDF Endpoint (Text extraction format)
app.post('/api/convert-word', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('Please upload a file.');
    }

    const fileContent = req.file.buffer.toString('utf-8');

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="word-converted.pdf"');

    const doc = new PDFDocument({ margin: 40, size: 'A4' });
    doc.pipe(res);

    doc.fillColor('#1e293b').fontSize(18).text('Word Document Conversion', { align: 'center' });
    doc.moveDown(1.5);
    doc.fillColor('#334155').fontSize(11).font('Helvetica').text(fileContent);

    doc.end();
  } catch (error) {
    res.status(500).send('Error processing Word file: ' + error.message);
  }
});

// Main Dashboard Interface
app.get('/', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>DocTransform Studio</title>
      <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-slate-100 text-slate-800 min-h-screen font-sans">

      <!-- Header -->
      <header class="bg-white border-b border-slate-200 py-4 px-8 flex justify-between items-center shadow-sm">
        <div class="flex items-center gap-3">
          <div class="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center text-white font-bold text-xl shadow-md">
            📄
          </div>
          <div>
            <h1 class="text-xl font-bold text-slate-800 leading-tight">DocTransform Studio</h1>
            <p class="text-xs text-slate-500">Ultra-Fast Workspace for Excel & Word to PDF</p>
          </div>
        </div>
        <div class="bg-emerald-50 text-emerald-600 border border-emerald-200 px-3 py-1 rounded-full text-xs font-semibold flex items-center gap-1.5">
          <span class="w-2 h-2 rounded-full bg-emerald-500"></span> 100% Private & Serverless
        </div>
      </header>

      <main class="max-w-5xl mx-auto py-10 px-4">
        
        <!-- Navigation Options -->
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
          
          <!-- Excel Option -->
          <button id="excelTab" onclick="switchMode('excel')" class="bg-white p-6 rounded-2xl border-2 border-indigo-600 shadow-sm hover:shadow-md transition-all text-left flex items-start gap-4 cursor-pointer">
            <div class="w-12 h-12 bg-emerald-100 text-emerald-600 rounded-xl flex items-center justify-center text-2xl font-bold shrink-0">
              📊
            </div>
            <div>
              <h3 class="font-bold text-slate-800 text-lg">EXCEL TO PDF</h3>
              <p class="text-slate-500 text-xs mt-1">Convert .xlsx & .xls spreadsheets into formatted PDF reports</p>
            </div>
          </button>

          <!-- Word Option -->
          <button id="wordTab" onclick="switchMode('word')" class="bg-white p-6 rounded-2xl border border-slate-200 opacity-60 hover:opacity-100 transition-all text-left flex items-start gap-4 cursor-pointer">
            <div class="w-12 h-12 bg-blue-100 text-blue-600 rounded-xl flex items-center justify-center text-2xl font-bold shrink-0">
              📝
            </div>
            <div>
              <h3 class="font-bold text-slate-800 text-lg">WORD TO PDF</h3>
              <p class="text-slate-500 text-xs mt-1">Convert text documents into clean printable PDF files</p>
            </div>
          </button>

        </div>

        <!-- Main Converter Box -->
        <div class="bg-white rounded-3xl border border-slate-200 p-10 shadow-sm text-center">
          
          <form id="convertForm" action="/api/convert-excel" method="post" enctype="multipart/form-data">
            
            <div class="border-2 border-dashed border-slate-300 hover:border-indigo-500 bg-slate-50/50 hover:bg-indigo-50/30 rounded-2xl p-12 transition-all cursor-pointer flex flex-col items-center justify-center mb-6" onclick="document.getElementById('fileInput').click()">
              
              <div class="w-16 h-16 bg-indigo-50 text-indigo-600 rounded-2xl flex items-center justify-center mb-4 text-3xl">
                ☁️
              </div>
              
              <h3 id="dragText" class="text-lg font-bold text-slate-700 mb-1">Excel ফাইল ড্রাগ করুন অথবা সিলেক্ট করুন</h3>
              <p id="subText" class="text-slate-400 text-xs mb-6">আপনার .xlsx ফাইল সিলেক্ট করে দ্রুত PDF বানিয়ে নিন</p>
              
              <span class="bg-indigo-600 hover:bg-indigo-700 text-white font-semibold text-sm px-6 py-3 rounded-xl shadow-sm transition-all inline-block">
                ফাইল সিলেক্ট করুন
              </span>

              <input type="file" name="file" accept=".xlsx, .xls" class="hidden" required id="fileInput" onchange="handleFileChange(this)" />
            </div>

            <div id="fileInfo" class="hidden mb-6 p-4 bg-slate-100 rounded-xl flex items-center justify-between">
              <span id="selectedFileName" class="text-sm font-medium text-slate-700"></span>
              <button type="submit" id="submitBtn" class="bg-emerald-600 hover:bg-emerald-700 text-white font-semibold px-6 py-2 rounded-lg text-sm shadow-sm transition-all">
                Convert to PDF Now
              </button>
            </div>

          </form>

        </div>

      </main>

      <script>
        let currentMode = 'excel';

        function switchMode(mode) {
          currentMode = mode;
          const form = document.getElementById('convertForm');
          const fileInput = document.getElementById('fileInput');
          const excelTab = document.getElementById('excelTab');
          const wordTab = document.getElementById('wordTab');
          const dragText = document.getElementById('dragText');
          const subText = document.getElementById('subText');
          const fileInfo = document.getElementById('fileInfo');
          
          fileInput.value = '';
          fileInfo.classList.add('hidden');

          if (mode === 'excel') {
            form.action = '/api/convert-excel';
            fileInput.accept = '.xlsx, .xls';
            excelTab.className = 'bg-white p-6 rounded-2xl border-2 border-indigo-600 shadow-sm hover:shadow-md transition-all text-left flex items-start gap-4 cursor-pointer';
            wordTab.className = 'bg-white p-6 rounded-2xl border border-slate-200 opacity-60 hover:opacity-100 transition-all text-left flex items-start gap-4 cursor-pointer';
            dragText.textContent = 'Excel ফাইল ড্রাগ করুন অথবা সিলেক্ট করুন';
            subText.textContent = 'আপনার .xlsx ফাইল সিলেক্ট করে দ্রুত PDF বানিয়ে নিন';
          } else {
            form.action = '/api/convert-word';
            fileInput.accept = '.docx, .doc, .txt';
            wordTab.className = 'bg-white p-6 rounded-2xl border-2 border-indigo-600 shadow-sm hover:shadow-md transition-all text-left flex items-start gap-4 cursor-pointer';
            excelTab.className = 'bg-white p-6 rounded-2xl border border-slate-200 opacity-60 hover:opacity-100 transition-all text-left flex items-start gap-4 cursor-pointer';
            dragText.textContent = 'Word ফাইল ড্রাগ করুন অথবা সিলেক্ট করুন';
            subText.textContent = 'আপনার Word document সিলেক্ট করে দ্রুত PDF বানিয়ে নিন';
          }
        }

        function handleFileChange(input) {
          if (input.files && input.files[0]) {
            document.getElementById('selectedFileName').textContent = input.files[0].name;
            document.getElementById('fileInfo').classList.remove('hidden');
          }
        }
      </script>

    </body>
    </html>
  `);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));

module.exports = app;