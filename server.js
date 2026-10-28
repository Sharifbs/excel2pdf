const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Excel to PDF Endpoint (Bold Option Label + Explanation Included, No 'Answer' Text)
app.post('/api/convert-excel', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('Please upload an Excel file.');
    }

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="question-paper.pdf"');

    const doc = new PDFDocument({ margin: 40, size: 'A4' });
    doc.pipe(res);

    // 1. Register Bangla Font
    const fontPath = path.join(__dirname, 'fonts', 'kalpurush.ttf');
    doc.registerFont('BanglaFont', fontPath);

    // 2. Question Paper Header
    doc.font('BanglaFont').fontSize(18).fillColor('#0f172a').text('ঢাবি ভর্তি পরীক্ষা : প্রশ্নপত্র ও সমাধান', { align: 'center' });
    doc.moveDown(0.3);
    doc.fontSize(10).fillColor('#64748b').text('উৎসব / বিষয়: C-Unit Admission Test', { align: 'center' });
    doc.moveDown(0.5);
    
    // Header underline divider
    doc.strokeColor('#cbd5e1').lineWidth(1).moveTo(40, doc.y).lineTo(555, doc.y).stroke();
    doc.moveDown(1);

    // 3. Render Questions
    rawData.forEach((q, index) => {
      const qText = q.question_text || q.Question || q.question || '';
      const optA = q.option_a || q['Option A'] || q.a || '';
      const optB = q.option_b || q['Option B'] || q.b || '';
      const optC = q.option_c || q['Option C'] || q.c || '';
      const optD = q.option_d || q['Option D'] || q.d || '';
      
      const correctOpt = String(q.correct_option || q.Answer || q.correct || '').trim().toUpperCase();
      const explanation = (q['Explanation '] || q.Explanation || q.explanation || '').trim();

      if (qText) {
        // Render Question Text
        doc.font('BanglaFont').fontSize(11).fillColor('#0f172a').text(`প্রশ্ন ${index + 1}. ${qText}`, { lineGap: 3 });
        
        const options = [
          { key: 'A', label: '(A)', val: optA },
          { key: 'B', label: '(B)', val: optB },
          { key: 'C', label: '(C)', val: optC },
          { key: 'D', label: '(D)', val: optD }
        ];

        let startX = 55;
        const currentY = doc.y;

        // Render Options: Only bold the label like (D) for the correct option
        options.forEach((opt) => {
          const isCorrect = (opt.key === correctOpt);
          
          if (isCorrect) {
            doc.font('BanglaFont').fontSize(10.5).fillColor('#000000').text(`${opt.label} `, startX, currentY, { continued: true });
            doc.font('BanglaFont').fontSize(10).fillColor('#475569').text(`${opt.val}`, { continued: true });
          } else {
            doc.font('BanglaFont').fontSize(10).fillColor('#475569').text(`${opt.label} ${opt.val}`, startX, currentY, { continued: true });
          }
          
          doc.text('      ', { continued: true });
        });

        doc.text('');
        doc.moveDown(0.4);

        // Render Explanation Text Only (Without 'সঠিক উত্তর: X')
        if (explanation) {
          doc.font('BanglaFont').fontSize(9).fillColor('#047857').text(`ব্যাখ্যা: ${explanation}`, { indent: 15, lineGap: 3 });
        }

        doc.moveDown(0.8);
      }
    });

    doc.end();
  } catch (error) {
    res.status(500).send('Error generating question paper PDF: ' + error.message);
  }
});

// Word to PDF Endpoint
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

    const fontPath = path.join(__dirname, 'fonts', 'kalpurush.ttf');
    doc.registerFont('BanglaFont', fontPath);

    doc.font('BanglaFont').fontSize(18).fillColor('#1e293b').text('Word Document Conversion', { align: 'center' });
    doc.moveDown(1.5);
    doc.font('BanglaFont').fontSize(11).fillColor('#334155').text(fileContent);

    doc.end();
  } catch (error) {
    res.status(500).send('Error processing Word file: ' + error.message);
  }
});

// UI Dashboard
app.get('/', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>DocTransform Studio - Question Paper</title>
      <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-slate-100 text-slate-800 min-h-screen font-sans">
      <header class="bg-white border-b border-slate-200 py-4 px-8 flex justify-between items-center shadow-sm">
        <div class="flex items-center gap-3">
          <div class="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center text-white font-bold text-xl shadow-md">📝</div>
          <div>
            <h1 class="text-xl font-bold text-slate-800 leading-tight">DocTransform Studio</h1>
            <p class="text-xs text-slate-500">MCQ Excel to PDF Question Generator</p>
          </div>
        </div>
      </header>

      <main class="max-w-4xl mx-auto py-10 px-4">
        <div class="bg-white rounded-3xl border border-slate-200 p-10 shadow-sm text-center">
          <form id="convertForm" action="/api/convert-excel" method="post" enctype="multipart/form-data">
            <div class="border-2 border-dashed border-slate-300 hover:border-indigo-500 bg-slate-50/50 hover:bg-indigo-50/30 rounded-2xl p-12 transition-all cursor-pointer flex flex-col items-center justify-center mb-6" onclick="document.getElementById('fileInput').click()">
              <div class="w-16 h-16 bg-indigo-50 text-indigo-600 rounded-2xl flex items-center justify-center mb-4 text-3xl">📥</div>
              <h3 class="text-lg font-bold text-slate-700 mb-1">Excel প্রশ্ন ফাইল আপলোড করুন</h3>
              <p class="text-slate-400 text-xs mb-6">ফাইল সিলেক্ট করে বোল্ড অপশন লেবেল ও ব্যাখ্যা সহ Question Paper ডাউনলোড করুন</p>
              <span class="bg-indigo-600 hover:bg-indigo-700 text-white font-semibold text-sm px-6 py-3 rounded-xl shadow-sm transition-all inline-block">ফাইল সিলেক্ট করুন</span>
              <input type="file" name="file" accept=".xlsx, .xls" class="hidden" required id="fileInput" onchange="handleFileChange(this)" />
            </div>

            <div id="fileInfo" class="hidden mb-6 p-4 bg-slate-100 rounded-xl flex items-center justify-between">
              <span id="selectedFileName" class="text-sm font-medium text-slate-700"></span>
              <button type="submit" class="bg-emerald-600 hover:bg-emerald-700 text-white font-semibold px-6 py-2 rounded-lg text-sm shadow-sm transition-all">
                Download PDF
              </button>
            </div>
          </form>
        </div>
      </main>

      <script>
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