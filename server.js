const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');
const docx = require('docx');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Helper function to build PDF logic safely
function generateQuestionPDF(rawData, res, docTitle = 'ঢাবি ভর্তি পরীক্ষা : প্রশ্নপত্র ও সমাধান') {
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', 'attachment; filename="question-paper.pdf"');

  const doc = new PDFDocument({ margin: 40, size: 'A4' });
  doc.pipe(res);

  // 1. Register Font
  const fontPath = path.join(__dirname, 'fonts', 'kalpurush.ttf');
  doc.registerFont('BanglaFont', fontPath);

  // 2. Document Header
  doc.font('BanglaFont').fontSize(18).fillColor('#0f172a').text(docTitle, { align: 'center' });
  doc.moveDown(0.3);
  doc.fontSize(10).fillColor('#64748b').text('উৎসব / বিষয়: C-Unit Admission Test', { align: 'center' });
  doc.moveDown(0.5);

  // Header Divider Line
  doc.strokeColor('#cbd5e1').lineWidth(1).moveTo(40, doc.y).lineTo(555, doc.y).stroke();
  doc.moveDown(1.2);

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
      // Question Text
      doc.font('BanglaFont').fontSize(11).fillColor('#0f172a').text(`প্রশ্ন ${index + 1}. ${qText}`, { lineGap: 4 });
      doc.moveDown(0.4);

      // Options Array
      const options = [
        { key: 'A', label: '(A)', val: optA },
        { key: 'B', label: '(B)', val: optB },
        { key: 'C', label: '(C)', val: optC },
        { key: 'D', label: '(D)', val: optD }
      ];

      // Highlight Correct Answer Option
      const formattedOptions = options.map(opt => {
        const isCorrect = (opt.key === correctOpt);
        if (isCorrect) {
          return `▶ ${opt.label} ${opt.val}`;
        }
        return `   ${opt.label} ${opt.val}`;
      }).join('     ');

      doc.font('BanglaFont').fontSize(10).fillColor('#334155').text(formattedOptions, { indent: 15, lineGap: 4 });
      doc.moveDown(0.5);

      // Render Explanation
      if (explanation) {
        doc.font('BanglaFont').fontSize(9.5).fillColor('#047857').text(`ব্যাখ্যা: ${explanation}`, { indent: 15, lineGap: 4 });
      }

      doc.moveDown(1.4);
    }
  });

  doc.end();
}

// 1. Excel to PDF Endpoint
app.post('/api/convert-excel', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('Please upload an Excel file.');
    }

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    generateQuestionPDF(rawData, res, 'ঢাবি ভর্তি পরীক্ষা : প্রশ্নপত্র ও সমাধান');
  } catch (error) {
    res.status(500).send('Error generating question paper PDF: ' + error.message);
  }
});

// 2. Excel to Word (.docx) Endpoint
app.post('/api/convert-excel-to-word', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('Please upload an Excel file.');
    }

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    const docParagraphs = [
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: "ঢাবি ভর্তি পরীক্ষা : প্রশ্নপত্র ও সমাধান", bold: true, size: 32 })
        ],
        alignment: docx.AlignmentType.CENTER
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: "উৎসব / বিষয়: C-Unit Admission Test", color: "64748B", size: 20 })
        ],
        alignment: docx.AlignmentType.CENTER
      }),
      new docx.Paragraph({ text: "" })
    ];

    rawData.forEach((q, index) => {
      const qText = q.question_text || q.Question || q.question || '';
      const optA = q.option_a || q['Option A'] || q.a || '';
      const optB = q.option_b || q['Option B'] || q.b || '';
      const optC = q.option_c || q['Option C'] || q.c || '';
      const optD = q.option_d || q['Option D'] || q.d || '';
      const correctOpt = String(q.correct_option || q.Answer || q.correct || '').trim().toUpperCase();
      const explanation = (q['Explanation '] || q.Explanation || q.explanation || '').trim();

      if (qText) {
        // Question Line
        docParagraphs.push(
          new docx.Paragraph({
            children: [
              new docx.TextRun({ text: `প্রশ্ন ${index + 1}. ${qText}`, bold: true, size: 22 })
            ]
          })
        );

        // Options Line
        const optionsStr = `(A) ${optA}      (B) ${optB}      (C) ${optC}      (D) ${optD}`;
        docParagraphs.push(
          new docx.Paragraph({
            children: [
              new docx.TextRun({ text: optionsStr, size: 20 })
            ],
            indent: { left: 360 }
          })
        );

        // Explanation Line
        if (explanation) {
          docParagraphs.push(
            new docx.Paragraph({
              children: [
                new docx.TextRun({ text: `ব্যাখ্যা: ${explanation}`, color: "047857", size: 18 })
              ],
              indent: { left: 360 }
            })
          );
        }

        docParagraphs.push(new docx.Paragraph({ text: "" }));
      }
    });

    const docFile = new docx.Document({
      sections: [{ children: docParagraphs }]
    });

    const buffer = await docx.Packer.toBuffer(docFile);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="question-paper.docx"');
    res.send(buffer);
  } catch (error) {
    res.status(500).send('Error generating Word document: ' + error.message);
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
      <title>DocTransform Studio</title>
      <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-slate-100 text-slate-800 min-h-screen font-sans">
      <header class="bg-white border-b border-slate-200 py-4 px-8 flex justify-between items-center shadow-sm">
        <div class="flex items-center gap-3">
          <div class="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center text-white font-bold text-xl shadow-md">📄</div>
          <div>
            <h1 class="text-xl font-bold text-slate-800 leading-tight">DocTransform Studio</h1>
            <p class="text-xs text-slate-500">Excel to PDF & Word Question Generator</p>
          </div>
        </div>
      </header>

      <main class="max-w-4xl mx-auto py-10 px-4">
        <!-- Dashboard Navigation Tabs -->
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
          <button id="excelTab" onclick="switchMode('excel')" class="bg-white p-6 rounded-2xl border-2 border-indigo-600 shadow-sm hover:shadow-md transition-all text-left flex items-start gap-4 cursor-pointer">
            <div class="w-12 h-12 bg-emerald-100 text-emerald-600 rounded-xl flex items-center justify-center text-2xl font-bold shrink-0">📊</div>
            <div>
              <h3 class="font-bold text-slate-800 text-lg">EXCEL TO PDF</h3>
              <p class="text-slate-500 text-xs mt-1">MCQ Excel to PDF Question Generator</p>
            </div>
          </button>

          <button id="wordTab" onclick="switchMode('word')" class="bg-white p-6 rounded-2xl border border-slate-200 opacity-60 hover:opacity-100 transition-all text-left flex items-start gap-4 cursor-pointer">
            <div class="w-12 h-12 bg-blue-100 text-blue-600 rounded-xl flex items-center justify-center text-2xl font-bold shrink-0">📝</div>
            <div>
              <h3 class="font-bold text-slate-800 text-lg">EXCEL TO WORD</h3>
              <p class="text-slate-500 text-xs mt-1">Convert MCQ Excel into Word Document</p>
            </div>
          </button>
        </div>

        <!-- Form Card -->
        <div class="bg-white rounded-3xl border border-slate-200 p-10 shadow-sm text-center">
          <form id="convertForm" action="/api/convert-excel" method="post" enctype="multipart/form-data">
            <div class="border-2 border-dashed border-slate-300 hover:border-indigo-500 bg-slate-50/50 hover:bg-indigo-50/30 rounded-2xl p-12 transition-all cursor-pointer flex flex-col items-center justify-center mb-6" onclick="document.getElementById('fileInput').click()">
              <div class="w-16 h-16 bg-indigo-50 text-indigo-600 rounded-2xl flex items-center justify-center mb-4 text-3xl">📥</div>
              <h3 id="dragText" class="text-lg font-bold text-slate-700 mb-1">Excel প্রশ্ন ফাইল আপলোড করুন</h3>
              <p id="subText" class="text-slate-400 text-xs mb-6">ফাইল সিলেক্ট করে সুন্দর Question Paper PDF বানিয়ে নিন</p>
              <span class="bg-indigo-600 hover:bg-indigo-700 text-white font-semibold text-sm px-6 py-3 rounded-xl shadow-sm transition-all inline-block">ফাইল সিলেক্ট করুন</span>
              <input type="file" name="file" accept=".xlsx, .xls" class="hidden" required id="fileInput" onchange="handleFileChange(this)" />
            </div>

            <div id="fileInfo" class="hidden mb-6 p-4 bg-slate-100 rounded-xl flex items-center justify-between">
              <span id="selectedFileName" class="text-sm font-medium text-slate-700"></span>
              <button type="submit" id="submitBtn" class="bg-emerald-600 hover:bg-emerald-700 text-white font-semibold px-6 py-2 rounded-lg text-sm shadow-sm transition-all">
                Download PDF Now
              </button>
            </div>
          </form>
        </div>
      </main>

      <script>
        function switchMode(mode) {
          const form = document.getElementById('convertForm');
          const fileInput = document.getElementById('fileInput');
          const excelTab = document.getElementById('excelTab');
          const wordTab = document.getElementById('wordTab');
          const dragText = document.getElementById('dragText');
          const subText = document.getElementById('subText');
          const fileInfo = document.getElementById('fileInfo');
          const submitBtn = document.getElementById('submitBtn');
          
          fileInput.value = '';
          fileInfo.classList.add('hidden');

          if (mode === 'excel') {
            form.action = '/api/convert-excel';
            fileInput.accept = '.xlsx, .xls';
            excelTab.className = 'bg-white p-6 rounded-2xl border-2 border-indigo-600 shadow-sm hover:shadow-md transition-all text-left flex items-start gap-4 cursor-pointer';
            wordTab.className = 'bg-white p-6 rounded-2xl border border-slate-200 opacity-60 hover:opacity-100 transition-all text-left flex items-start gap-4 cursor-pointer';
            dragText.textContent = 'Excel প্রশ্ন ফাইল আপলোড করুন';
            subText.textContent = 'ফাইল সিলেক্ট করে সুন্দর Question Paper PDF বানিয়ে নিন';
            submitBtn.textContent = 'Download PDF Now';
          } else {
            form.action = '/api/convert-excel-to-word';
            fileInput.accept = '.xlsx, .xls';
            wordTab.className = 'bg-white p-6 rounded-2xl border-2 border-indigo-600 shadow-sm hover:shadow-md transition-all text-left flex items-start gap-4 cursor-pointer';
            excelTab.className = 'bg-white p-6 rounded-2xl border border-slate-200 opacity-60 hover:opacity-100 transition-all text-left flex items-start gap-4 cursor-pointer';
            dragText.textContent = 'Excel প্রশ্ন ফাইল আপলোড করুন';
            subText.textContent = 'ফাইল সিলেক্ট করে Word Document (.docx) বানিয়ে নিন';
            submitBtn.textContent = 'Download Word (.docx) Now';
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