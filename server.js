const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');
const docx = require('docx');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// 1. Excel Template Download Endpoint
app.get('/api/download-template', (req, res) => {
  try {
    const templateData = [
      {
        question_text: "মহিমা শব্দের সঠিক প্রকৃতি প্রত্যয় কোনটি?",
        option_a: "মহি+মা",
        option_b: "মহা+ইমা",
        option_c: "মহা+ইমন",
        option_d: "মহৎ+ইমন",
        correct_option: "D",
        Explanation: "মহিমা = মহৎ + ইমন (সংস্কৃত তদ্ধিত প্রত্যয় যোগে গঠিত)।"
      },
      {
        question_text: "চন্দ্র শব্দের সমার্থক শব্দ নয় কোনটি?",
        option_a: "ইন্দু",
        option_b: "কলাভৃৎ",
        option_c: "সোম",
        option_d: "বিভাকর",
        correct_option: "D",
        Explanation: "বিভাকর অর্থ সূর্য। ইন্দু, কলাভৃৎ, সোম ও কলাধর শব্দগুলো চন্দ্র এর সমার্থক শব্দ।"
      }
    ];

    const worksheet = XLSX.utils.json_to_sheet(templateData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Questions");

    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="mcq_question_template.xlsx"');
    res.send(buffer);
  } catch (error) {
    res.status(500).send('Error downloading template: ' + error.message);
  }
});

// 2. Excel to PDF Endpoint
app.post('/api/convert-excel', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('Please upload an Excel file.');
    }

    const testName = req.body.testName || 'ঢাবি ভর্তি পরীক্ষা : প্রশ্নপত্র ও সমাধান';
    const subject = req.body.subject || 'C-Unit Admission Test';
    const fullMarks = req.body.fullMarks || '১০০';
    const duration = req.body.duration || '১ ঘণ্টা';

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    res.setHeader('Content-Type', 'application/pdf');

    const doc = new PDFDocument({ margin: 40, size: 'A4' });
    doc.pipe(res);

    // Register Bangla Font
    const fontPath = path.join(__dirname, 'fonts', 'kalpurush.ttf');
    doc.registerFont('BanglaFont', fontPath);

    // Document Header - Title & Subject
    doc.font('BanglaFont').fontSize(18).fillColor('#0f172a').text(testName, { align: 'center' });
    doc.moveDown(0.2);
    doc.fontSize(11).fillColor('#475569').text(`বিষয়: ${subject}`, { align: 'center' });
    doc.moveDown(0.3);

    // Document Header - Full Marks & Time Duration Line
    const headerInfoY = doc.y;
    doc.fontSize(10).fillColor('#64748b').text(`পূর্ণমান: ${fullMarks}`, 40, headerInfoY, { align: 'left' });
    doc.fontSize(10).fillColor('#64748b').text(`সময়: ${duration}`, 40, headerInfoY, { align: 'right' });
    doc.moveDown(0.6);

    // Header Divider Line
    doc.strokeColor('#cbd5e1').lineWidth(1).moveTo(40, doc.y).lineTo(555, doc.y).stroke();
    doc.moveDown(1.2);

    // Render Questions
    rawData.forEach((q, index) => {
      const qText = q.question_text || q.Question || q.question || '';
      const optA = q.option_a || q['Option A'] || q.a || '';
      const optB = q.option_b || q['Option B'] || q.b || '';
      const optC = q.option_c || q['Option C'] || q.c || '';
      const optD = q.option_d || q['Option D'] || q.d || '';
      
      const correctOpt = String(q.correct_option || q.Answer || q.correct || '').trim().toUpperCase();
      const explanation = (q['Explanation '] || q.Explanation || q.explanation || '').trim();

      if (qText) {
        doc.font('BanglaFont').fontSize(11).fillColor('#0f172a').text(`প্রশ্ন ${index + 1}. ${qText}`, { lineGap: 4 });
        doc.moveDown(0.3);

        const options = [
          { key: 'A', label: '(A)', val: optA },
          { key: 'B', label: '(B)', val: optB },
          { key: 'C', label: '(C)', val: optC },
          { key: 'D', label: '(D)', val: optD }
        ];

        const optionsStr = options.map(opt => {
          const isCorrect = (opt.key === correctOpt);
          return isCorrect ? `[ ${opt.label} ${opt.val} ]` : `${opt.label} ${opt.val}`;
        }).join('     ');

        doc.font('BanglaFont').fontSize(10).fillColor('#334155').text(optionsStr, { indent: 15, lineGap: 4 });
        doc.moveDown(0.4);

        if (explanation) {
          doc.font('BanglaFont').fontSize(9.5).fillColor('#047857').text(`ব্যাখ্যা: ${explanation}`, { indent: 15, lineGap: 4 });
        }

        doc.moveDown(1.2);
      }
    });

    doc.end();
  } catch (error) {
    res.status(500).send('Error generating question paper PDF: ' + error.message);
  }
});

// 3. Excel to Word (.docx) Endpoint
app.post('/api/convert-excel-to-word', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('Please upload an Excel file.');
    }

    const testName = req.body.testName || 'ঢাবি ভর্তি পরীক্ষা : প্রশ্নপত্র ও সমাধান';
    const subject = req.body.subject || 'C-Unit Admission Test';
    const fullMarks = req.body.fullMarks || '১০০';
    const duration = req.body.duration || '১ ঘণ্টা';

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    const docParagraphs = [
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: testName, bold: true, size: 32 })
        ],
        alignment: docx.AlignmentType.CENTER
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: `বিষয়: ${subject}`, color: "475569", size: 22 })
        ],
        alignment: docx.AlignmentType.CENTER
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: `পূর্ণমান: ${fullMarks}                    সময়: ${duration}`, color: "64748B", size: 20 })
        ],
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: 180 }
      })
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
        docParagraphs.push(
          new docx.Paragraph({
            children: [
              new docx.TextRun({ text: `প্রশ্ন ${index + 1}. ${qText}`, bold: true, size: 22 })
            ],
            spacing: { before: 120, after: 60 }
          })
        );

        const options = [
          { key: 'A', label: '(A)', val: optA },
          { key: 'B', label: '(B)', val: optB },
          { key: 'C', label: '(C)', val: optC },
          { key: 'D', label: '(D)', val: optD }
        ];

        const optionTextRuns = [];
        options.forEach((opt, idx) => {
          const isCorrect = (opt.key === correctOpt);
          optionTextRuns.push(
            new docx.TextRun({
              text: `${opt.label} ${opt.val}`,
              bold: isCorrect,
              underline: isCorrect ? { type: docx.UnderlineType.SINGLE } : undefined,
              size: 20
            })
          );
          if (idx < options.length - 1) {
            optionTextRuns.push(new docx.TextRun({ text: "     ", size: 20 }));
          }
        });

        docParagraphs.push(
          new docx.Paragraph({
            children: optionTextRuns,
            indent: { left: 360 },
            spacing: { after: 60 }
          })
        );

        if (explanation) {
          docParagraphs.push(
            new docx.Paragraph({
              children: [
                new docx.TextRun({ text: `ব্যাখ্যা: ${explanation}`, color: "047857", size: 18 })
              ],
              indent: { left: 360 },
              spacing: { after: 120 }
            })
          );
        }
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

// 4. UI Dashboard
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
    <body class="bg-slate-100 text-slate-800 min-h-screen font-sans pb-16">
      <header class="bg-white border-b border-slate-200 py-4 px-8 flex justify-between items-center shadow-sm">
        <div class="flex items-center gap-3">
          <div class="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center text-white font-bold text-xl shadow-md">📄</div>
          <div>
            <h1 class="text-xl font-bold text-slate-800 leading-tight">DocTransform Studio</h1>
            <p class="text-xs text-slate-500">Excel to PDF & Word Question Generator</p>
          </div>
        </div>
        <a href="/api/download-template" class="bg-indigo-50 hover:bg-indigo-100 text-indigo-600 border border-indigo-200 px-4 py-2 rounded-xl text-xs font-semibold flex items-center gap-2 transition-all">
          <span>📥</span> Download Excel Template
        </a>
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

        <!-- Header Info Customization Form Card -->
        <div class="bg-white rounded-3xl border border-slate-200 p-8 shadow-sm mb-6">
          <h3 class="text-md font-bold text-slate-800 mb-4 flex items-center gap-2">
            <span>⚙️</span> প্রশ্নপত্রের হেডার বিবরণি (Header Information)
          </h3>
          <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <label class="block text-xs font-semibold text-slate-600 mb-1">পরীক্ষার নাম (Test Name)</label>
              <input type="text" form="convertForm" name="testName" value="ঢাবি ভর্তি পরীক্ষা : প্রশ্নপত্র ও সমাধান" class="w-full px-4 py-2 text-sm border border-slate-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:outline-none">
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-600 mb-1">বিষয় (Subject)</label>
              <input type="text" form="convertForm" name="subject" value="C-Unit Admission Test" class="w-full px-4 py-2 text-sm border border-slate-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:outline-none">
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-600 mb-1">পূর্ণমান (Full Marks)</label>
              <input type="text" form="convertForm" name="fullMarks" value="১০০" class="w-full px-4 py-2 text-sm border border-slate-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:outline-none">
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-600 mb-1">সময় (Time Duration)</label>
              <input type="text" form="convertForm" name="duration" value="১ ঘণ্টা" class="w-full px-4 py-2 text-sm border border-slate-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:outline-none">
            </div>
          </div>
        </div>

        <!-- File Upload Form Card -->
        <div class="bg-white rounded-3xl border border-slate-200 p-10 shadow-sm text-center mb-8">
          <form id="convertForm" onsubmit="handleFormSubmit(event)" enctype="multipart/form-data">
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
                Generate & Preview
              </button>
            </div>
          </form>
        </div>

        <!-- Live PDF Show Section -->
        <div id="pdfShowSection" class="hidden bg-white rounded-3xl border border-slate-200 p-6 shadow-sm">
          <div class="flex items-center justify-between mb-4">
            <h3 class="text-lg font-bold text-slate-800 flex items-center gap-2">
              <span>👁️</span> PDF Live Preview & Download
            </h3>
            <a id="downloadBtn" href="#" download="question-paper.pdf" class="bg-indigo-600 hover:bg-indigo-700 text-white font-semibold text-xs px-5 py-2.5 rounded-xl shadow-sm transition-all flex items-center gap-1.5">
              <span>⬇️</span> Download File
            </a>
          </div>
          <div class="w-full h-[650px] rounded-2xl overflow-hidden border border-slate-200 bg-slate-50">
            <iframe id="pdfFrame" class="w-full h-full border-none"></iframe>
          </div>
        </div>
      </main>

      <script>
        let currentMode = 'excel';

        function switchMode(mode) {
          currentMode = mode;
          const fileInput = document.getElementById('fileInput');
          const excelTab = document.getElementById('excelTab');
          const wordTab = document.getElementById('wordTab');
          const dragText = document.getElementById('dragText');
          const subText = document.getElementById('subText');
          const fileInfo = document.getElementById('fileInfo');
          const submitBtn = document.getElementById('submitBtn');
          const pdfShowSection = document.getElementById('pdfShowSection');
          
          fileInput.value = '';
          fileInfo.classList.add('hidden');
          pdfShowSection.classList.add('hidden');

          if (mode === 'excel') {
            fileInput.accept = '.xlsx, .xls';
            excelTab.className = 'bg-white p-6 rounded-2xl border-2 border-indigo-600 shadow-sm hover:shadow-md transition-all text-left flex items-start gap-4 cursor-pointer';
            wordTab.className = 'bg-white p-6 rounded-2xl border border-slate-200 opacity-60 hover:opacity-100 transition-all text-left flex items-start gap-4 cursor-pointer';
            dragText.textContent = 'Excel প্রশ্ন ফাইল আপলোড করুন';
            subText.textContent = 'ফাইল সিলেক্ট করে সুন্দর Question Paper PDF বানিয়ে নিন';
            submitBtn.textContent = 'Generate & Preview PDF';
          } else {
            fileInput.accept = '.xlsx, .xls';
            wordTab.className = 'bg-white p-6 rounded-2xl border-2 border-indigo-600 shadow-sm hover:shadow-md transition-all text-left flex items-start gap-4 cursor-pointer';
            excelTab.className = 'bg-white p-6 rounded-2xl border border-slate-200 opacity-60 hover:opacity-100 transition-all text-left flex items-start gap-4 cursor-pointer';
            dragText.textContent = 'Excel প্রশ্ন ফাইল আপলোড করুন';
            subText.textContent = 'ফাইল সিলেক্ট করে Word Document (.docx) বানিয়ে নিন';
            submitBtn.textContent = 'Generate & Download Word';
          }
        }

        function handleFileChange(input) {
          if (input.files && input.files[0]) {
            document.getElementById('selectedFileName').textContent = input.files[0].name;
            document.getElementById('fileInfo').classList.remove('hidden');
          }
        }

        async function handleFormSubmit(event) {
          event.preventDefault();
          const form = document.getElementById('convertForm');
          const formData = new FormData(form);
          const endpoint = currentMode === 'excel' ? '/api/convert-excel' : '/api/convert-excel-to-word';
          const submitBtn = document.getElementById('submitBtn');

          submitBtn.disabled = true;
          submitBtn.textContent = 'Processing...';

          try {
            const response = await fetch(endpoint, {
              method: 'POST',
              body: formData
            });

            if (!response.ok) throw new Error('Conversion failed');

            const blob = await response.blob();
            const blobUrl = URL.createObjectURL(blob);
            
            const pdfShowSection = document.getElementById('pdfShowSection');
            const pdfFrame = document.getElementById('pdfFrame');
            const downloadBtn = document.getElementById('downloadBtn');

            if (currentMode === 'excel') {
              pdfFrame.src = blobUrl;
              downloadBtn.href = blobUrl;
              downloadBtn.download = 'question-paper.pdf';
              downloadBtn.textContent = '⬇️ Download PDF File';
              pdfShowSection.classList.remove('hidden');
              pdfShowSection.scrollIntoView({ behavior: 'smooth' });
            } else {
              downloadBtn.href = blobUrl;
              downloadBtn.download = 'question-paper.docx';
              downloadBtn.textContent = '⬇️ Download Word File';
              pdfShowSection.classList.add('hidden');
              const a = document.createElement('a');
              a.href = blobUrl;
              a.download = 'question-paper.docx';
              a.click();
            }
          } catch (err) {
            alert('Error generating document: ' + err.message);
          } finally {
            submitBtn.disabled = false;
            submitBtn.textContent = currentMode === 'excel' ? 'Generate & Preview PDF' : 'Generate & Download Word';
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