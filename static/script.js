// ===== MODE SWITCHING =====
function showMode(mode) {
  const sections = ['modeSelector', 'customSection', 'aiSection', 'pdfSection', 'pptSection', 'aiPreview'];
  sections.forEach(s => { const el = document.getElementById(s); if (el) el.classList.add('hidden'); });
  if (mode === 'none') {
    document.getElementById('modeSelector').classList.remove('hidden');
  } else if (mode === 'custom') {
    document.getElementById('customSection').classList.remove('hidden');
    if (questionCount === 0) addQuestion();
  } else if (mode === 'ai') {
    document.getElementById('aiSection').classList.remove('hidden');
  } else if (mode === 'pdf') {
    document.getElementById('pdfSection').classList.remove('hidden');
  } else if (mode === 'ppt') {
    document.getElementById('pptSection').classList.remove('hidden');
  } else if (mode === 'exam') {
    document.getElementById('examSection').classList.remove('hidden');
    if (document.getElementById('examSectionsContainer').children.length === 0) addExamSection();
  }
}

function updateFileName() {
  const fileInput = document.getElementById('pdfFile');
  const display = document.getElementById('fileNameDisplay');
  if (fileInput.files.length > 0) {
    display.innerText = "📄 " + fileInput.files[0].name;
    display.style.color = "var(--accent)";
  } else {
    display.innerText = "Click to select or drag and drop your PDF here";
    display.style.color = "var(--text-secondary)";
  }
}

function updatePPTFileName() {
  const fileInput = document.getElementById('pptFile');
  const display = document.getElementById('pptFileNameDisplay');
  if (fileInput.files.length > 0) {
    display.innerText = "📊 " + fileInput.files[0].name;
    display.style.color = "#d97706";
  } else {
    display.innerText = "Click to select or drag and drop your PPTX here";
    display.style.color = "var(--text-secondary)";
  }
}

function updateExamFileName() {
  const fileInput = document.getElementById('examFile');
  const display = document.getElementById('examFileNameDisplay');
  if (fileInput.files.length > 0) {
    display.innerText = "📄 " + fileInput.files[0].name;
    display.style.color = "var(--accent)";
  } else {
    display.innerText = "Select or drop the material to generate from";
    display.style.color = "var(--text-secondary)";
  }
}

// ===== DYNAMIC QUESTION BUILDER =====
let questionCount = 0;

function getQuizType() {
  const typeSelect = document.getElementById('customType');
  return typeSelect ? typeSelect.value : 'MCQ';
}

function addQuestion() {
  const container = document.getElementById('questionsContainer');
  if (!container) return;
  questionCount++;
  const qType = getQuizType();
  const qDiv = document.createElement('div');
  qDiv.className = 'question-block';
  qDiv.id = `question-${questionCount}`;
  qDiv.setAttribute('data-qnum', questionCount);

  let optionsHTML = '';
  if (qType === 'MCQ') {
    optionsHTML = `
      <div class="options-grid">
        <div class="form-group"><label class="form-label">Option A</label><input type="text" class="form-input opt-input" data-opt="A" placeholder="Option A" required></div>
        <div class="form-group"><label class="form-label">Option B</label><input type="text" class="form-input opt-input" data-opt="B" placeholder="Option B" required></div>
        <div class="form-group"><label class="form-label">Option C</label><input type="text" class="form-input opt-input" data-opt="C" placeholder="Option C" required></div>
        <div class="form-group"><label class="form-label">Option D</label><input type="text" class="form-input opt-input" data-opt="D" placeholder="Option D" required></div>
      </div>
      <div class="form-group">
        <label class="form-label">Correct Answer</label>
        <select class="form-select correct-answer-select"><option value="A">A</option><option value="B">B</option><option value="C">C</option><option value="D">D</option></select>
      </div>`;
  } else if (qType === 'Rapid Fire') {
    optionsHTML = `<div class="form-group"><label class="form-label">Correct Answer</label><input type="text" class="form-input correct-answer-text" placeholder="Short answer" required></div>`;
  } else if (qType === 'Fill in the Blank') {
    optionsHTML = `<div class="form-group"><label class="form-label">Answer (the blank)</label><input type="text" class="form-input correct-answer-text" placeholder="Word or phrase" required></div>`;
  } else if (qType === 'Coding') {
    optionsHTML = `<div class="form-group"><label class="form-label">Expected Output / Answer</label><textarea class="form-input correct-answer-text" rows="3" required></textarea></div>`;
  }

  qDiv.innerHTML = `
    <div class="question-header">
      <span class="question-number">Question ${questionCount}</span>
      <button type="button" class="btn-remove" onclick="removeQuestion(${questionCount})">✕</button>
    </div>
    <div class="form-group">
      <label class="form-label">Question</label>
      <input type="text" class="form-input question-text" placeholder="Enter your question" required>
    </div>
    ${optionsHTML}`;
  container.appendChild(qDiv);
  qDiv.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function removeQuestion(num) {
  const el = document.getElementById(`question-${num}`);
  if (el) el.remove();
  document.querySelectorAll('.question-block').forEach((block, i) => {
    block.querySelector('.question-number').textContent = `Question ${i + 1}`;
  });
}

function updateQuestionTemplate() {
  const container = document.getElementById('questionsContainer');
  if (container) { container.innerHTML = ''; questionCount = 0; addQuestion(); }
}

function prepareCustomSubmit() {
  const blocks = document.querySelectorAll('.question-block');
  if (blocks.length === 0) { alert('Please add at least one question!'); return false; }
  const qType = getQuizType();
  const questions = [];
  for (const block of blocks) {
    const questionText = block.querySelector('.question-text').value.trim();
    if (!questionText) { alert('Please fill in all question fields!'); return false; }
    const q = { question: questionText };
    if (qType === 'MCQ') {
      const opts = block.querySelectorAll('.opt-input');
      q.options = [];
      for (const opt of opts) {
        if (!opt.value.trim()) { alert('Please fill in all options!'); return false; }
        q.options.push(opt.value.trim());
      }
      const selVal = block.querySelector('.correct-answer-select').value;
      const idx = ['A','B','C','D'].indexOf(selVal);
      q.answer = idx >= 0 && q.options[idx] ? q.options[idx] : selVal;
    } else {
      const answerEl = block.querySelector('.correct-answer-text');
      if (!answerEl.value.trim()) { alert('Please fill in all answer fields!'); return false; }
      q.answer = answerEl.value.trim();
    }
    questions.push(q);
  }
  document.getElementById('questionsJsonInput').value = JSON.stringify(questions);
  return true;
}

// ===== AI QUIZ GENERATION =====
let aiGeneratedQuestions = [];

async function generateAIQuiz() {
  const title = document.getElementById('aiTitle').value.trim();
  const subject = document.getElementById('aiSubject').value.trim();
  const qType = document.getElementById('aiType').value;
  const count = parseInt(document.getElementById('aiCount').value) || 5;
  const difficulty = document.getElementById('aiDifficulty').value;
  if (!subject) { alert('Please enter a subject!'); return; }
  if (!title) { alert('Please enter a quiz name!'); return; }

  document.getElementById('aiLoading').classList.remove('hidden');
  document.getElementById('generateBtn').disabled = true;
  document.getElementById('aiPreview').classList.add('hidden');

  try {
    const response = await fetch('/generate-ai', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ subject, type: qType, count, difficulty })
    });
    const data = await response.json();
    if (data.error) { alert('Error: ' + data.error); return; }
    aiGeneratedQuestions = data.questions;
    renderAIPreview(aiGeneratedQuestions, qType);
  } catch (err) {
    alert('Failed to generate quiz: ' + err.message);
  } finally {
    document.getElementById('aiLoading').classList.add('hidden');
    document.getElementById('generateBtn').disabled = false;
  }
}

async function generatePDFQuiz() {
  const fileInput = document.getElementById('pdfFile');
  const type = document.getElementById('pdfType').value;
  const count = document.getElementById('pdfCount').value;
  const difficulty = document.getElementById('pdfDifficulty').value;
  if (!fileInput.files[0]) { alert("Please select a PDF file first."); return; }
  const btn = document.getElementById('pdfGenerateBtn');
  const loading = document.getElementById('pdfLoading');
  btn.disabled = true; loading.classList.remove('hidden');
  const formData = new FormData();
  formData.append('pdf', fileInput.files[0]);
  formData.append('type', type); formData.append('count', count); formData.append('difficulty', difficulty);
  try {
    const res = await fetch('/generate-from-pdf', { method: 'POST', body: formData });
    if (!res.ok) throw new Error(`Server error ${res.status}`);
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    aiGeneratedQuestions = data.questions;
    if (!aiGeneratedQuestions || aiGeneratedQuestions.length === 0) throw new Error("No questions returned.");
    renderAIPreview(aiGeneratedQuestions, type);
  } catch (err) {
    alert("Generation Failed: " + err.message);
  } finally {
    btn.disabled = false; loading.classList.add('hidden');
  }
}

async function generatePPTQuiz() {
  const fileInput = document.getElementById('pptFile');
  const title = document.getElementById('pptTitle').value || "PPT Generated Quiz";
  const type = document.getElementById('pptType').value;
  const count = document.getElementById('pptCount').value;
  const difficulty = document.getElementById('pptDifficulty').value;
  if (!fileInput.files[0]) { alert("Please select a PPTX file first."); return; }
  const btn = document.getElementById('pptGenerateBtn');
  const loading = document.getElementById('pptLoading');
  btn.disabled = true; loading.classList.remove('hidden');
  const formData = new FormData();
  formData.append('ppt', fileInput.files[0]); formData.append('type', type);
  formData.append('count', count); formData.append('difficulty', difficulty);
  try {
    const res = await fetch('/generate-from-ppt', { method: 'POST', body: formData });
    if (!res.ok) throw new Error(`Server error ${res.status}`);
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    aiGeneratedQuestions = data.questions;
    if (!aiGeneratedQuestions || aiGeneratedQuestions.length === 0) throw new Error("No questions returned.");
    renderAIPreview(aiGeneratedQuestions, type);
  } catch (err) {
    alert("Generation Failed: " + err.message);
  } finally {
    btn.disabled = false; loading.classList.add('hidden');
  }
}

function renderAIPreview(questions, qType) {
  const container = document.getElementById('aiPreviewContent');
  container.innerHTML = '';
  questions.forEach((q, i) => {
    let optionsHTML = '';
    if (qType === 'MCQ' && q.options) {
      const labels = ['A','B','C','D'];
      optionsHTML = q.options.map((opt, j) => {
        const isCorrect = q.answer === labels[j] || q.answer === opt;
        return `<div class="preview-option ${isCorrect ? 'correct' : ''}">${labels[j]}. ${opt}</div>`;
      }).join('');
      optionsHTML += `<p class="preview-answer">✅ Correct: ${q.answer}</p>`;
    } else {
      optionsHTML = `<p class="preview-answer">✅ Answer: ${q.answer}</p>`;
    }
    container.innerHTML += `<div class="quiz-card preview-card"><p class="preview-question">Q${i+1}. ${q.question}</p>${optionsHTML}</div>`;
  });
  document.getElementById('aiPreview').classList.remove('hidden');
}

async function submitAIQuiz() {
  const isPDF = document.getElementById('pdfSection') && !document.getElementById('pdfSection').classList.contains('hidden');
  const isPPT = document.getElementById('pptSection') && !document.getElementById('pptSection').classList.contains('hidden');
  let title, subject, qType, qClass, openAt = '', closeAt = '';
  let maxAttempts = 0, shuffleQuestions = false, shuffleOptions = false, rapidFire = false;

  if (isPDF) {
    title = document.getElementById('pdfTitle').value.trim() || "PDF Generated Quiz";
    subject = document.getElementById('pdfSubject').value.trim() || "PDF Import";
    qType = document.getElementById('pdfType').value; qClass = document.getElementById('pdfClass').value;
    maxAttempts = parseInt(document.getElementById('pdfMaxAttempts').value) || 0;
    shuffleQuestions = document.getElementById('pdfShuffleQ').checked;
    shuffleOptions = document.getElementById('pdfShuffleOpt').checked;
    rapidFire = document.getElementById('pdfRapidFire') ? document.getElementById('pdfRapidFire').checked : false;
  } else if (isPPT) {
    title = document.getElementById('pptTitle').value.trim() || "PPT Generated Quiz";
    subject = document.getElementById('pptSubject').value.trim() || "PPT Import";
    qType = document.getElementById('pptType').value; qClass = document.getElementById('pptClass').value;
    maxAttempts = parseInt(document.getElementById('pptMaxAttempts').value) || 0;
    shuffleQuestions = document.getElementById('pptShuffleQ').checked;
    shuffleOptions = document.getElementById('pptShuffleOpt').checked;
    rapidFire = document.getElementById('pptRapidFire') ? document.getElementById('pptRapidFire').checked : false;
  } else {
    title = document.getElementById('aiTitle').value.trim() || "AI Generated Quiz";
    subject = document.getElementById('aiSubject').value.trim() || "General";
    qType = document.getElementById('aiType').value; qClass = document.getElementById('aiClass').value;
    openAt = document.getElementById('aiOpenAt') ? document.getElementById('aiOpenAt').value : '';
    closeAt = document.getElementById('aiCloseAt') ? document.getElementById('aiCloseAt').value : '';
    maxAttempts = parseInt(document.getElementById('aiMaxAttempts').value) || 0;
    shuffleQuestions = document.getElementById('aiShuffleQ').checked;
    shuffleOptions = document.getElementById('aiShuffleOpt').checked;
    rapidFire = document.getElementById('aiRapidFire') ? document.getElementById('aiRapidFire').checked : false;
  }

  if (aiGeneratedQuestions.length === 0) { alert('Please generate a quiz first!'); return; }

  try {
    const response = await fetch('/create_quiz', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ title, subject, type: qType, class: qClass, questions: aiGeneratedQuestions, open_at: openAt, close_at: closeAt, max_attempts: maxAttempts, shuffle_questions: shuffleQuestions, shuffle_options: shuffleOptions, rapid_fire: rapidFire })
    });
    const data = await response.json();
    if (data.success) { alert('Quiz created successfully!'); window.location.href = '/mentor'; }
    else { alert('Error: ' + (data.error || 'Unknown error')); }
  } catch (err) {
    alert('Failed to save quiz: ' + err.message);
  }
}

// ===== EXAM GENERATOR =====
let examSectionCount = 0;
function addExamSection() {
  const container = document.getElementById('examSectionsContainer');
  examSectionCount++;
  const div = document.createElement('div');
  div.className = 'settings-group section-block';
  div.style = "margin-top: 10px; background: rgba(255,255,255,0.03); border: 1px solid var(--border); padding: 16px; border-radius: 12px; position: relative;";
  div.innerHTML = `
    <button type="button" class="btn-remove" onclick="this.parentElement.remove()" style="position:absolute; top:12px; right:12px;">✕</button>
    <div class="form-row">
      <div class="form-group" style="flex:2;">
        <label class="form-label">Section Name</label>
        <input type="text" class="form-input exam-sec-name" placeholder="e.g. Section A: Theory">
      </div>
      <div class="form-group">
        <label class="form-label">Type</label>
        <select class="form-select exam-sec-type">
          <option value="MCQ">MCQ</option>
          <option value="Short Answer">Short Answer</option>
          <option value="Long Answer">Long Answer</option>
          <option value="Coding">Coding</option>
        </select>
      </div>
    </div>
    <div class="form-row" style="margin-top: 10px;">
      <div class="form-group">
        <label class="form-label">Questions</label>
        <input type="number" class="form-input exam-sec-count" value="5" min="1" max="50">
      </div>
      <div class="form-group">
        <label class="form-label">Marks Per Q</label>
        <input type="number" class="form-input exam-sec-marks" value="2" min="1">
      </div>
      <div class="form-group">
        <label class="form-label">Difficulty</label>
        <select class="form-select exam-sec-diff">
          <option value="Easy">Easy</option>
          <option value="Medium" selected>Medium</option>
          <option value="Hard">Hard</option>
        </select>
      </div>
    </div>
  `;
  container.appendChild(div);
}

// ── Shared helper: collect exam form values ───────────────────────────────────
function collectExamFormData() {
  const fileInput = document.getElementById('examFile');
  if (!fileInput.files[0]) { alert("Please upload a reference PDF or PPTX first!"); return null; }

  const sections = [];
  document.querySelectorAll('.section-block').forEach(block => {
    sections.push({
      name:       block.querySelector('.exam-sec-name').value  || "Section",
      type:       block.querySelector('.exam-sec-type').value,
      count:      block.querySelector('.exam-sec-count').value,
      marks:      block.querySelector('.exam-sec-marks').value,
      difficulty: block.querySelector('.exam-sec-diff').value
    });
  });
  if (sections.length === 0) { alert("Please add at least one section!"); return null; }

  const header = {
    institution: document.getElementById('examInst').value,
    degree:      document.getElementById('examDegree').value,
    title:       document.getElementById('examTitle').value,
    subject:     document.getElementById('examSub').value,
    date:        document.getElementById('examDate').value,
    time:        document.getElementById('examTime').value,
    totalMarks:  document.getElementById('examTotalMarks').value
  };

  return { file: fileInput.files[0], header, sections };
}

// ── Generate question paper OR answer key (single file) ───────────────────────
async function generateExamPaper(mode = 'paper') {
  const formValues = collectExamFormData();
  if (!formValues) return;
  const { file, header, sections } = formValues;

  const loading     = document.getElementById('examLoading');
  const loadingText = loading.querySelector('p');
  const btnPaper    = document.getElementById('examPaperBtn');
  const btnKey      = document.getElementById('examKeyBtn');
  const btnBoth     = document.getElementById('examBothBtn');

  btnPaper.disabled = btnKey.disabled = btnBoth.disabled = true;
  loading.classList.remove('hidden');

  const formData = new FormData();
  formData.append('file',     file);
  formData.append('header',   JSON.stringify(header));
  formData.append('sections', JSON.stringify(sections));
  formData.append('mode',     mode);

  try {
    loadingText.innerText = "AI is generating questions — this may take a moment...";
    const res  = await fetch('/generate-exam', { method: 'POST', body: formData });
    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data.success) throw new Error(data.error || "Failed to generate document.");

    loadingText.innerText = "Downloading your file...";
    window.location.href  = data.url;
    alert(`✅ ${mode === 'key' ? 'Answer Key' : 'Question Paper'} is downloading!`);
  } catch (err) {
    alert("❌ Error: " + err.message);
  } finally {
    btnPaper.disabled = btnKey.disabled = btnBoth.disabled = false;
    loading.classList.add('hidden');
    loadingText.innerText = "AI is structuring your exam and generating files...";
  }
}

// ── Generate BOTH files as a single ZIP download ─────────────────────────────
async function generateBothDocs() {
  const formValues = collectExamFormData();
  if (!formValues) return;
  const { file, header, sections } = formValues;

  const loading     = document.getElementById('examLoading');
  const loadingText = loading.querySelector('p');
  const btnPaper    = document.getElementById('examPaperBtn');
  const btnKey      = document.getElementById('examKeyBtn');
  const btnBoth     = document.getElementById('examBothBtn');

  btnPaper.disabled = btnKey.disabled = btnBoth.disabled = true;
  loading.classList.remove('hidden');
  loadingText.innerText = "AI is generating questions — building both documents...";

  const formData = new FormData();
  formData.append('file',     file);
  formData.append('header',   JSON.stringify(header));
  formData.append('sections', JSON.stringify(sections));

  try {
    const res = await fetch('/generate-exam-both', { method: 'POST', body: formData });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || `Server error ${res.status}`);
    }

    // Stream the ZIP blob and trigger download
    loadingText.innerText = "Packaging ZIP — downloading now...";
    const blob     = await res.blob();
    const url      = URL.createObjectURL(blob);
    const a        = document.createElement('a');
    // Try to read filename from Content-Disposition header
    const cd       = res.headers.get('Content-Disposition') || '';
    const match    = cd.match(/filename[^;=\n]*=["']?([^"'\n;]+)/);
    a.download     = match ? match[1] : 'ExamPack.zip';
    a.href         = url;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    alert("✅ ZIP downloaded! It contains:\n• QuestionPaper_…docx\n• AnswerKey_…docx");
  } catch (err) {
    alert("❌ Error: " + err.message);
  } finally {
    btnPaper.disabled = btnKey.disabled = btnBoth.disabled = false;
    loading.classList.add('hidden');
    loadingText.innerText = "AI is structuring your exam and generating files...";
  }
}

// ===== DRAG AND DROP =====
document.addEventListener('DOMContentLoaded', () => {
  if (typeof questionCount !== 'undefined' && document.getElementById('modeSelector')) {
    showMode('none');
  }
  setupDragAndDrop('dropArea', 'pdfFile', updateFileName);
  setupDragAndDrop('pptDropArea', 'pptFile', updatePPTFileName);
  setupDragAndDrop('examDropArea', 'examFile', updateExamFileName);
});

function setupDragAndDrop(areaId, inputId, updateCallback) {
  const dropArea = document.getElementById(areaId);
  if (!dropArea) return;
  ['dragenter','dragover','dragleave','drop'].forEach(evt => {
    dropArea.addEventListener(evt, e => { e.preventDefault(); e.stopPropagation(); }, false);
  });
  ['dragenter','dragover'].forEach(evt => dropArea.addEventListener(evt, () => dropArea.classList.add('highlight'), false));
  ['dragleave','drop'].forEach(evt => dropArea.addEventListener(evt, () => dropArea.classList.remove('highlight'), false));
  dropArea.addEventListener('drop', e => {
    const files = e.dataTransfer.files;
    const fileInput = document.getElementById(inputId);
    if (fileInput && files.length > 0) { fileInput.files = files; updateCallback(); }
  }, false);
}
async function gradeAnswer() {
  const fileInput = document.querySelector("input[type='file']");
  const file = fileInput.files[0];

  if (!file) {
    alert("Please upload a file first");
    return;
  }

  const formData = new FormData();

  // MUST match Flask backend
  formData.append("pdf", file);
  formData.append("subject", document.getElementById("subject")?.value || "General");
  formData.append("totalMarks", document.getElementById("totalMarks")?.value || "30");
  formData.append("context", document.getElementById("context")?.value || "");
  formData.append("strictness", document.getElementById("strictness")?.value || "standard");
  formData.append("examName", document.getElementById("examName")?.value || "Exam");

  try {
    const res = await fetch("/check-answers", {
      method: "POST",
      body: formData
    });

    if (!res.ok) {
      throw new Error("Server error: " + res.status);
    }

    const data = await res.json();

    if (data.error) {
      alert(data.error);
      return;
    }

    console.log("RESULT:", data);
    alert("Grading completed! Check console.");

  } catch (err) {
    alert("Failed to fetch: " + err.message);
    console.error(err);
  }
}