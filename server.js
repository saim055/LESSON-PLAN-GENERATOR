// ================= IMPORTS & SETUP (same as before) =================
const express = require("express");
const fs = require("fs");
const path = require("path");
const cors = require("cors");
const multer = require("multer");
const pdf = require("pdf-parse");
const mammoth = require("mammoth");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
require("dotenv").config();

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(express.json({ limit: "50mb" }));

const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir, { recursive: true });

const upload = multer({ dest: uploadsDir });

// Groq Client + Multi-Model System (same powerful version as last)
const Groq = require("groq-sdk");
const client = new Groq({ apiKey: process.env.GROQ_API_KEY });

const MODEL_QUEUE = [
  "meta-llama/llama-4-scout-17b-16e-instruct",
  "llama-3.3-70b-versatile",
  "qwen/qwen3-32b",
  "llama-3.1-70b-versatile",
  "qwen/qwen3.6-27b"
];

const modelCooldowns = {};
const requestCache = new Map();

MODEL_QUEUE.forEach(m => { if (!modelCooldowns[m]) modelCooldowns[m] = 0; });

// ... (keep the callAIWithFallback, safe(), getMonthlyValue(), extractFileContent() functions same as previous version)

// ================= EXPANDED STANDARDS (same as last) =================
const STANDARDS_FRAMEWORK = { /* same as previous response */ };
function getStandardsFramework(subject) { /* same as previous */ }

// ================= FINAL STRONG PROMPT =================
const EXPERT_SYSTEM_PROMPT = `You are a master American curriculum designer.

Create **outstanding, highly detailed, and classroom-ready** lesson plans.

Focus especially on making:
- Cooperative Tasks: 3 truly different activities (Support, Average, Upper) with clear instructions, scaffolds, and deliverables.
- Independent Tasks: 3 different tasks (not repeating cooperative ones) with strong scaffolding for support and high challenge for upper level.

Be subject-specific:
- Math/Science: problem-solving, inquiry, calculations
- English/Public Speaking: discussion, presentation, analysis
- Arts/Digital Arts: creation, critique, technique
- PE: skills practice, fitness, teamwork
- Economics/Business: real-world cases, decision making
- Geography: maps, analysis, connections

Always maintain excellent differentiation.`;

const DOK_PROFILE = {
  introductory: ["DOK1", "DOK2", "DOK3"],
  intermediate: ["DOK2", "DOK3", "DOK4"],
  mastery: ["DOK3", "DOK4", "DOK4"]
};

// ================= MAIN API =================
app.post("/api/generate", upload.single("file"), async (req, res) => {
  try {
    const { subject, grade, topic, level, period, date, semester, giftedTalented } = req.body;

    if (!subject || !grade || !topic || !level) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    const standardsFramework = getStandardsFramework(subject);
    const dokLevels = DOK_PROFILE[level.toLowerCase()] || DOK_PROFILE.introductory;

    let syllabusContent = "";
    if (req.file) {
      syllabusContent = await extractFileContent(req.file.path);
      fs.unlinkSync(req.file.path).catch(() => {});
    }

    const userPrompt = `Generate a high-quality lesson plan for:
Subject: ${subject}
Grade: ${grade}
Topic: ${topic}
Level: ${level}
Standards: ${standardsFramework}

${syllabusContent ? `Syllabus Context: ${syllabusContent}` : ''}
${giftedTalented === 'yes' ? 'Include advanced objective for gifted students.' : ''}

Pay special attention to creating **varied, engaging, and well-differentiated** Cooperative Tasks and Independent Tasks.`;

    const result = await callAIWithFallback(EXPERT_SYSTEM_PROMPT, userPrompt);
    const aiData = result.parsed;

    const templateData = {
      date: safe(new Date(date || Date.now()).toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' })),
      semester: safe(semester || '1'),
      grade: safe(grade),
      subject: safe(subject),
      topic: safe(topic),
      period: safe(period || '1'),
      value: safe(getMonthlyValue(date)),

      standardText: safe(aiData.standardText),
      objective1: safe(aiData.objectives?.[0]?.text),
      objective2: safe(aiData.objectives?.[1]?.text),
      objective3: safe(aiData.objectives?.[2]?.text),

      outcomeAll: safe(aiData.outcomes?.all?.text),
      outcomeMost: safe(aiData.outcomes?.most?.text),
      outcomeSome: safe(aiData.outcomes?.some?.text),

      vocabulary: safe(Array.isArray(aiData.vocabulary) ? aiData.vocabulary.join('\n') : aiData.vocabulary),
      resources: safe(Array.isArray(aiData.resources) ? aiData.resources.join('\n') : aiData.resources),
      skills: safe(aiData.skills),

      starter: safe(aiData.starter),
      teaching: safe(aiData.teaching),

      // Cooperative Tasks
      coopSupport: safe(aiData.cooperative?.support),
      coopAverage: safe(aiData.cooperative?.average),
      coopUpper: safe(aiData.cooperative?.upper),

      // Independent Tasks
      indepSupport: safe(aiData.independent?.support),
      indepAverage: safe(aiData.independent?.average),
      indepUpper: safe(aiData.independent?.upper),

      plenary: safe(Array.isArray(aiData.plenary) ? aiData.plenary.map((p,i) => `${i+1}. (${p.dok}) ${p.q}`).join('\n') : aiData.plenary),

      myIdentity: safe(aiData.identity ? `Domain: ${aiData.identity.domain} - Element: ${aiData.identity.element}\n\n${aiData.identity.description}` : ''),
      moralEducation: safe(aiData.moralEducation),
      steam: safe(aiData.steam),
      linksToSubjects: safe(aiData.linksToSubjects),
      environment: safe(aiData.environment),
      realWorld: safe(aiData.realWorld),
      alnObjectives: giftedTalented === 'yes' ? safe(aiData.alnObjective) : ''
    };

    // Template Rendering (unchanged)
    const templatePath = path.join(__dirname, 'LESSON PLAN TEMPLATE.docx');
    const zip = new PizZip(fs.readFileSync(templatePath));
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

    doc.setData(templateData);
    doc.render();

    const buffer = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="Lesson_Plan_${subject.replace(/\s+/g,'_')}_${topic.replace(/\s+/g,'_')}.docx"`);
    res.send(buffer);

    console.log(`✅ Successfully generated: ${subject} - ${topic}`);

  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to generate lesson", details: error.message });
  }
});

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'enhanced-lesson-planner.html')));

app.get('/api/test', (req, res) => res.json({ status: 'READY - Cooperative & Independent Tasks Optimized' }));

app.listen(PORT, '0.0.0.0', () => {
  console.log('═══════════════════════════════════════════════');
  console.log('   WAAO POWERFUL LESSON PLANNER - FINAL VERSION');
  console.log('═══════════════════════════════════════════════');
  console.log(`Server running → http://localhost:${PORT}`);
  console.log('Cooperative & Independent tasks optimized for all subjects');
  console.log('═══════════════════════════════════════════════');
});
