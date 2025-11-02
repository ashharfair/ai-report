const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const cors = require("cors");
const PizZip = require("pizzip");
const { DOMParser, XMLSerializer } = require("xmldom");
const { GoogleGenerativeAI } = require("@google/generative-ai");

const app = express();
app.use(cors());
app.use(express.json());
const port = 3000;

// --- CONFIGURATION ---
const GOOGLE_API_KEY = "AIzaSyClmusFjx_AViVSfe5Tzmtj2gyYxFTIc4g"; // 🔒 store in env variable ideally
const genAI = new GoogleGenerativeAI(GOOGLE_API_KEY);
const model = genAI.getGenerativeModel({ model: "gemini-2.0-flash" });
// ensure uploadsDir exists (you already have this)
const uploadsDir = path.join(__dirname, "uploads");
if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir);

// multer storage — use originalname or a sanitized timestamped name
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, uploadsDir);
  },
  filename: function (req, file, cb) {
    // create a safer filename: timestamp + original name
    const safeName = `${Date.now()}-${file.originalname.replace(/\s+/g, "_")}`;
    cb(null, safeName);
  },
});
const upload = multer({ storage });

// Upload endpoint: return the saved filename (fileId)
app.post("/upload-template", upload.single("template"), async (req, res) => {
  if (!req.file) return res.status(400).send("No file uploaded.");

  // `req.file.filename` is the filename saved under uploadsDir
  const savedFilename = req.file.filename; // e.g. "169xxx-template.docx"
  try {
    const content = fs.readFileSync(path.join(uploadsDir, savedFilename), "binary");
    const zip = new PizZip(content);
    const xml = zip.file("word/document.xml").asText();

    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const paragraphs = Array.from(doc.getElementsByTagName("w:p"));
    const headings = [];

    paragraphs.forEach((p) => {
      const style = p.getElementsByTagName("w:pStyle")[0];
      const textNodes = p.getElementsByTagName("w:t");
      const text = Array.from(textNodes).map((t) => t.textContent).join("");
      if (style && /^Heading[1-3]$/i.test(style.getAttribute("w:val"))) {
        headings.push(text.trim());
      }
    });

    // Return the saved filename as the fileId (client will send this back)
    res.json({
      fields: headings,
      fileId: savedFilename,
    });
  } catch (err) {
    console.error("Error extracting headings:", err);
    res.status(500).send("Error processing document.");
  }
});

// ------------------------------------
// 2️⃣ Enhance + Insert Back into DOCX
// ------------------------------------
app.post("/generate-new-report", async (req, res) => {
  const { fileId, userInputs } = req.body;
  if (!fileId || !userInputs) return res.status(400).send("Missing fileId or userInputs.");
  let filePath = fileId;
  try {
    // --- 1) Call AI to polish (same as before) ---
    const prompt = `
Polish the following formal document JSON. return a JSON object with the exact same keys and structure as the original, but make it more professional, clear, and detailed. Return valid JSON only.

${JSON.stringify(userInputs, null, 2)}
`;
    const result = await model.generateContent(prompt);
    const response = await result.response;
    const responseText = await response.text();
    const jsonString = responseText.replace(/```json/g, "").replace(/```/g, "").trim();

    let enhancedData;
    try {
      enhancedData = JSON.parse(jsonString);
//       enhancedData = {
//   "Project:": [
//     "Carrier Ingestion Platform Development",
//   ],
//   "Deliverables:": [
//     "Integration of Far Eye API",
//     "Completion of UI/UX design for all application pages",
//   ],
//   "Blockers:": [
//     "Awaiting necessary deployment access privileges",
//   ],
//   "Next Steps:": [
//     "Conduct comprehensive end-to-end testing of the implemented features",
//   ],
// };
    } catch (err) {
      console.error("Failed to parse AI JSON:", jsonString);
      return res.status(500).send("AI response could not be parsed as JSON.");
    }

    // Normalize enhancedData keys: remove trailing punctuation like ":" and lowercase
    const normalizeKey = (k) =>
      k
        .toString()
        .trim()
        .replace(/[:\uFEFF\s]+$/g, "") // remove trailing colons and spaces
        .replace(/[^\w\s]/g, "") // remove other punctuation for safe matching
        .toLowerCase();

    const normalizedData = {};
    for (const [k, v] of Object.entries(enhancedData)) {
      normalizedData[normalizeKey(k)] = v;
    }

    // --- 2) Load DOCX and parse document.xml ---
    const content = fs.readFileSync(`uploads/${filePath}`, "binary");
    const zip = new PizZip(content);
    const xml = zip.file("word/document.xml").asText();

    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const doc = parser.parseFromString(xml, "text/xml");

    const paragraphs = Array.from(doc.getElementsByTagName("w:p"));

    // find a sample body paragraph (to clone style) - first non-heading paragraph
    let sampleBody = paragraphs.find((p) => {
      const style = p.getElementsByTagName("w:pStyle")[0];
      return !(style && /^Heading[1-6]$/i.test(style.getAttribute("w:val")));
    });

    // fallback to any paragraph if none found
    if (!sampleBody && paragraphs.length > 0) sampleBody = paragraphs[0];

    const matchesLog = [];

    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];

      // Extract visible text for this paragraph (concatenate all w:t)
      const textNodes = Array.from(p.getElementsByTagName("w:t"));
      const visibleText = textNodes.map((t) => t.textContent || "").join("").trim();

      // Determine heading style if exists
      const styleNode = p.getElementsByTagName("w:pStyle")[0];
      const pStyle = styleNode ? styleNode.getAttribute("w:val") : null;
      const isHeadingStyle = !!(pStyle && /^Heading[1-6]$/i.test(pStyle));

      // Normalize paragraph text for matching with keys
      const normalizedParaKey = normalizeKey(visibleText);

      // If heading style, try to match based on normalized heading
      let matchedKey = null;
      if (isHeadingStyle && normalizedData[normalizedParaKey]) {
        matchedKey = normalizedParaKey;
        matchesLog.push({ type: "heading-style", paraText: visibleText, key: matchedKey });
      } else {
        // if not heading style, also try exact visible text match (with or without trailing colon)
        // try a few variants to be robust
        const candidates = [
          visibleText,
          visibleText.replace(/[:\s]+$/g, ""),
          visibleText.toLowerCase(),
        ].map((s) => normalizeKey(s));

        for (const c of candidates) {
          if (normalizedData[c]) {
            matchedKey = c;
            matchesLog.push({ type: "visible-match", paraText: visibleText, key: matchedKey });
            break;
          }
        }
      }

      if (!matchedKey) continue; // no matching enhanced data for this paragraph

      const enhancedValue = normalizedData[matchedKey];
      // Build text to insert. enhancedValue may be string or array
      let insertionText = "";
      if (Array.isArray(enhancedValue)) insertionText = enhancedValue.join("\n");
      else insertionText = String(enhancedValue);

      // --- Insert or replace behavior ---
      if (isHeadingStyle) {
        // --- Replacement Logic ---
        // Find the next paragraph sibling and remove it if it's not a heading
        let nextP = p.nextSibling;
        while (nextP && nextP.nodeName !== "w:p") {
          nextP = nextP.nextSibling; // skip non-paragraph nodes
        }

        if (nextP) {
          const nextStyleNode = nextP.getElementsByTagName("w:pStyle")[0];
          const nextPStyle = nextStyleNode ? nextStyleNode.getAttribute("w:val") : null;
          const isNextPHeading = !!(nextPStyle && /^Heading[1-6]$/i.test(nextPStyle));

          if (!isNextPHeading) {
            nextP.parentNode.removeChild(nextP);
          }
        }

        // --- Create and Insert New Paragraph ---
        const newP = doc.createElement("w:p");
        // Optional: Apply a default style if needed, e.g., 'Normal'
        const pPr = doc.createElement("w:pPr");
        const pStyle = doc.createElement("w:pStyle");
        pStyle.setAttribute("w:val", "Normal"); // or your default body style
        pPr.appendChild(pStyle);
        newP.appendChild(pPr);

        const lines = insertionText.split(/\r?\n/);
        lines.forEach((line, liIdx) => {
          const newRun = doc.createElement("w:r");
          const newText = doc.createElement("w:t");
          if (/^\s|\s$/.test(line)) newText.setAttribute("xml:space", "preserve");
          newText.appendChild(doc.createTextNode(line));
          newRun.appendChild(newText);
          newP.appendChild(newRun);

          if (liIdx < lines.length - 1) {
            const brRun = doc.createElement("w:r");
            const br = doc.createElement("w:br");
            brRun.appendChild(br);
            newP.appendChild(brRun);
          }
        });

        const parent = p.parentNode;
        if (p.nextSibling) parent.insertBefore(newP, p.nextSibling);
        else parent.appendChild(newP);
      } else {
        // Not a styled heading — replace the next paragraph's runs (if exists),
        // otherwise insert a new paragraph after current
        const nextP = paragraphs[i + 1];
        if (nextP) {
          // Remove existing runs
          const runsToRemove = Array.from(nextP.getElementsByTagName("w:r"));
          for (const r of runsToRemove) r.parentNode.removeChild(r);

          // Build new runs like above
          const lines = insertionText.split(/\r?\n/);
          lines.forEach((line, liIdx) => {
            const newRun = doc.createElement("w:r");
            const newText = doc.createElement("w:t");
            if (/^\s|\s$/.test(line)) newText.setAttribute("xml:space", "preserve");
            newText.appendChild(doc.createTextNode(line));
            newRun.appendChild(newText);
            nextP.appendChild(newRun);
            if (liIdx < lines.length - 1) {
              const brRun = doc.createElement("w:r");
              const br = doc.createElement("w:br");
              brRun.appendChild(br);
              nextP.appendChild(brRun);
            }
          });
        } else {
          // Insert a new paragraph after current
          const newP = sampleBody ? sampleBody.cloneNode(true) : doc.createElement("w:p");
          const runs = Array.from(newP.getElementsByTagName("w:r"));
          for (const r of runs) r.parentNode.removeChild(r);

          const lines = insertionText.split(/\r?\n/);
          lines.forEach((line, liIdx) => {
            const newRun = doc.createElement("w:r");
            const newText = doc.createElement("w:t");
            if (/^\s|\s$/.test(line)) newText.setAttribute("xml:space", "preserve");
            newText.appendChild(doc.createTextNode(line));
            newRun.appendChild(newText);
            newP.appendChild(newRun);
            if (liIdx < lines.length - 1) {
              const brRun = doc.createElement("w:r");
              const br = doc.createElement("w:br");
              brRun.appendChild(br);
              newP.appendChild(brRun);
            }
          });

          const parent = p.parentNode;
          if (p.nextSibling) parent.insertBefore(newP, p.nextSibling);
          else parent.appendChild(newP);
        }
      }
    } // end for paragraphs

    // Optional: log matches for debugging
    console.log("DOCX -> matched insertions:", JSON.stringify(matchesLog, null, 2));

    // --- 3) serialize and save ---
    const updatedXml = serializer.serializeToString(doc);
    zip.file("word/document.xml", updatedXml);

    const outputBuffer = zip.generate({ type: "nodebuffer" });
    const outputFile = `enhanced-${Date.now()}.docx`;
    fs.writeFileSync(outputFile, outputBuffer);

    res.download(outputFile, (err) => {
      if (err) console.error("Download error:", err);
      fs.unlinkSync(outputFile);
    });
  } catch (err) {
    console.error("Error generating document:", err);
    res.status(500).send("Error generating enhanced document.");
  }
});

app.listen(port, () =>
  console.log(`🚀 Server running at http://localhost:${port}`)
);
