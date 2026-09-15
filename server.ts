import express from "express";
import path from "path";
import fs from "fs";
import { createServer as createViteServer } from "vite";
import * as XLSXModule from "xlsx";
const XLSX = (XLSXModule as any).default || XLSXModule;
import nodemailer from "nodemailer";

const app = express();
const PORT = process.env.PORT ? parseInt(process.env.PORT, 10) : 3000;

app.use(express.json());

// Ensure data directory exists for storing Excel workbook
const DATA_DIR = path.join(process.cwd(), "data");
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

const EXCEL_FILE_PATH = path.join(DATA_DIR, "portfolio_database.xlsx");

// Initial data for seeding Excel sheets if not existing
const INITIAL_PROJECTS = [
  {
    ID: "PRJ-001",
    Name: "Disaster Management System",
    Category: "Web Platforms",
    TechStack: "HTML5, CSS3, JavaScript, Node.js",
    Description: "Real-time edge event mapping and public emergency alert communication framework routing support assets with Gmail-based emergency alert integration.",
    LiveLink: "https://bespoke-palmier-314313.netlify.app/",
    GitHubLink: "https://github.com/shashankpersonal18-2006/Disaster-Management-System"
  },
  {
    ID: "PRJ-002",
    Name: "Quick Cook AI",
    Category: "Cloud Deployments",
    TechStack: "HTML5, CSS3, JavaScript, Cloudflare Workers, GPT API",
    Description: "High-availability distributed dynamic recipe ecosystem platform utilizing Cloudflare workers and OpenAI GPT API for personalized recipe recommendations.",
    LiveLink: "https://quickcook.shashank-personal18.workers.dev/",
    GitHubLink: "https://github.com/shashankpersonal18-2006/QuickCook-AI"
  }
];

const INITIAL_CERTIFICATIONS = [
  { ID: "CRT-001", Name: "Python Foundation Certificate", Issuer: "Infosys Springboard & Sololearn", Category: "Python", Year: "2024" },
  { ID: "CRT-002", Name: "SQL Databases (Intro & Intermediate)", Issuer: "Sololearn", Category: "Databases", Year: "2024" },
  { ID: "CRT-003", Name: "Cloud Computing Fundamentals", Issuer: "Simplilearn", Category: "Cloud", Year: "2024" },
  { ID: "CRT-004", Name: "AWS Cloud Practitioner", Issuer: "Simplilearn", Category: "Cloud", Year: "2025" },
  { ID: "CRT-005", Name: "Azure Fundamentals (Microsoft Elevate)", Issuer: "Microsoft & AICTE", Category: "Cloud", Year: "2025" }
];

const INITIAL_SKILLS = [
  { ID: "SKL-001", SkillName: "Frontend Engineering (HTML/CSS/JS)", Proficiency: "90%", Category: "Web Tech" },
  { ID: "SKL-002", SkillName: "Python Programming", Proficiency: "90%", Category: "Languages" },
  { ID: "SKL-003", SkillName: "AWS Cloud Practitioner", Proficiency: "80%", Category: "Cloud" },
  { ID: "SKL-004", SkillName: "Relational Databases & SQL", Proficiency: "85%", Category: "Databases" },
  { ID: "SKL-005", SkillName: "Java & C++ Programming", Proficiency: "80%", Category: "Languages" },
  { ID: "SKL-006", SkillName: "Microsoft Azure Fundamentals", Proficiency: "85%", Category: "Cloud" }
];

const INITIAL_INTERNSHIPS = [
  { ID: "INT-001", Role: "Web Development Intern", Organization: "UpToSkills", Duration: "Internship", Description: "Developed responsive web interfaces, enhanced usability, applied Git/GitHub version control." },
  { ID: "INT-002", Role: "Microsoft Elevate Intern", Organization: "AICTE & Microsoft", Duration: "Internship", Description: "Acquired foundational knowledge of Azure services, architecture, security, governance, and cloud pricing." }
];

const INITIAL_SUBMISSIONS = [
  {
    ID: "SUB-001",
    Timestamp: "2026-08-01 10:30:00",
    Name: "Academic Reviewer",
    Email: "reviewer@panimalar.edu.in",
    Phone: "+91 9876543210",
    Message: "Impressive portfolio and Cloud project architecture!",
    GmailStatus: "Sent to shashanksenthil2006@gmail.com"
  }
];

// Function to initialize or read the Excel Database
function getOrInitWorkbook(): XLSXModule.WorkBook {
  if (fs.existsSync(EXCEL_FILE_PATH)) {
    try {
      return XLSX.readFile(EXCEL_FILE_PATH);
    } catch (e) {
      console.error("Error reading existing Excel workbook, recreating:", e);
    }
  }

  const wb = XLSX.utils.book_new();

  const wsSubmissions = XLSX.utils.json_to_sheet(INITIAL_SUBMISSIONS);
  const wsProjects = XLSX.utils.json_to_sheet(INITIAL_PROJECTS);
  const wsCertifications = XLSX.utils.json_to_sheet(INITIAL_CERTIFICATIONS);
  const wsSkills = XLSX.utils.json_to_sheet(INITIAL_SKILLS);
  const wsInternships = XLSX.utils.json_to_sheet(INITIAL_INTERNSHIPS);

  XLSX.utils.book_append_sheet(wb, wsSubmissions, "Contact Submissions");
  XLSX.utils.book_append_sheet(wb, wsProjects, "Projects");
  XLSX.utils.book_append_sheet(wb, wsCertifications, "Certifications");
  XLSX.utils.book_append_sheet(wb, wsSkills, "Skills");
  XLSX.utils.book_append_sheet(wb, wsInternships, "Internships");

  XLSX.writeFile(wb, EXCEL_FILE_PATH);
  return wb;
}

// Function to append a contact submission to the Excel sheet
function saveSubmissionToExcel(submission: {
  name: string;
  email: string;
  phone: string;
  message: string;
  gmailStatus: string;
}) {
  const wb = getOrInitWorkbook();
  const sheetName = "Contact Submissions";
  let submissions: any[] = [];

  if (wb.SheetNames.includes(sheetName)) {
    submissions = XLSX.utils.sheet_to_json(wb.Sheets[sheetName]);
  }

  const rowId = `SUB-${String(submissions.length + 1).padStart(3, "0")}`;
  const now = new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" });

  const newEntry = {
    ID: rowId,
    Timestamp: now,
    Name: submission.name,
    Email: submission.email,
    Phone: submission.phone || "N/A",
    Message: submission.message,
    GmailStatus: submission.gmailStatus
  };

  submissions.push(newEntry);

  const newWs = XLSX.utils.json_to_sheet(submissions);
  wb.Sheets[sheetName] = newWs;
  if (!wb.SheetNames.includes(sheetName)) {
    XLSX.utils.book_append_sheet(wb, newWs, sheetName);
  }

  XLSX.writeFile(wb, EXCEL_FILE_PATH);
  return newEntry;
}

// API Routes

// Health check endpoint
app.get("/api/health", (req, res) => {
  res.json({
    status: "ok",
    excelDatabase: fs.existsSync(EXCEL_FILE_PATH) ? "Connected" : "Initialized",
    targetEmail: "shashanksenthil2006@gmail.com"
  });
});

// GET Excel Database Content in JSON format
app.get("/api/excel/database", (req, res) => {
  try {
    const wb = getOrInitWorkbook();
    const data: Record<string, any[]> = {};

    wb.SheetNames.forEach((name) => {
      data[name] = XLSX.utils.sheet_to_json(wb.Sheets[name]);
    });

    res.json({
      success: true,
      filePath: EXCEL_FILE_PATH,
      lastModified: fs.statSync(EXCEL_FILE_PATH).mtime,
      sheets: data
    });
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// Download Excel File directly (.xlsx)
app.get("/api/excel/download", (req, res) => {
  try {
    getOrInitWorkbook();
    res.download(EXCEL_FILE_PATH, "Shashank_Portfolio_Database.xlsx");
  } catch (error: any) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// Send Contact Message API -> Logs to Excel + Transmits to Gmail
app.post("/api/contact/send", async (req, res) => {
  const { name, email, phone, message } = req.body;

  if (!name || !email || !message) {
    return res.status(400).json({
      success: false,
      error: "Please provide Name, Email, and Message."
    });
  }

  const targetEmail = "shashanksenthil2006@gmail.com";
  let gmailStatus = "Simulated Dispatch to " + targetEmail;
  let emailSent = false;

  // Try real SMTP dispatch if credentials exist in env
  const gmailUser = process.env.GMAIL_USER || targetEmail;
  const gmailPass = process.env.GMAIL_APP_PASSWORD;

  if (gmailPass) {
    try {
      const transporter = nodemailer.createTransport({
        service: "gmail",
        auth: {
          user: gmailUser,
          pass: gmailPass
        }
      });

      await transporter.sendMail({
        from: `"${name}" <${email}>`,
        to: targetEmail,
        replyTo: email,
        subject: `[Portfolio Contact] New Transmission from ${name}`,
        text: `Name: ${name}\nEmail: ${email}\nPhone: ${phone || "N/A"}\n\nMessage:\n${message}`,
        html: `
          <div style="font-family: Arial, sans-serif; padding: 20px; background-color: #f4f6f8; color: #1e293b;">
            <h2 style="color: #22c55e;">New Portfolio Contact Transmission</h2>
            <p><strong>Sender Name:</strong> ${name}</p>
            <p><strong>Email:</strong> ${email}</p>
            <p><strong>Phone:</strong> ${phone || "N/A"}</p>
            <hr style="border: 1px solid #cbd5e1; margin: 15px 0;" />
            <p><strong>Message Packet:</strong></p>
            <div style="background-color: #ffffff; padding: 15px; border-radius: 8px; border-left: 4px solid #22c55e;">
              ${message.replace(/\n/g, "<br/>")}
            </div>
            <p style="font-size: 12px; color: #64748b; margin-top: 20px;">Logged in Excel Database workbook at server runtime.</p>
          </div>
        `
      });

      gmailStatus = `Delivered to Gmail (${targetEmail}) via SMTP`;
      emailSent = true;
    } catch (err: any) {
      console.warn("SMTP Transmission warning, falling back to simulated dispatch:", err.message);
      gmailStatus = `Logged to Excel & Prepared for Gmail (${targetEmail})`;
    }
  } else {
    gmailStatus = `Logged to Excel & Prepared for Gmail (${targetEmail})`;
  }

  // Save entry to Excel workbook
  const excelRecord = saveSubmissionToExcel({
    name,
    email,
    phone,
    message,
    gmailStatus
  });

  res.json({
    success: true,
    message: `Message transmitted and logged successfully!`,
    details: {
      recordId: excelRecord.ID,
      timestamp: excelRecord.Timestamp,
      recipient: targetEmail,
      gmailStatus,
      emailSent,
      excelFile: "portfolio_database.xlsx"
    }
  });
});

// Explicit routes to serve resume PDF with proper attachment download headers
app.get(["/Shashank_S_Resume.pdf", "/Shashank%20S%20Resume.pdf"], (req, res) => {
  const publicPath = path.join(process.cwd(), "public", "Shashank_S_Resume.pdf");
  const distPath = path.join(process.cwd(), "dist", "Shashank_S_Resume.pdf");
  const filePath = fs.existsSync(publicPath) ? publicPath : distPath;

  if (fs.existsSync(filePath)) {
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", "attachment; filename=\"Shashank_S_Resume.pdf\"");
    res.sendFile(filePath);
  } else {
    res.status(404).send("Resume PDF file not found");
  }
});

async function startServer() {
  // Ensure workbook initialized on startup
  getOrInitWorkbook();

  // Serve static assets from public folder in all environments
  const publicDir = path.join(process.cwd(), "public");
  if (fs.existsSync(publicDir)) {
    app.use(express.static(publicDir));
  }

  // Vite middleware for dev or static server in prod
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa"
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`\n🚀 Shashank Portfolio Server is running!`);
    console.log(`   > Local:   http://localhost:${PORT}`);
    console.log(`   > Network: http://0.0.0.0:${PORT}\n`);
  });
}

startServer();
