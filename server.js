const express = require("express");
const multer = require("multer");
const nodemailer = require("nodemailer");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");
const app = express();
const PORT = process.env.PORT || 3000;
// Upload folder setup
const uploadDir = path.join(__dirname, "uploads");
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);

const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, uploadDir),
    filename: (req, file, cb) => cb(null, Date.now() + "-" + file.originalname),
});
const upload = multer({ storage });

// Serve static files
app.use(express.static(path.join(__dirname, "public")));

app.get("/health", (req, res)=>{
    console.log("Backend run properly....**")
})
// ============================================
// SEND EMAILS API
// ============================================
app.post("/send", upload.single("file"), async (req, res) => {
    try {
        const { gmail, appPassword, subject, body, delay } = req.body;

        if (!req.file) {
            return res.json({ success: false, error: "No file uploaded" });
        }

        if (!gmail || !appPassword || !subject || !body) {
            return res.json({ success: false, error: "Missing required fields" });
        }

        // Read uploaded file
        const filePath = req.file.path;
        const workbook = XLSX.readFile(filePath);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet);

        if (!data[0] || !data[0].email) {
            return res.json({ success: false, error: "File must have an 'email' column" });
        }

        console.log(`\n📧 Starting bulk send: ${data.length} emails`);
        console.log(`From: ${gmail}`);
        console.log(`Subject: ${subject}`);

        // Create transporter
        const transporter = nodemailer.createTransport({
            service: "gmail",
            auth: { user: gmail, pass: appPassword },
        });

        // Replace placeholders in subject and body
        function personalize(text, row) {
            let result = text;
            for (const key of Object.keys(row)) {
                const regex = new RegExp(`\\{${key}\\}`, "g");
                result = result.replace(regex, row[key] || "");
            }
            return result;
        }

        // Send emails
        let sent = 0;
        let failed = 0;
        const delayMs = (parseInt(delay) || 2) * 1000;

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const personalizedSubject = personalize(subject, row);
            const personalizedBody = personalize(body, row).replace(/\n/g, "<br>");

            try {
                await transporter.sendMail({
                    from: `"${gmail.split("@")[0]}" <${gmail}>`,
                    to: row.email,
                    subject: personalizedSubject,
                    html: `<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;padding:20px;">
                        ${personalizedBody}
                    </div>`,
                    replyTo: gmail,
                });
                console.log(`✅ [${i + 1}/${data.length}] Sent to ${row.email}`);
                sent++;
            } catch (err) {
                console.log(`❌ [${i + 1}/${data.length}] Failed: ${row.email} - ${err.message}`);
                failed++;
            }

            // Delay between emails
            if (i < data.length - 1) {
                await new Promise((resolve) => setTimeout(resolve, delayMs));
            }
        }

        // Save results
        const results = data.map((row, i) => ({
            email: row.email,
            status: i < sent ? "sent" : "failed",
        }));
        const resultSheet = XLSX.utils.json_to_sheet(results);
        const resultBook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(resultBook, resultSheet, "Results");
        const resultPath = path.join(__dirname, "send-results.xlsx");
        XLSX.writeFile(resultBook, resultPath);

        // Clean up uploaded file
        fs.unlinkSync(filePath);

        console.log(`\n📊 Done! Sent: ${sent}, Failed: ${failed}`);

        res.json({
            success: true,
            total: data.length,
            sent,
            failed,
            resultsFile: "send-results.xlsx",
        });
    } catch (error) {
        console.error("Server error:", error);
        res.json({ success: false, error: error.message });
    }
});

// Download results
app.get("/download-results", (req, res) => {
    const filePath = path.join(__dirname, "send-results.xlsx");
    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.status(404).json({ error: "No results file found" });
    }
});

// Start server
app.listen(PORT, () => {
    console.log("\n========================================");
    console.log("   BULK EMAIL SENDER - SERVER");
    console.log("========================================");
    console.log(`\n🚀 Server running at: http://localhost:${PORT}`);
    console.log("\nSteps:");
    console.log("1. Open browser: http://localhost:3000");
    console.log("2. Enter Gmail + App Password");
    console.log("3. Write email subject & body");
    console.log("4. Upload CSV/Excel file");
    console.log("5. Click Send Emails");
    console.log("\n========================================\n");
});
