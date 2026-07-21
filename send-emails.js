const nodemailer = require("nodemailer");
const XLSX = require("xlsx");
const path = require("path");

// ============================================
// CONFIGURATION - Yahan apni details daalo
// ============================================
const CONFIG = {
    gmail: "your@gmail.com",           // Apna Gmail daalo
    appPassword: "xxxx xxxx xxxx xxxx", // Gmail App Password daalo
    subject: "Hello from UniSpark",     // Email subject
    delay: 2000,                        // Delay between emails (ms)
};

// ============================================
// HTML EMAIL TEMPLATE
// ============================================
function getEmailHTML(name, company) {
    return `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #1F3C88;">Hi ${name},</h2>
        <p>We are reaching out to <strong>${company}</strong> regarding a potential partnership opportunity.</p>
        <p>Our team specializes in providing end-to-end solutions that can help your business grow.</p>
        <br>
        <p>Best regards,</p>
        <p><strong>Your Name</strong><br>
        Your Designation<br>
        Your Company<br>
        Phone: +971 XX XXX XXXX</p>
    </div>
    `;
}

// ============================================
// TRANSPORTER SETUP
// ============================================
const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
        user: CONFIG.gmail,
        pass: CONFIG.appPassword,
    },
});

// ============================================
// READ EXCEL/CSV FILE
// ============================================
function readFile(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);
    console.log(`\n📁 File loaded: ${data.length} emails found`);
    console.log("Columns:", Object.keys(data[0]).join(", "));
    return data;
}

// ============================================
// SEND SINGLE EMAIL
// ============================================
async function sendOneEmail(user) {
    const mailOptions = {
        from: `"Your Name" <${CONFIG.gmail}>`,
        to: user.email,
        subject: CONFIG.subject.replace("{name}", user.name || ""),
        html: getEmailHTML(user.name || "there", user.company || "your organization"),
        replyTo: CONFIG.gmail,
    };

    try {
        const info = await transporter.sendMail(mailOptions);
        return { success: true, email: user.email, messageId: info.messageId };
    } catch (error) {
        return { success: false, email: user.email, error: error.message };
    }
}

// ============================================
// MAIN FUNCTION
// ============================================
async function main() {
    const filePath = process.argv[2] || path.join(__dirname, "emails.xlsx");

    console.log("========================================");
    console.log("   BULK EMAIL SENDER");
    console.log("========================================");
    console.log(`From: ${CONFIG.gmail}`);
    console.log(`File: ${filePath}`);

    const data = readFile(filePath);

    // Validate data
    if (!data[0].email) {
        console.error("\n❌ Error: File mein 'email' column nahi mila!");
        console.log("Required columns: email, name (optional), company (optional)");
        process.exit(1);
    }

    let sent = 0;
    let failed = 0;
    const results = [];

    console.log("\n📧 Sending emails...\n");

    for (let i = 0; i < data.length; i++) {
        const user = data[i];
        const progress = `[${i + 1}/${data.length}]`;

        process.stdout.write(`${progress} Sending to ${user.email}... `);

        const result = await sendOneEmail(user);
        results.push(result);

        if (result.success) {
            console.log("✅ Sent");
            sent++;
        } else {
            console.log(`❌ Failed: ${result.error}`);
            failed++;
        }

        // Delay between emails
        if (i < data.length - 1) {
            await new Promise((resolve) => setTimeout(resolve, CONFIG.delay));
        }
    }

    // Summary
    console.log("\n========================================");
    console.log("   SUMMARY");
    console.log("========================================");
    console.log(`Total:   ${data.length}`);
    console.log(`Sent:    ${sent} ✅`);
    console.log(`Failed:  ${failed} ❌`);
    console.log("========================================");

    // Save results to file
    const resultsWorkbook = XLSX.utils.json_to_sheet(results);
    const resultsBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(resultsBook, resultsWorkbook, "Results");
    XLSX.writeFile(resultsBook, path.join(__dirname, "send-results.xlsx"));
    console.log("\n📊 Results saved to: send-results.xlsx");
}

main().catch(console.error);
