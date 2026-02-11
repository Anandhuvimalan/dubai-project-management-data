const puppeteer = require("puppeteer");
const fs = require("fs");

const mdPath = "/home/anandhuvimalan/.gemini/antigravity/brain/5f7fdecb-31e8-4138-90bd-ad13eca76b3b/dashboard_blueprint.md";
const htmlPath = "/home/anandhuvimalan/dubai_data/dubai-project-management-data/dashboard_blueprint.html";
const pdfPath = "/home/anandhuvimalan/dubai_data/dubai-project-management-data/dashboard_blueprint.pdf";

(async () => {
  try {
    // Dynamic import for ESM module 'marked'
    const { marked } = await import('marked');

    const mdContent = fs.readFileSync(mdPath, "utf8");
    const htmlContent = marked.parse(mdContent);

    // Clean, professional CSS for text-heavy specification
    const fullHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
          
          body { 
            font-family: 'Inter', 'Helvetica', 'Arial', sans-serif; 
            padding: 40px; 
            line-height: 1.6; 
            color: #1e293b; 
            max-width: 900px; 
            margin: 0 auto; 
            -webkit-print-color-adjust: exact; 
          }
          h1, h2, h3, h4 { color: #0f172a; margin-top: 1.5em; font-weight: 700; }
          h1 { border-bottom: 2px solid #e2e8f0; padding-bottom: 15px; font-size: 28px; letter-spacing: -0.02em; color: #4f46e5; }
          h2 { font-size: 22px; border-bottom: 1px solid #e2e8f0; padding-bottom: 8px; margin-top: 30px; }
          h3 { font-size: 18px; color: #334155; margin-top: 25px; }
          p { margin-bottom: 15px; color: #475569; }
          
          table { width: 100%; border-collapse: collapse; margin: 25px 0; font-size: 14px; }
          th, td { padding: 10px 12px; border: 1px solid #e2e8f0; text-align: left; }
          th { background-color: #f8fafc; font-weight: 600; color: #0f172a; }
          tr:nth-child(even) { background-color: #f8fafc; }
          
          ul, ol { margin-left: 20px; color: #475569; }
          li { margin-bottom: 8px; }
          
          code { background-color: #f1f5f9; padding: 2px 5px; border-radius: 4px; font-family: 'Menlo', monospace; font-size: 0.9em; color: #64748b; }
          
          /* Mermaid diagram placeholder style */
          .mermaid { display: none; } 
          
          @media print {
            body { padding: 0; max-width: 100%; }
            .page-break { page-break-after: always; }
          }
        </style>
      </head>
      <body>
        ${htmlContent}
      </body>
      </html>
    `;

    fs.writeFileSync(htmlPath, fullHtml);
    console.log("Created HTML:", htmlPath);

    const browser = await puppeteer.launch({
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    const page = await browser.newPage();
    await page.setContent(fullHtml, { waitUntil: 'networkidle0' });
    await page.pdf({
      path: pdfPath,
      format: 'A4',
      printBackground: true,
      margin: { top: '2cm', right: '2cm', bottom: '2cm', left: '2cm' }
    });

    await browser.close();
    console.log("Created PDF:", pdfPath);
  } catch (err) {
    console.error("Error generating PDF:", err);
    process.exit(1);
  }
})();
