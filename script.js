function myFunction() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Sheet1");

  if (!sheet) { 
    Logger.log("Sheet not found! Check the sheet name.");
    return;
  }

  var data = sheet.getDataRange().getValues();
  var fileId = "1PCrbEDrL948jjKcEaELzsquvBEQhlN-7"; // Resume File ID
  var file = DriveApp.getFileById(fileId);
  
  for (var i = 0; i < data.length; i++) { 
    var name = data[i][1];
    var email = data[i][2];
    var title = data[i][3];
    var company = data[i][4];
    var status = data[i][6]; // Checking if "Sent" status exists

    if (status === "Sent" || !email) {
      continue; // Skip already sent emails or missing email entries
    }

    var subject = `Application for Opportunities at ${company}`;

    var message = `
Dear ${name},<br>
I hope this email finds you well at ${company}.<br><br>
My name is Samay Rathod, a passionate Full-Stack Developer actively seeking exciting opportunities where I can contribute my skills in building scalable, high-performance applications.<br>
<b>Why I Am a Stronger Candidate Than Others:-</b><br>
<b>Full-Stack Expertise:</b> Proficient in both frontend <b>(React.js, JavaScript, Tailwind)</b> and backend <b>(Node.js, Express, Spring Boot, MongoDB, MySQL)</b>, enabling me to build seamless, responsive, and scalable applications.<br>
<b>Cloud & DevOps Knowledge:</b> Hands-on experience with <b>Kubernetes and Docker</b>, showcasing my ability to handle containerization and cloud deployment efficiently.<br>
<b>Proven Track Record:</b> Recognized as the <b>top-performing intern at Boost Star Expert</b>, where I contributed to <b>website development and SEO optimization</b>, improving website performance and engagement.<br>
<b>AI & ML Research:</b>Published a <b>research paper</b> on Frostbite Detection using ML, demonstrating my ability to work on AI-based solutions and understand deep learning and computer vision.<br>
<b>Hackathon Experience:</b> Proven ability to work under pressure, solve complex problems, and deliver high-quality solutions within tight deadlines, making me adaptable and efficient.<br>
<b>Immediate Joiner:</b> Available for internship or full-time roles, ready to contribute immediately and drive impact from day one.<br><br>
Best regards,<br>
Samay
    `;

    try {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: message,
        attachments: [file.getAs(MimeType.PDF)]
      });

      sheet.getRange(i + 1, 6).setValue("Sent"); // Mark as sent
      Logger.log(`Email sent to: ${email}`);
      
    } catch (e) {
      Logger.log(`Error sending email to ${email}: ${e.message}`);
    }
  }
}
