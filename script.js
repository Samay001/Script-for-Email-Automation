function myFunction() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Sheet1");
  
    if (!sheet) { 
      Logger.log("Sheet not found! Check the sheet name.");
      return;
    }
  
    var data = sheet.getDataRange().getValues();
    var fileId = "1UHAt3HJXBzmc1ukQ3q8q1xERJJR3gB2Y"; // Resume File ID
    var file = DriveApp.getFileById(fileId);
    
    for (var i = 1; i < data.length; i++) { 
      var name = data[i][1];
      var email = data[i][2];
      var title = data[i][3];
      var company = data[i][4];
      var status = data[i][5]; // Checking if "Sent" status exists
  
      if (status === "Sent" || !email) {
        continue; // Skip already sent emails or missing email entries
      }
  
      var subject = `Application for Opportunities at ${company}`;
  
      var message = `
  Dear ${name},<br><br>
  I hope this email finds you well at ${company}.<br><br>
  I am a full-stack developer with experience in both frontend and backend technologies, cloud computing, and scalable systems. On the frontend, I have worked with React.js, JavaScript, and Tailwind, ensuring intuitive and responsive user interfaces. On the backend, my expertise includes Node.js, Express, Spring Boot, and database management (MongoDB and MySQL), allowing me to build secure, efficient applications with smooth user experiences.Additionally, I have worked extensively with Kubernetes and Helm, showcasing my expertise in containerization and cloud deployment.<br><br>
  During my previous internship at Boost Star Expert, I was recognized as the top performer for my batch. In this role, I contributed to website development and SEO optimization, improving website performance and user engagement.This experience enhanced my ability to develop scalable, high-performance applications while ensuring seamless user interaction.<br><br>
  Published research paper on Frostbite Detection using ML highlights my understanding of deep learning and computer vision, aligning well with roles requiring AI-based solutions.I have also worked on other projects in similar domains, showcasing my versatility and ability to adapt to various challenges.<br><br>
  My hackathon experience has honed my ability to work under pressure, solve problems efficiently, and deliver high-quality solutions within tight deadlines.With a passion for innovation and optimization, I am eager to contribute my skills to a fast-paced environment where efficiency, scalability, and rapid development are key.<br><br>
  If there are any internship or full-time opportunities available, I would be happy to join and contribute to the team.<br>
  Looking forward to discussing this opportunity further.<br><br>
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
  