// 📩 Induction Mail Automation Script for Nlink Tech
function sendInductionMails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const nameCol = headers.indexOf("Name");
  const emailCol = headers.indexOf("Email");
  const passwrdCol = headers.indexOf("Password")
  const dojCol = headers.indexOf("Date of Joining");
  const welcomeCol = headers.indexOf("Welcome mail Status");
  const inductionCol = headers.indexOf("Induction Mail Status");
  const policyCol = headers.indexOf("HR Policy Mail Status");
  const completionCol = headers.indexOf("Induction Completion Mail Status");

  if ([nameCol,emailCol,dojCol,welcomeCol,inductionCol,policyCol,completionCol].includes(-1)) {
    Logger.log("Missing required columns");
    return;
  }
    // === Inline image setup (Drive file IDs) ===
  // From your original URLs: id=1GLvDYAB8w_K81MvzWQvhpeB7zWrGOKLb and id=1_8T73h0rCtitBRLfgQ4ajgqVEaNiRiGh
  const BANNER_FILE_ID   = '1GLvDYAB8w_K81MvzWQvhpeB7zWrGOKLb';
  const TECH_FILE_ID     = '1_8T73h0rCtitBRLfgQ4ajgqVEaNiRiGh';

  // CID keys used inside HTML
  const BANNER_CID       = 'banner';
  const TECH_LOGO_CID    = 'techlogo';

  // Preload blobs once
  const inlineImagesBase = {
    [BANNER_CID]:   DriveApp.getFileById(BANNER_FILE_ID).getBlob(),
    [TECH_LOGO_CID]:DriveApp.getFileById(TECH_FILE_ID).getBlob()
  };

  // === Banner block (uses cid:) + container ===
  // === Banner block + shared container (padded) ===
  
  // Wide banner (e.g., 700px max)
  const BANNER_HTM = `
  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
    <tr>
      <td align="center">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="max-width:600px;">
          <tr>
            <td style="padding:0 0 16px 0;">
              <img src="cid:${BANNER_CID}" width="600"
                   alt="Company Banner"
                   style="display:block;width:100%;max-width:600px;height:auto;border:0;">
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
`;

  // Narrower body (e.g., 560px max) with side padding
  const BODY_OPEN = `
  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
    <tr>
      <td align="center" style="padding:0 20px;">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="max-width:560px;">
          <tr>
            <td style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:20px;color:#111;">
`;

  const BODY_CLOSE = `
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
`;

  // === Tech signature block (uses cid:) ===
  const TECH_HTM = `
  <div style="margin-top:15px;text-align:left;">
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="width:100%;max-width:520px;">
      <tr>
        <td style="width:86px;padding:6px 12px 6px 0;">
          <img src="cid:${TECH_LOGO_CID}" width="86" alt="Nlink Tech Logo"
               style="display:block;height:auto;border:0;border-radius:12px;">
        </td>
        <td style="vertical-align:top;">
          <div style="font-weight:700; font-size:16px; line-height:20px;">Nlink Tech Pvt Ltd</div>
          <div style="font-size:13px; color:#333; margin:2px 0 8px;">Technology | Innovation | Development</div>

          <div style="font-size:12px; line-height:18px;">
            <em>Email:</em> <a href="mailto:sonalirana@nlink.tech" style="text-decoration:none;">hr@nlink.tech</a></a>
            &nbsp;&nbsp; <em>Web:</em> <a href="https://www.nlink.tech" style="text-decoration:none;">www.nlink.tech</a>
            &nbsp;&nbsp; <em>Tel:</em> +91 9997173164
          </div>

          <div style="font-size:12px; line-height:18px; margin-top:2px;">
            <em>Office:</em> D-81, Kalkaji, New Delhi, 110020
          </div>

          <div style="font-size:12px; line-height:18px; margin-top:8px;">
            <a href="https://www.linkedin.com/company/nlink-tech-pvt-ltd" style="text-decoration:none;">Linkedin</a>
            &nbsp;&nbsp; <a href="https://x.com/NLink_Tech" style="text-decoration:none;">Twitter</a>
            &nbsp;&nbsp; <a href="https://www.youtube.com/@NlinkTech/featured" style="text-decoration:none;">Youtube</a>
          </div>
        </td>
      </tr>
    </table>
  </div>
`;

  
  const DIS = `
  <!-- Disclaimer block: gradient bar + dark panel -->
  <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="width:100%;max-width:700px;margin:0 auto;border-collapse:collapse;">
    <tr>
      <td>
        <!-- Top gradient (Outlook fallback = solid color) -->
        <div style="height:12px;background:linear-gradient(90deg,#e9eefb,#c8ecf3,#d7c6f3);background-color:#e9eefb;"></div>

        <!-- Disclaimer panel -->
        <div style="background-color:#142466;color:#c9cdd6;font-family:Arial,Helvetica,sans-serif;
                    font-size:11px;line-height:1.7;padding:24px 28px;margin:0;">
          This email may contain privileged and confidential information and is intended for the use of the addressee only.
          If you are not an intended recipient, you are requested not to reproduce, copy, disseminate or in any manner distribute
          this email communication as it is strictly prohibited. If you have received this email in error, please notify the
          sender immediately by return email and delete the communication sent in error. Email communications cannot be guaranteed
          to be secure & error-free and Corptive Research Private Limited is not liable for any errors.
        </div>
      </td>
    </tr>
  </table>
`;

  const today = new Date();
  const normalize = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const todayNorm = normalize(today);

  const reportingTime = "12:30 PM";
  const officeAddress = "1 Darshani Gate, Behind Hotel Siddharth Residency, Dehradun, Uttarakhand 248001";
  const signature = 
  `<p>Best regards,<br>
  Sonali Rana<br>
  Human Resource<br>
  +91 9045293556<br></p>`
  ;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[nameCol];
    const email = row[emailCol];
    const doj = row[dojCol];
    const welcomeStatus = row[welcomeCol] ;
    const inductionStatus = row[inductionCol];
    const policyStatus = row[policyCol];
    const completionStatus = row[completionCol];
    const passwrd = row[passwrdCol];

    // Welcome mail 
    if (
      welcomeStatus.toString().toLowerCase() !== 'sent' &&
      doj instanceof Date &&
      normalize(doj).getTime() === todayNorm.getTime()
    ) {
      const subject = "Welcome to Nlink Tech Pvt. Ltd.!";
      let message = 
      BANNER_HTM +            // wide banner
      BODY_OPEN +
      `<p>Dear ${name},</p>

      <p>Congratulations and a warm welcome to Nlink Tech Pvt. Ltd.!</p>

      <p>We are delighted to have you join our team, and we are confident that your skills, experience, and enthusiasm will make a valuable contribution to our organization.</p>

      <p>In the coming days, you will have the opportunity to meet your colleagues and start integrating into our community. We are excited to introduce ourselves and get to know you better.</p>

      <p>To help you get started, we have compiled a list of essential steps and resources:</p>

      <p><b>Email Setup:</b></p>
      <p>Please configure your computer with your work e-mail address using the following credentials in Gmail. Visit <a href="https://mail.google.com">https://mail.google.com</a>.</p>

      <p>Email: ${email}</p>
      <p>Temporary Password: ${passwrd}</p>
      <p><b>Communication:</b> For easy communication, please log in to Skype on your computer. You will also be added to your respective team group for smooth interaction.</p>

      <p>If there is anything else we can do to make your transition as seamless as possible, please let us know.</p>

      <p>Once again, we are thrilled to have you on board, and we hope you are looking forward to your first day as much as we are.</p>`;

      message += signature +
      `<div style="height:20px;"></div>` +  // space between footer and signature
      TECH_HTM +
      `<div style="height:10px;"></div>` +
      DIS  +
      BODY_CLOSE ;

      MailApp.sendEmail({
      to: email,
      cc: 'ad@nlink.tech',
      subject: subject,
      htmlBody: message,
      inlineImages: inlineImagesBase
});

      sheet.getRange(i + 1, welcomeCol + 1).setValue('Sent');
    }
    
    
    // Induction mail
    if (
      inductionStatus.toString().toLowerCase() !== 'sent' &&
      doj instanceof Date &&
      normalize(doj).getTime() === todayNorm.getTime()
    ) {
      const subject = "Welcome to Nlink Tech: Induction Details and Resources";
      let message = 
      BANNER_HTM +            // wide banner
      BODY_OPEN +
      `<p>Dear ${name},</p>
       <p>On behalf of the entire team at Nlink Tech, I extend a warm welcome to you! We are excited to have you join our organization and look forward to the contributions you will bring.</p>
       <p>As you embark on this journey with us, we want to ensure your onboarding experience is seamless and informative.Below, you will find details about our induction program and important resources to help you get started:<p/>

       <h3>Induction Program Details:</h3>

       <p><b>Date:</b> ${doj.toLocaleDateString()}</p>
       <p><b>Time:</b> ${reportingTime}</p>
       <p><b>Location:</b> ${officeAddress}</p>
       
       <p>We are committed to making your transition into Nlink Tech as smooth as possible. Once again, welcome to the team!</p>`
      message += signature +
       `<div style="height:20px;"></div>` +  // space between footer and signature
        TECH_HTM +
      `<div style="height:10px;"></div>`+
        DIS  +
        BODY_CLOSE ;

      MailApp.sendEmail({
      to: email,
      cc: 'ad@nlink.tech',
      subject: subject,
      htmlBody: message,
      inlineImages: inlineImagesBase
});
      sheet.getRange(i + 1, inductionCol + 1).setValue('Sent');
    }

    // HR Policy mail
    if (
      policyStatus.toString().toLowerCase() !== 'sent' &&
      doj instanceof Date &&
      normalize(doj).getTime() === todayNorm.getTime()
    ) {
      const subject = "HR Policy Guidelines";
      let message = 
      BANNER_HTM +            // wide banner
      BODY_OPEN +
      `<p>Dear ${name},</p>
      <p>As discussed during your introduction, I’m sharing the written HR policies for your records and future reference. Please find the attached documents:</p>

      <a href="https://drive.google.com/file/d/1_WyITI3Y729_0mn3aImUKDTQcz-08dc2/view?usp=sharing"><b>HR Policy Guidelines</b></a>

      <p>Kindly review these at your convenience. If you have any questions, feel free to reach out. Please reply to confirm receipt.</p>`
      message += signature +
       `<div style="height:20px;"></div>` +  // space between footer and signature
        TECH_HTM +
      `<div style="height:10px;"></div>`+
        DIS  +
        BODY_CLOSE ;

      MailApp.sendEmail({
      to: email,
      cc: 'ad@nlink.tech',
      subject: subject,
      htmlBody: message,
      inlineImages: inlineImagesBase
});
      sheet.getRange(i + 1, policyCol + 1).setValue('Sent');
    }

// ⏰ Induction Completion Mail (only at 2:00 PM)
    const currentHour = today.getHours();
    const currentMinute = today.getMinutes();
    const is2PM = currentHour === 14 && currentMinute < 5; // 2:00 PM to 2:04 PM window

    if (
     completionStatus.toString().toLowerCase() !== 'sent' &&
     doj instanceof Date &&
     normalize(doj).getTime() === todayNorm.getTime() &&
     is2PM
    ) {
      const subject = "Completion of Your Induction Program";
      let message = 
      BANNER_HTM +            // wide banner
      BODY_OPEN +
      `<p>Dear ${name},</p>
      <p>I am writing to formally acknowledge the successful completion of your induction program at Nlink Tech Pvt. Ltd. Your dedication and engagement throughout this period have been commendable, and we are excited to have you officially on board.</p>

      <p>To ensure that we have properly concluded your induction process and that all necessary aspects have been covered, please acknowledge your induction program's completion by replying to this email.This will help us ensure that everything is in order and that you have received all the necessary information.</p>

      <p>Should you have any questions or require further assistance as you transition into your role, please feel free to contact us.</p>

      <p>Thank you once again for your commitment and enthusiasm. We look forward to your continued success as a part of our team.</p>`
      message += signature +
       `<div style="height:20px;"></div>` +  // space between footer and signature
        TECH_HTM +
      `<div style="height:10px;"></div>`+
        DIS  +
        BODY_CLOSE ;

      MailApp.sendEmail({
      to: email,
      cc: 'ad@nlink.tech',
      subject: subject,
      htmlBody: message,
      inlineImages: inlineImagesBase
});
  sheet.getRange(i + 1, completionCol + 1).setValue('Sent');
}

  }
}
