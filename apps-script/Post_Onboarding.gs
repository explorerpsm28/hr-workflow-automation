function processOnboardingEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Column lookups
  const nameCol = headers.indexOf('Name');
  const emailCol = headers.indexOf('Email');
  const dojCol = headers.indexOf('Date of Joining');
  const surveyCol = headers.indexOf('Survey Mail');
  const mgrNameCol = headers.indexOf('Manager Name');
  const mgrEmailCol = headers.indexOf('Manager Email');
  const assessmentCol = headers.indexOf('Assessment Mail');
  const designationCol = headers.indexOf('Designation');
  const departmentCol = headers.indexOf('Department');
  const probationCol = headers.indexOf('Probation Period');
  const probationReminderCol = headers.indexOf('Probation Mail');
  const probationEndCol = headers.indexOf('Probation Completion Mail'); 

  if ([nameCol, emailCol, dojCol, surveyCol, mgrNameCol, mgrEmailCol, assessmentCol, designationCol, departmentCol, probationCol, probationReminderCol, probationEndCol].includes(-1)) {
    Logger.log("One or more required columns are missing.");
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

  // === Single, centered container (600px) for consistent alignment ===
  const WRAP_OPEN =
    '<table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background:#ffffff;">' +
      '<tr><td align="center">' +
        '<table role="presentation" width="600" cellspacing="0" cellpadding="0" border="0" style="width:600px;max-width:600px;">' +
          '<tr><td style="padding:0 20px 0 20px;">';

  const WRAP_CLOSE =          '</td></tr></table></td></tr></table>';

  // === Banner (uses cid:) ===
  // === Banner block + shared container (padded) ===
  
  // Wide banner (e.g., 700px max)
  const BANNER_HTML = `
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
  // === Body text wrapper preserves your line breaks exactly (no words changed) ===
  const TEXT_OPEN  = '<div style="white-space:pre-line;font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.6;color:#111111;">';
  const TEXT_CLOSE = '</div>';

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
            <em>Email:</em> <a href="mailto:sonalirana@nlink.tech" style="text-decoration:none;">hr@nlink.tech</a>
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

  // === Disclaimer (kept verbatim; width now follows the same container) ===
  const DIS = `
  <!-- Disclaimer block: gradient bar + dark panel -->
  <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="width:100%;max-width:600px;margin:0 auto;border-collapse:collapse;">
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
  const normalizeDate = date => new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const todayNorm = normalizeDate(today);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[nameCol];
    const email = row[emailCol];
    const doj = row[dojCol];
    const surveyStatus = row[surveyCol];
    const mgrName = row[mgrNameCol];
    const mgrEmail = row[mgrEmailCol];
    const assessmentStatus = row[assessmentCol];
    const designation = row[designationCol];
    const department = row[departmentCol];
    const probation = row[probationCol];
    const probationFeedbackStatus = row[probationReminderCol];
    const probationEndStatus = row[probationEndCol];

    if (!(doj instanceof Date)) continue;
    const dojNorm = normalizeDate(doj);
    const diffDays = Math.floor((todayNorm - dojNorm) / (1000 * 60 * 60 * 24));

    // 📮 1. Survey Mail (7 days)
    if (diffDays >= 7 && surveyStatus.toString().toLowerCase() !== 'yes') {
      const subject = 'Feedback Request: Your First Week at Nlink Tech';
      let message =
        WRAP_OPEN +
        BANNER_HTML +
        BODY_OPEN +
        TEXT_OPEN +          // wide banner 
        
`Hi ${name},

It's been a week since you joined us, and we're delighted to have you on board! We hope your initial experience with us has been positive and fulfilling.</p><p>To help us continually improve our onboarding process, we kindly request your feedback through a quick survey:

<a href="https://docs.google.com/forms/d/e/1FAIpQLSdy4UVOZoJIxluo8YbGkPdy4WrhBSmcC7-JBVieo-Vnnxix4w/viewform?usp=pp_url">On-Boarding Survey</a>

Please complete the survey by tomorrow. Your input is precious to us.

Feel free to reach out if you have any questions or need assistance.

Best regards,  
Sonali Rana  
Human Resource  
+91 9045293556  
Nlink Tech Pvt Ltd` +
          TEXT_CLOSE +
          '<div style="height:20px;"></div>' +
          TECH_HTM +
          '<div style="height:10px;"></div>' +
          DIS +
          BODY_CLOSE  +
          WRAP_CLOSE ;

      MailApp.sendEmail({
        to: email,
        cc: 'ad@nlink.tech',
        subject: subject,
        htmlBody: message,
        inlineImages: inlineImagesBase
      });
      sheet.getRange(i + 1, surveyCol + 1).setValue('Yes');
    }

    // 📮 2. Assessment Mail (21 days)
    if (diffDays === 21 && assessmentStatus.toString().toLowerCase() !== 'yes') {
      const subject = 'New Joiner Assessment and Feedback';
      let message =
        WRAP_OPEN +
        BANNER_HTML +
        BODY_OPEN +
        TEXT_OPEN +  

`Hi ${mgrName},

I hope you’re doing well! As part of our onboarding process, we conduct an assessment feedback session for new joiners. Your insights will help us understand their progress and development.

Employee Name: ${name}  
DOJ: ${doj.toLocaleDateString()}  
Designation: ${designation}  
Performance Rating (out of 5): ____________________  
Feedback by RM: ____________________  

Your feedback is invaluable in supporting their integration into the team and ensuring their success moving forward.

Thank you for your cooperation!

Best regards,  
Sonali Rana  
Human Resource  
+91 9045293556  
Nlink Tech Pvt Ltd` +
          TEXT_CLOSE +
          '<div style="height:20px;"></div>' +
          TECH_HTM +
          '<div style="height:10px;"></div>' +
          DIS +
          BODY_CLOSE  +
          WRAP_CLOSE ;

      MailApp.sendEmail({
        to: mgrEmail,
        cc: 'ad@nlink.tech',
        subject: subject,
        htmlBody: message,
        inlineImages: inlineImagesBase
      });
      sheet.getRange(i + 1, assessmentCol + 1).setValue('Yes');
    }

    // 📮 3. Probation Reminder (1 month before end)
    if (probation && probation.toString().toLowerCase().includes('month') && probationFeedbackStatus.toString().toLowerCase() !== 'yes') {
      const probationMonths = parseInt(probation.toString().toLowerCase().replace('months', '').trim());
      if (!isNaN(probationMonths)) {
        const confirmationDate = new Date(doj);
        confirmationDate.setMonth(confirmationDate.getMonth() + probationMonths);
        const reminderDate = new Date(confirmationDate);
        reminderDate.setMonth(reminderDate.getMonth() - 1);

        const reminderDateNorm = normalizeDate(reminderDate);

        if (todayNorm.getTime() === reminderDateNorm.getTime()) {
          const subject = "Request for Feedback on Employee's Probation Completion";
          let message =
            WRAP_OPEN +
            BANNER_HTML +
            BODY_OPEN +
            TEXT_OPEN + 
`Dear ${mgrName},

I hope this message finds you well. I am writing to request your feedback on employee's performance as they approach the completion of their probationary period.

The employee has shown exceptional dedication and commitment during these past ${probationMonths - 1} months.

Please provide your valuable feedback in the format below:

Employee Name: ${name}  
Department: ${department}  
Date of Joining: ${doj.toLocaleDateString()}  
Date of Confirmation: ${confirmationDate.toLocaleDateString()}  
Rating (1 to 5): ____________________  
Feedback: ____________________  

Thank you for your attention to this matter. I look forward to your feedback.

Best regards,  
Sonali Rana  
Human Resource  
+91 90452 93556` +
              TEXT_CLOSE +
              '<div style="height:20px;"></div>' +
              TECH_HTM +
              '<div style="height:10px;"></div>' +
              DIS +
              BODY_CLOSE  +
              WRAP_CLOSE ;

          MailApp.sendEmail({
            to: mgrEmail,
            cc: 'ad@nlink.tech',
            subject: subject,
            htmlBody: message,
            inlineImages: inlineImagesBase
          });
          sheet.getRange(i + 1, probationReminderCol + 1).setValue('Yes');
        }
      }
    }

    // 📮 4. Probation Completion Notification (on last day)
    if (probation && probation.toString().toLowerCase().includes('month') && probationEndStatus.toString().toLowerCase() !== 'yes') {
      const probationMonths = parseInt(probation.toString().toLowerCase().replace('months', '').trim());
      if (!isNaN(probationMonths)) {
        const endDate = new Date(doj);
        endDate.setMonth(endDate.getMonth() + probationMonths);
        const endDateNorm = normalizeDate(endDate);

        if (todayNorm.getTime() === endDateNorm.getTime()) {
          const subject = `Completion of Probation Period for ${name}`;
          let message =
            WRAP_OPEN +
            BANNER_HTML +
            BODY_OPEN +
            TEXT_OPEN + 
`Dear ${mgrName},

I hope this message finds you well.

I am writing to inform you that ${name} has successfully completed her ${probation} probation period as of ${endDate.toLocaleDateString()}.

At this stage, please confirm whether we should extend her probation period or proceed with confirming her appointment as a permanent employee.

Your prompt feedback on this matter would be greatly appreciated.

Thank you for your attention to this.

Best regards,  
Sonali Rana  
Human Resource  
+91 90452 93556` +
              TEXT_CLOSE +
              '<div style="height:20px;"></div>' +
              TECH_HTM +
              '<div style="height:10px;"></div>' +
              DIS +
              BODY_CLOSE  +
              WRAP_CLOSE ;

          MailApp.sendEmail({
            to: mgrEmail,
            cc: 'ad@nlink.tech',
            subject: subject,
            htmlBody: message,
            inlineImages: inlineImagesBase
          });
          sheet.getRange(i + 1, probationEndCol + 1).setValue('Yes');
        }
      }
    }
  }
}

