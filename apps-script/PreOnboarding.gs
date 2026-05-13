// 🔁 Automated Onboarding Email Script for Nlink Tech
function sendFormalityMails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 🔍 Column validation with logging
  const requiredHeaders = [
    'Name', 'Email', 'Date of Joining', 'Date of Offer Acceptance',
    'Formality Mail Status', 'Declaration Mail Status', 'ID Creation Mail Status',
    'Quick Update Mail Status', 'Background Check Required', 'Background Mail Status',
    'Email of Previous Employeer', 'Previous Employeer Mail Status',
    'Designation', 'Department'
    // LOI columns removed
  ];

  const headerMap = {};
  requiredHeaders.forEach(header => {
    const index = headers.indexOf(header);
    if (index === -1) {
      Logger.log(`❌ Missing column: "${header}"`);
    } else {
      headerMap[header] = index;
    }
  });

  // If any column is missing, stop execution
  if (Object.keys(headerMap).length !== requiredHeaders.length) {
    throw new Error('⛔ One or more required columns are missing. Check logs.');
  }

  const nameCol = headers.indexOf('Name');
  const emailCol = headers.indexOf('Email');
  const dojCol = headers.indexOf('Date of Joining');
  const offerDateCol = headers.indexOf('Date of Offer Acceptance');
  const formalityCol = headers.indexOf('Formality Mail Status');
  const declarationCol = headers.indexOf('Declaration Mail Status');
  const idCreationCol = headers.indexOf('ID Creation Mail Status');
  const quickUpdateCol = headers.indexOf('Quick Update Mail Status');
  const backgroundReqCol = headers.indexOf('Background Check Required');
  const backgroundStatusCol = headers.indexOf('Background Mail Status');
  const prevEmpEmailCol = headers.indexOf('Email of Previous Employeer');
  const prevEmpStatusCol = headers.indexOf('Previous Employeer Mail Status');
  const designationCol = headers.indexOf('Designation');
  const departmentCol = headers.indexOf('Department');

  if ([nameCol, emailCol, dojCol, offerDateCol, formalityCol, declarationCol, idCreationCol, quickUpdateCol, backgroundReqCol, backgroundStatusCol, prevEmpEmailCol, prevEmpStatusCol, designationCol, departmentCol].includes(-1)) {
    Logger.log('One or more required columns are missing.');
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

  const signature = 
  `<p>Best regards,<br>
  Sonali Rana<br>
  Human Resource<br>
  +91 9045293556<br></p>`
  ;
  const reportingTime = "12:30 PM";
  const officeAddress = "1 Darshani Gate, Behind Hotel Siddharth Residency, Dehradun, Uttarakhand 248001";
  const contactPerson = "Sonali Rana";
  const contactDetails = "+91 90452 93556";

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[nameCol];
    const email = row[emailCol];
    const doj = row[dojCol];
    const offerDate = row[offerDateCol];
    const formalityStatus = row[formalityCol];
    const declarationStatus = row[declarationCol];
    const idCreationStatus = row[idCreationCol];
    const quickUpdateStatus = row[quickUpdateCol];
    const bgRequired = row[backgroundReqCol];
    const bgStatus = row[backgroundStatusCol];
    const prevEmpEmail = row[prevEmpEmailCol];
    const prevEmpStatus = row[prevEmpStatusCol];
    const designation = row[designationCol];
    const department = row[departmentCol];

    // 📩 Formality Mail
    if (
      formalityStatus.toString().toLowerCase() !== 'sent' &&
      offerDate instanceof Date &&
      normalize(offerDate).getTime() === todayNorm.getTime()
    ) {
      const subject = 'Joining Formalities with Nlink Tech';
      let message = 
      BANNER_HTM +            // wide banner
      BODY_OPEN +
     `<p>Dear ${name},</p>
      <p>Hope you are doing well! We are thrilled to welcome you to Nlink Tech!</p> <p> As part of your pre-onboarding process, please upload the required documents using the link below: Document Upload: <a href="https://docs.google.com/forms/d/e/1FAIpQLSeYYwT6LxTMiEAwRbgmiFPeOr6rB8H48HqiC4YP1ZHFvpJ5jA/viewform"><em>Form</em></a></p><p>Please complete this at the earliest so we can proceed with your onboarding smoothly.</p>
      <p>If you have any questions or need help filling out the form, feel free to reach out.Looking forward to having you on board!\nOnce again, congratulations!</p>`
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
      sheet.getRange(i + 1, formalityCol + 1).setValue('Sent');
    }

    // 📄 Declaration Mail
    if (
      declarationStatus.toString().toLowerCase() !== 'sent' &&
      offerDate instanceof Date &&
      normalize(new Date(offerDate.getTime() + 86400000)).getTime() === todayNorm.getTime()
    ) {
      const subject = 'Declaration Acknowledgment';
      let message = 
      BANNER_HTM +            // wide banner
      BODY_OPEN +
      `<p>Hi ${name},</p>
      <p>Please review the attached NDA/IT Security/Code of Conduct documents.</p>
      <p>Kindly provide a self-attested copy or acknowledge acceptance via email.</p>
      <p>When acknowledging, confirm agreement to the terms of the Code of Conduct, IT Security Policy, and NDA shared here: <a href="https://drive.google.com/file/d/1R4k8HxnxCps54BsNfPIjJYSru3D11_En/view?usp=sharing">Document</a>
      </p>`
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
      sheet.getRange(i + 1, declarationCol + 1).setValue('Sent');
    }

    // ✅ Background Check Mail
    if (bgRequired.toString().toLowerCase() === 'yes' && bgStatus.toString().toLowerCase() !== 'sent') {
      const subject = 'Request for Reference Details';
      let message = 
      BANNER_HTM +            // wide banner
      BODY_OPEN +
     `<p>Dear ${name},</p>
     <p>As part of the hiring process for ${designation}, we kindly request references from your previous employers or professional contacts. Required details:</br></p>
     <p><br>• Reference Name</br>
     <br>• Company/Organization</br>
     <br>• Position/Title</br>
     <br>• Email</br>
     <br>• Phone</br>
     <br>• Relationship</br></p>
    <p>Your cooperation is appreciated.<p/>`
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
      sheet.getRange(i + 1, backgroundStatusCol + 1).setValue('Sent');
    }

    // 🏢 Previous Employer Mail
    if (prevEmpEmail && prevEmpStatus.toString().toLowerCase() !== 'sent') {
      const subject = `Background Verification of ${name}`;
      let message = 
      BANNER_HTM +            // wide banner
      BODY_OPEN +
     `<p><br>Dear Sir/Ma’am,</br>
      <p>Greetings</p>
      I am Sonali, and I am working with Nlink Tech Pvt. Ltd. You must be aware that during the process of recruitment, employee verification is an important step, and therefore, I am writing to you to seek your assistance in verifying the details given below of the following candidate:
      <p><br>Name of the Candidate:</br> 
      <br>Company’s Name: Nlink Tech Pvt Ltd</br>
      <br>Position/ Profile:</br></p>
      Kindly provide the following details about the candidate mentioned above.
     <p><br>• Designation</br>
     <br>• Department</br>
     <br>• Date of Joining</br>
     <br>• Date of Leaving</br>
     <br>• Salary Gross/ CTC</br>
     <br>• Reason for Leaving</br>
     <br>• Eligible to hire (Yes or No)</br>
     <br>-Name of the Verifier:</br>
     <br>-Position of the Verifier:</br></p>

     <p>We look forward to your responses and would appreciate it if you could fill in the details as soon as possible. Please get in touch with us at +91 90452 93556  or revert us on the same email.</p>

     We are thankful for your time and help.</p>`
      message += signature +
      `<div style="height:20px;"></div>` +  // space between footer and signature
        TECH_HTM +
      `<div style="height:10px;"></div>`+
        DIS  +
        BODY_CLOSE ;

      MailApp.sendEmail({
      to: prevEmpEmail,
      cc: 'ad@nlink.tech',
      subject: subject,
      htmlBody: message,
      inlineImages: inlineImagesBase
});
      sheet.getRange(i + 1, prevEmpStatusCol + 1).setValue('Sent');
    }

    // 🛠️ ID Creation Mail
    if (
      idCreationStatus.toString().toLowerCase() !== 'sent' &&
      doj instanceof Date &&
      normalize(new Date(doj.getTime() - 86400000)).getTime() === todayNorm.getTime()
    ) {
      const subject = 'Joining Information for Email ID Creation';
      let message = 
      BANNER_HTM +            // wide banner
      BODY_OPEN +
      `<p>Hi Mam,</p>
      <p>Please create email IDs for the following new joiner:</p> 
      <br>• Name: ${name}</br>
      <br>• DOJ: ${doj.toLocaleDateString()}</br> 
      <br>• Designation: ${designation}</br>
      <br>• Department: ${department}</br> 
      <br>• Email: ${email}</br> 
      <p>Proceed with the creation of their email IDs and add the same ID in the Group mail.</p>
      <p>Thank you for your prompt attention to this matter.</p>`
      message += signature +
       `<div style="height:20px;"></div>` +  // space between footer and signature
        TECH_HTM +
      `<div style="height:10px;"></div>`+
        DIS  +
        BODY_CLOSE ;

      MailApp.sendEmail({
      to: 's2@corptive.com',
      cc: 'ad@nlink.tech',
      subject: subject,
      htmlBody: message,
      inlineImages: inlineImagesBase
});
      sheet.getRange(i + 1, idCreationCol + 1).setValue('Sent');
    }

    // 🕒 Quick Update Mail
    if (
      quickUpdateStatus.toString().toLowerCase() !== 'sent' &&
      doj instanceof Date &&
      normalize(new Date(doj.getTime() - 86400000)).getTime() === todayNorm.getTime()
    ) {
      const subject = 'Quick Update: Joining Details';
      let message = 
      BANNER_HTM +            // wide banner
      BODY_OPEN +
      `<p>Hi ${name},</p>
      <p>We are delighted to welcome you to Nlink Tech Pvt Ltd! We are looking forward to your first day with us and are confident that you will be a valuable addition to our team.</p>
      <p>Here are the details you will need for your first day:</p>    
      <p><b>Date of Joining (DOJ):</b> ${doj.toLocaleDateString()}</p>
      <p><b>Reporting Time:</b> ${reportingTime}</p>
      <p><b>Contact Person:</b> ${contactPerson}</p>
      <p><b>Contact Number:</b> ${contactDetails}</p>
      <p><b>Office Address:</b> ${officeAddress}</p>
      
      <p>Welcome aboard!</p>`
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
      sheet.getRange(i + 1, quickUpdateCol + 1).setValue('Sent');
    }

    // ✅ LOI logic removed
  }
}
