function runLOIAutomation() {
  /***** CONFIG *****/
  const sheetName = 'Sheet1'; // <-- change if your tab is named differently
  const approvalRecipients = ['s2@corptive.com', 'sonalirana@nlink.tech'];
  const ccOnFinal = 'ad@nlink.tech'; // <-- CC for final candidate mail
  let DRY_RUN = false; // set to false to actually send emails

  // === File IDs for the DOCX templates & (optional) output folder ===
  // Upload both DOCX files to Drive and paste their FILE IDs here:
  const TEMPLATE_DOCX_EMP_FILE_ID    = '105Q-FUDB4THZbVGpZliQfVPAfCV5aLIh5a6EuX3oCSU';     // e.g., "LOI- Employee Format.docx"
  const TEMPLATE_DOCX_INTERN_FILE_ID = '1LWEpGO0VbGtBVNDSxG1piPKpcFSd4amj2rKXvhXCF1c';   // e.g., "LOI- Internship.docx"
  // (Optional) Put generated Google Docs + PDFs in this folder id; leave '' to use root
  const OUTPUT_FOLDER_ID = '';

  // NOTE: Requires Advanced Drive Service enabled (Services → Drive API v2).

  /***** UTILITIES *****/
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName) || ss.getActiveSheet();
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Sheet is empty.');

  const headers = values[0].map(h => String(h).trim());
  const norm = s => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
  const headerIndex = (target, alts=[]) => {
    const wanted = [target, ...alts].map(norm);
    for (let i = 0; i < headers.length; i++) {
      const h = norm(headers[i]);
      if (wanted.includes(h)) return i;
    }
    return -1;
  };

  const col = {
  name: headerIndex('Name'),
  designation: headerIndex('Designation'),
  dept: headerIndex('Department'),
  email: headerIndex('Email'),
  doj: headerIndex('Date of Joining', ['DOJ']),
  salary: headerIndex('Salary', ['Stipend','CTC','Cost To Company']),

  // 👇 NEW: "Valid Till" is for LOI validity
  validTill: headerIndex('Valid Till', [
    'Offer Validity','LOI Valid Till','Offer Validity Date','Validity'
  ]),

  // 👇 NEW: separate column for internship / contract end date
  endDate: headerIndex('End Date', [
    'Internship End','End Duration','Contract End'
  ]),

  letterDate: headerIndex('Letter Date', ['Offer Date','LOI Date']),
  approvalMailStatus: headerIndex('Approval Mail Status', ['LOI Approval Mail Status']),
  approvalOfLOI: headerIndex('Approval of LOI', ['Aprroval of LOI']),
  internStatus: headerIndex('LOI Intern Status'),
  employeeStatus: headerIndex('LOI Employee Status'),
  backgroundRequired: headerIndex('Background Check Required', ['Background required'])
  };

  // === Inline image setup (Drive file IDs) ===
  const BANNER_FILE_ID   = '1GLvDYAB8w_K81MvzWQvhpeB7zWrGOKLb';
  const TECH_FILE_ID     = '1_8T73h0rCtitBRLfgQ4ajgqVEaNiRiGh';
  const BANNER_CID       = 'banner';
  const TECH_LOGO_CID    = 'techlogo';
  const inlineImagesBase = {
    [BANNER_CID]:    DriveApp.getFileById(BANNER_FILE_ID).getBlob(),
    [TECH_LOGO_CID]: DriveApp.getFileById(TECH_FILE_ID).getBlob()
  };

  // === Shared HTML blocks ===
  const BANNER_HTML = `
  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
    <tr>
      <td align="center">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="max-width:600px;">
          <tr>
            <td style="padding:0 0 16px 0;">
              <img src="cid:${BANNER_CID}" width="600" alt="Company Banner" style="display:block;width:100%;max-width:600px;height:auto;border:0;">
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>`;

  const BODY_OPEN = `
  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
    <tr>
      <td align="center" style="padding:0 20px;">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="max-width:560px;">
          <tr>
            <td style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:20px;color:#111;">`;

  const BODY_CLOSE = `
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>`;

  const TECH_HTML = `
  <div style="margin-top:15px;text-align:left;">
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="width:100%;max-width:520px;">
      <tr>
        <td style="width:86px;padding:6px 12px 6px 0;">
          <img src="cid:${TECH_LOGO_CID}" width="86" alt="Nlink Tech Logo" style="display:block;height:auto;border:0;border-radius:12px;">
        </td>
        <td style="vertical-align:top;">
          <div style="font-weight:700; font-size:16px; line-height:20px;">Nlink Tech Pvt Ltd</div>
          <div style="font-size:13px; color:#333; margin:2px 0 8px;">Technology | Innovation | Development</div>
          <div style="font-size:12px; line-height:18px;">
            <em>Email:</em> <a href="mailto:hr@nlink.tech" style="text-decoration:none;">hr@nlink.tech</a>
            &nbsp;&nbsp; <em>Web:</em> <a href="https://www.nlink.tech" style="text-decoration:none;">www.nlink.tech</a>
            &nbsp;&nbsp; <em>Tel:</em> +91 9997173164
          </div>
          <div style="font-size:12px; line-height:18px; margin-top:2px;">
            <em>Office:</em> D-81, Kalkaji, New Delhi, 110020
          </div>
          <div style="font-size:12px; line-height:18px; margin-top:8px;">
            <a href="https://www.linkedin.com/company/nlink-tech-pvt-ltd" style="text-decoration:none;">LinkedIn</a>
            &nbsp;&nbsp; <a href="https://x.com/NLink_Tech" style="text-decoration:none;">Twitter</a>
            &nbsp;&nbsp; <a href="https://www.youtube.com/@NlinkTech/featured" style="text-decoration:none;">YouTube</a>
          </div>
        </td>
      </tr>
    </table>
  </div>`;

  const DISC = `
  <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="width:100%;max-width:700px;margin:0 auto;border-collapse:collapse;">
    <tr>
      <td>
        <div style="height:12px;background:linear-gradient(90deg,#e9eefb,#c8ecf3,#d7c6f3);background-color:#e9eefb;"></div>
        <div style="background-color:#142466;color:#c9cdd6;font-family:Arial,Helvetica,sans-serif;font-size:11px;line-height:1.7;padding:24px 28px;margin:0;">
          This email may contain privileged and confidential information and is intended for the use of the addressee only. If you are not an intended recipient, you are requested not to reproduce, copy, disseminate or in any manner distribute this email communication as it is strictly prohibited. If you have received this email in error, please notify the sender immediately by return email and delete the communication sent in error. Email communications cannot be guaranteed to be secure & error-free and Corptive Research Private Limited is not liable for any errors.
        </div>
      </td>
    </tr>
  </table>`;

  // Log missing columns (non-fatal)
  Object.entries(col).forEach(([k, v]) => { if (v < 0) Logger.log(`Missing column: ${k} (check header spelling)`); });

  const tz = Session.getScriptTimeZone();
  const MS_PER_DAY = 24 * 60 * 60 * 1000;
  const fmtDate = d => d ? Utilities.formatDate(d, tz, 'dd-MMM-yyyy') : '';
  const fmtDateTime = d => Utilities.formatDate(d, tz, 'dd-MMM-yyyy HH:mm');
  const toDate = v => {
    if (v === null || v === undefined || v === '') return null;
    if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return v;
    const p = new Date(v); return isNaN(p) ? null : p;
  };
  const toMidnight = d => { const x = new Date(d); x.setHours(0,0,0,0); return x; };
  const text = v => (v === null || v === undefined) ? '' : String(v).trim();
  const looksSent = v => /^\s*sent\b/i.test(text(v));
  const isYes = v => /^yes$/i.test(text(v));
  const INTERN_REGEX = /\b(intern|internship|industrial\s*trainee|trainee|apprentice|apprenticeship|ojt)\b/i;
  const isInternish = row => {
    const dsg = col.designation >= 0 ? text(row[col.designation]) : '';
    const dept = col.dept >= 0 ? text(row[col.dept]) : '';
    return INTERN_REGEX.test(dsg) || INTERN_REGEX.test(dept);
  };

  // ✅ NEW: LOI will send only if these columns are filled for that row:
  // Name, Designation, Department, Email, Date of Joining, Valid Till, Salary
  const isRowCompleteForLOI = (row, rowNum) => {
    const missing = [];

    const isFilled = (idx) => idx >= 0 && text(row[idx]) !== '';
    const isDateFilled = (idx) => idx >= 0 && toDate(row[idx]) !== null;

    if (!isFilled(col.name)) missing.push('Name');
    if (!isFilled(col.designation)) missing.push('Designation');
    if (!isFilled(col.dept)) missing.push('Department');
    if (!isFilled(col.email)) missing.push('Email');
    if (!isDateFilled(col.doj)) missing.push('Date of Joining');
    if (!isDateFilled(col.validTill)) missing.push('Valid Till');
    if (col.endDate >= 0 && !isDateFilled(col.endDate)) missing.push('End Date');
    if (!isFilled(col.salary)) missing.push('Salary');

    if (missing.length) {
      Logger.log(`Row ${rowNum}: Skipping LOI (required fields missing): ${missing.join(', ')}`);
      return false;
    }
    return true;
  };

  // === Shared date helpers for both PDF builders ===
  const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const pad2 = n => (n < 10 ? '0' : '') + n;
  const daySuffix = n => (n % 100 >= 11 && n % 100 <= 13) ? 'th' : (['th','st','nd','rd'][Math.min(n % 10, 4)] || 'th');
  const fmtDateLong = d => { if (!d) return ''; const day = d.getDate(); return `${pad2(day)}${daySuffix(day)} ${months[d.getMonth()]} ${d.getFullYear()}`; };
  const minusOneDay = d => { if (!d) return null; const x = new Date(d); x.setDate(x.getDate() - 1); return x; };

  const stripHtml = html => html.replace(/<[^>]*>/g, ' ');
  const sendEmail = (to, subject, htmlBody, cc, attachments) => {
    if (DRY_RUN) {
      Logger.log(`[DRY RUN] To: ${to} | CC: ${cc || '-'} | Subject: ${subject}`);
      if (attachments && attachments.length) Logger.log(`[DRY RUN] Attachments: ${attachments.map(b => b.getName()).join(', ')}`);
      return;
    }
    const opts = { htmlBody: htmlBody, inlineImages: inlineImagesBase };
    if (attachments && attachments.length) opts.attachments = attachments;
    if (cc) opts.cc = cc;
    GmailApp.sendEmail(to, subject, stripHtml(htmlBody), opts);
  };

  const buildApprovalMail = (row) => {
    const intern = isInternish(row);

    // Use existing candidate builders to get the full HTML body
    const base = intern ? buildInternLOI(row) : buildEmployeeLOI(row);

    const subject = intern
      ? 'LOI (Intern) - preview for approval'
      : 'LOI (Employee) - preview for approval';

    return { subject, htmlBody: base.htmlBody };
  };

  /***** Email bodies *****/
  const buildInternLOI = (row) => {
    const name = col.name >= 0 ? text(row[col.name]) : '';
    const designation = (col.designation >= 0 ? text(row[col.designation]) : '') || 'Intern';
    const subject = "Congratulations, your LOI is attached along";
    const htmlBody = BANNER_HTML + BODY_OPEN +
`<p>Dear ${name},</p>
<p>We are pleased to offer you the position of ${designation} at Nlink Tech Pvt Ltd. You will be directly working with the Nlink Tech team. Please accept this letter and send us back the signed copy.</p>
<p>Also, please send us the required documents at the earliest.</p>
<p>Best regards,<br>Sonali Rana<br>Human Resource<br>+91 90452 93556</p>` +
`<div style="height:20px;"></div>` + TECH_HTML + `<div style="height:10px;"></div>` + DISC + BODY_CLOSE;
    return { subject, htmlBody };
  };

  const buildEmployeeLOI = (row) => {
    const name = col.name >= 0 ? text(row[col.name]) : '';
    const designation = col.designation >= 0 ? text(row[col.designation]) : '';
    const subject = "Congratulations, your LOI is attached along";
    const htmlBody = BANNER_HTML + BODY_OPEN +
`<p>Dear ${name},</p>
<p>We are pleased to offer you the position of ${designation} at Nlink Tech Pvt Ltd. You will be directly working with the Nlink Tech team.</p>
<p>Please accept this letter and send us back the signed copy. Also, please send us the required documents at the earliest.</p>
<p>Your Appointment Letter would be provided after your probation period. We look forward to your acceptance and working with you soon.</p>
<p>Best regards,<br>Sonali Rana<br>Human Resource<br>+91 90452 93556</p>` +
`<div style="height:20px;"></div>` + TECH_HTML + `<div style="height:10px;"></div>` + DISC + BODY_CLOSE;
    return { subject, htmlBody };
  };

  /***** PDF builders *****/
  function makeEmployeePdfFromTemplate(row) {
    if (!TEMPLATE_DOCX_EMP_FILE_ID || TEMPLATE_DOCX_EMP_FILE_ID === 'PASTE_EMPLOYEE_DOCX_FILE_ID_HERE') {
      throw new Error('Set TEMPLATE_DOCX_EMP_FILE_ID to the employee DOCX file id in Drive.');
    }

    const name        = col.name        >= 0 ? text(row[col.name])        : '';
    const designation = col.designation >= 0 ? text(row[col.designation]) : '';
    const doj         = col.doj         >= 0 ? toDate(row[col.doj])       : null;
    const salary      = col.salary      >= 0 ? text(row[col.salary])      : '';

    // ✅ LOI validity date from "Valid Till" column
    const validTill   = col.validTill   >= 0 ? toDate(row[col.validTill]) : null;

    const letterDate  = col.letterDate  >= 0 ? toDate(row[col.letterDate]) : new Date();

    const fileNameBase = `LOI - ${name} - ${fmtDate(doj)}`;

    const copyResource = { title: fileNameBase, mimeType: 'application/vnd.google-apps.document' };
    if (OUTPUT_FOLDER_ID) copyResource.parents = [{id: OUTPUT_FOLDER_ID}];
    const gdocFile = Drive.Files.copy(copyResource, TEMPLATE_DOCX_EMP_FILE_ID);
    const gdocId   = gdocFile.id;

    const doc  = DocumentApp.openById(gdocId);
    const body = doc.getBody();
    const safe = s => (s == null ? '' : String(s));

    const map = [
      ['\\(\\s*Name\\s*\\)',          safe(name)],
      ['\\(\\s*Designation\\s*\\)',   safe(designation)],
      ['\\(\\s*Date of Joining\\s*\\)', fmtDate(doj)],
      ['\\bXXXX\\b',                  safe(salary)], // CTC placeholder
      ['Date:\\s*',                  'Date: ' + fmtDate(letterDate) + ' '],
    ];
    map.forEach(([pattern, repl]) => body.replaceText(pattern, repl));

  // ✅ Use the sheet Valid Till date for the heading + note
    if (validTill) {
      const vtHeader = fmtDateLong(validTill).toUpperCase();
      const vtNote   = fmtDateLong(validTill);

      body.replaceText('VALID\\s*TILL\\s*.*', 'VALID TILL ' + vtHeader);
      body.replaceText('i\\.e\\.\\s*[^,\\n]+', 'i.e. ' + vtNote);
    }

    doc.saveAndClose();
    const pdfBlob = DriveApp.getFileById(gdocId).getAs(MimeType.PDF).setName(fileNameBase + '.pdf');
    return { pdfBlob, gdocId };
  }

  function makeInternPdfFromTemplate(row) {
    if (!TEMPLATE_DOCX_INTERN_FILE_ID || TEMPLATE_DOCX_INTERN_FILE_ID === 'PASTE_INTERNSHIP_DOCX_FILE_ID_HERE') {
      throw new Error('Set TEMPLATE_DOCX_INTERN_FILE_ID to the intern DOCX file id in Drive.');
    }

    const name        = col.name        >= 0 ? text(row[col.name])        : '';
    const designation = (col.designation >= 0 ? text(row[col.designation]) : '') || 'Intern';

    // 👇 Internship start
    const doj         = col.doj       >= 0 ? toDate(row[col.doj])       : null;

    // 👇 Internship end / duration end from "End Date" column
    const endDate     = col.endDate   >= 0 ? toDate(row[col.endDate])   : null;

    // 👇 LOI validity from "Valid Till" column
    const validTill   = col.validTill >= 0 ? toDate(row[col.validTill]) : null;

    const stipend     = col.salary    >= 0 ? text(row[col.salary])      : '';
    const letterDate  = col.letterDate >= 0 ? toDate(row[col.letterDate]) : new Date();

    const fileNameBase = `LOI (Intern) - ${name} - ${fmtDate(doj)}`;

    const copyResource = { title: fileNameBase, mimeType: 'application/vnd.google-apps.document' };
    if (OUTPUT_FOLDER_ID) copyResource.parents = [{id: OUTPUT_FOLDER_ID}];
    const gdocFile = Drive.Files.copy(copyResource, TEMPLATE_DOCX_INTERN_FILE_ID);
    const gdocId   = gdocFile.id;

    const doc  = DocumentApp.openById(gdocId);
    const body = doc.getBody();

    body.replaceText('\\(\\s*Name\\s*\\)', name);
    body.replaceText('Date:\\s*', 'Date: ' + fmtDateLong(letterDate) + ' ');

    // ✅ LOI VALID TILL from sheet
    if (validTill) {
      const vtHeader = fmtDateLong(validTill).toUpperCase();
      const vtNote   = fmtDateLong(validTill);
      body.replaceText('VALID\\s*TILL\\s*.*', 'VALID TILL ' + vtHeader);
      body.replaceText('i\\.e\\.\\s*[^,\\n]+', 'i.e. ' + vtNote);
    }

    // Replace “AI Intern” with actual designation
    const quotedDesignation = '“' + designation + '”';
    body.replaceText('[“”]?AI\\s*Intern[“”]?', quotedDesignation);

    // ✅ Commence = DOJ  |  Conclude = End Date (duration end)
    if (doj)     body.replaceText('(?i)commence on\\s*[0-9]{1,2}(st|nd|rd|th)?\\s+[A-Za-z]+\\s+[0-9]{4}', 'commence on ' + fmtDateLong(doj));
    if (endDate) body.replaceText('(?i)conclude on\\s*[0-9]{1,2}(st|nd|rd|th)?\\s+[A-Za-z]+\\s+[0-9]{4}', 'conclude on ' + fmtDateLong(endDate));

    // Stipend
    if (stipend) {
      const s = String(stipend).trim();
      const stipendText = /^(?:Rs\.?|₹)/i.test(s) ? s : `Rs. ${s}`;
      body.replaceText('(Rs\\.?|₹)\\s*(?:XXXX|[0-9,]+)(?:\\s*/-)?', stipendText);
    }

    // Date of joining sentence
    if (doj) body.replaceText('(?i)\\bis\\s*[0-9]{1,2}(st|nd|rd|th)?\\s+[A-Za-z]+\\s+[0-9]{4}', 'is ' + fmtDateLong(doj));

    doc.saveAndClose();
    const pdfBlob = DriveApp.getFileById(gdocId).getAs(MimeType.PDF).setName(fileNameBase + '.pdf');
    return { pdfBlob, gdocId };
  }

  /***** MAIN SINGLE-PASS LOOP *****/
  const today = toMidnight(new Date());

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    // ✅ NEW: Gate LOI sending on required columns being filled for that person
    const readyForLOI = isRowCompleteForLOI(row, r + 1);

    // A) APPROVERS ONLY: send LOI (intern/employee template) once, when not already sent
    const doj = col.doj >= 0 ? toDate(row[col.doj]) : null;
    const approvalStatus = col.approvalMailStatus >= 0 ? text(row[col.approvalMailStatus]) : '';

    if (readyForLOI && doj && !looksSent(approvalStatus)) {
      const intern = isInternish(row);

      // Use same long HTML body as candidate, but with preview subject
      const { subject, htmlBody } = buildApprovalMail(row);

      // Attach the generated PDF for approvers to review
      let attachments = null;
      try {
        const { pdfBlob } = intern
          ? makeInternPdfFromTemplate(row)
          : makeEmployeePdfFromTemplate(row);
        attachments = [pdfBlob];
      } catch (e) {
        Logger.log(`Row ${r+1}: Approver PDF generation failed: ${e}`);
      }

      try {
        sendEmail(approvalRecipients.join(','), subject, htmlBody, null, attachments);
        if (col.approvalMailStatus >= 0) {
          sheet.getRange(r + 1, col.approvalMailStatus + 1)
              .setValue(`Sent ${fmtDateTime(new Date())}`);
        } else {
          Logger.log(`Row ${r+1}: Can't mark Approval Mail Status (column missing).`);
        }
      } catch (e) {
        Logger.log(`Row ${r+1}: Approver email failed: ${e}`);
      }
    }

    // B) CANDIDATE SENDS
    const approved = col.approvalOfLOI >= 0 ? isYes(row[col.approvalOfLOI]) : false;
    const email = col.email >= 0 ? text(row[col.email]) : '';
    const intern = isInternish(row);
    const bgCellRaw   = col.backgroundRequired >= 0 ? text(row[col.backgroundRequired]) : '';
    const bgCellLower = bgCellRaw.toLowerCase();
    const bgValid     = (bgCellLower === 'yes' || bgCellLower === 'no');

    if (approved && intern && readyForLOI) {
      if (!email) {
        Logger.log(`Row ${r+1}: Intern approved but Email is blank.`);
      } else {
        const current = col.internStatus >= 0 ? text(row[col.internStatus]) : '';
        if (!looksSent(current)) {
          const { subject, htmlBody } = buildInternLOI(row);
          let attachments = null;
          try {
            const { pdfBlob } = makeInternPdfFromTemplate(row);
            attachments = [pdfBlob];
          } catch (e) {
            Logger.log(`Row ${r+1}: Intern PDF generation failed: ${e}`);
          }
          try {
            sendEmail(email, subject, htmlBody, ccOnFinal, attachments);
            if (col.internStatus >= 0) {
              sheet.getRange(r + 1, col.internStatus + 1).setValue(`Sent ${fmtDateTime(new Date())}`);
            } else {
              Logger.log(`Row ${r+1}: Can't mark LOI Intern Status (column missing or mismatched header).`);
            }
          } catch (e) {
            Logger.log(`Row ${r+1}: Intern LOI send failed: ${e}`);
          }
        } else {
          Logger.log(`Row ${r+1}: Intern LOI already marked Sent.`);
        }
      }
    }

    if (approved && !intern && readyForLOI) {
      if (!doj) {
        Logger.log(`Row ${r+1}: Employee approved but DOJ is missing.`);
      } else if (!bgValid) {
        Logger.log(
          `Row ${r+1}: Employee approved but "Background Check Required" must be "Yes" or "No" (found "${bgCellRaw}").`
        );
      } else if (!email) {
        Logger.log(`Row ${r+1}: Employee approved but Email is blank.`);
      } else {
        const current = col.employeeStatus >= 0 ? text(row[col.employeeStatus]) : '';
        if (!looksSent(current)) {
          const { subject, htmlBody } = buildEmployeeLOI(row);
          let attachments = null;
          try {
            const { pdfBlob } = makeEmployeePdfFromTemplate(row);
            attachments = [pdfBlob];
          } catch (e) {
            Logger.log(`Row ${r+1}: Employee PDF generation failed: ${e}`);
          }
          try {
            sendEmail(email, subject, htmlBody, ccOnFinal, attachments);
            if (col.employeeStatus >= 0) {
              sheet.getRange(r + 1, col.employeeStatus + 1)
                  .setValue(`Sent ${fmtDateTime(new Date())}`);
            } else {
              Logger.log(`Row ${r+1}: Can't mark LOI Employee Status (column missing).`);
            }
          } catch (e) {
            Logger.log(`Row ${r+1}: Employee LOI send failed: ${e}`);
          }
        } else {
         Logger.log(`Row ${r+1}: Employee LOI already marked Sent.`);
        }
      }
    }
  }
  Logger.log('runLOIAutomation complete. DRY_RUN = ' + DRY_RUN);
}
