function sendOffboardingEmails(){
  const COMPANY_NAME = "Nlink Tech Pvt Ltd";
  const HR_TEAM = "ad@nlink.tech";
  // Backend-managed manager email (set this on your side; can be blank if not needed)

  // === Inline image setup (Drive file IDs) ===
  const BANNER_FILE_ID   = '1GLvDYAB8w_K81MvzWQvhpeB7zWrGOKLb';
  const TECH_FILE_ID     = '1_8T73h0rCtitBRLfgQ4ajgqVEaNiRiGh';
  
  // CID keys used inside HTML
  const BANNER_CID       = 'banner';
  const TECH_LOGO_CID    = 'techlogo';
  

  // Preload blobs once
  const inlineImagesBase = {
    [BANNER_CID]:    DriveApp.getFileById(BANNER_FILE_ID).getBlob(),
    [TECH_LOGO_CID]: DriveApp.getFileById(TECH_FILE_ID).getBlob()
  };

  
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


  // === Tech signature block (uses cid:) ===
  const TECH_HTML = `
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

  const DISC = `
  <!-- Disclaimer block: gradient bar + dark panel -->
  <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="width:100%;max-width:700px;margin:0 auto;border-collapse:collapse;">
    <tr>
      <td>
        <div style="height:12px;background:linear-gradient(90deg,#e9eefb,#c8ecf3,#d7c6f3);background-color:#e9eefb;"></div>
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

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    Logger.log("Another instance is running. Exiting.");
    return;
  }

  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName("Offboarding");
    if (!sheet) throw new Error("Sheet 'Offboarding' not found");

    const data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) {
      Logger.log("No data rows.");
      return;
    }

    const header = data[0];
    const tz = ss.getSpreadsheetTimeZone();

    const norm = v => (v == null ? "" : String(v).toLowerCase().trim());
    const isYes = v => ["y", "yes"].includes(norm(v));
    const isSent = v => norm(v) === "sent";
    const fmt = d => Utilities.formatDate(d, tz, "d MMM yyyy");

    // Robust parser
    const parseSheetDate = cell => {
      if (cell instanceof Date) {
        const d = new Date(cell.getFullYear(), cell.getMonth(), cell.getDate(), 12, 0, 0);
        return d;
      }
      if (cell == null) return null;
      const s = String(cell).trim();
      if (!s) return null;

      if (/^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}$/.test(s)) {
        const [y, m, d] = s.split(/[-\/]/).map(n => parseInt(n, 10));
        return new Date(y, m - 1, d, 12, 0, 0);
      }
      if (/^\d{1,2}[-\/]\d{1,2}[-\/]\d{4}$/.test(s)) {
        const [d, m, y] = s.split(/[-\/]/).map(n => parseInt(n, 10));
        return new Date(y, m - 1, d, 12, 0, 0);
      }
      const guess = new Date(s);
      if (!isNaN(guess.getTime())) {
        return new Date(guess.getFullYear(), guess.getMonth(), guess.getDate(), 12, 0, 0);
      }
      return null;
    };

    const today = new Date();
    const bannerHTML = ""; // kept for compatibility

    const footer = `
    <p>Best regards,<br>
    Sonali Rana<br>
    Human Resource<br>
    +91 9045293556<br></p>
    `;

    // Resolve column indices and validate
    const colMap = {
      NAME: header.indexOf("Employee Name"),
      EMAIL: header.indexOf("Personal Email"),
      MANAGER: header.indexOf("Manager Email"),
      EXIT_TYPE: header.indexOf("Exit Type"),
      LWD: header.indexOf("Last Working Day (LWD)"),
      ASSET_RETURNED: header.indexOf("Asset Returned (Y/N)"),
      ASSET_RETURN: header.indexOf("Asset Return Mail Status"),
      
      FNF: header.indexOf("FnF Mail Status"),
      TERMINATION: header.indexOf("Termination Mail Status"),
      
      CONTACT_ATTEMPT: header.indexOf("Contact Attempts Made"),
      ABSCONDING: header.indexOf("Absconding Mail Status"),
      EXIT_INTERVIEW: header.indexOf("Exit Interview Mail Status"),
      EXIT_DATE_INIT: header.indexOf("Exit Date Initiation"),
      FNF_DATE: header.indexOf("FnF Date"),
      NOTICE_PERIOD: header.indexOf("Notice Period (Days)"),
      EXIT_REASON: header.indexOf("Exit Reason"),
      EXIT_INTERVIEW_SUMMARY: header.indexOf("Exit Interview Summary"),
      EXIT_INTERVIEW_FEEDBACK: header.indexOf("Exit Feedback")
    };

    const missingCols = Object.entries(colMap)
      .filter(([, idx]) => idx === -1)
      .map(([k]) => k);
    if (missingCols.length) {
      throw new Error("Missing required columns: " + missingCols.join(", "));
    }

    // Process each row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row.every(c => c === "")) continue; // skip empty rows

      const exitType = norm(row[colMap.EXIT_TYPE]);
      const empName = row[colMap.NAME];

      const rawEmpEmail = row[colMap.EMAIL];
      const empEmail = norm(rawEmpEmail);

      const ccParts = [HR_TEAM].filter(Boolean);
      const ccList = Array.from(new Set(ccParts)).join(",");

      const assetReturned = norm(row[colMap.ASSET_RETURNED]);
      const fnfRequired = norm(row[colMap.FNF_REQUIRED]);
      const contactAttempt = norm(row[colMap.CONTACT_ATTEMPT]);

      // Read current LWD values
      let lwd = parseSheetDate(row[colMap.LWD]);
      let hasLwd = !!(lwd && !isNaN(lwd.getTime()));
      let lwdFormatted = hasLwd ? fmt(lwd) : "";
      const todayFormatted = fmt(today);

      // === resignation/termination/absconding: calculate LWD if missing ===
      if (exitType === "resignation" || exitType === "termination" || exitType === "absconding") {
        const exitDateInit = parseSheetDate(row[colMap.EXIT_DATE_INIT]);
        const noticeDays = parseInt(row[colMap.NOTICE_PERIOD], 10) || 0;

        if (exitDateInit && !hasLwd) {
          const calcLwd = new Date(exitDateInit);
          calcLwd.setHours(12, 0, 0, 0);
          calcLwd.setDate(calcLwd.getDate() + noticeDays);
          sheet.getRange(i + 1, colMap.LWD + 1).setValue(calcLwd);
          lwd = calcLwd;
          hasLwd = true;
          lwdFormatted = fmt(lwd);
        }
      }
      // === END LWD calc ===

      if (!empEmail) {
        if (ccList || row[colMap.MANAGER]) {
          Logger.log(`Row ${i + 1}: No personal email found. Will send only to manager/CC list (${[row[colMap.MANAGER], ccList].filter(Boolean).join(", ")}).`);
        } else {
          Logger.log(`Row ${i + 1}: No personal email and no manager/CC recipients. Skipping.`);
          continue;
        }
      }

      // Helper: send email with safe fallback (emp + manager), supports inline images and extra CC
      const sendSafely = ({ subject, htmlBody, ccExtra = "" }) => {
        const toList = [];
        if (empEmail) toList.push(empEmail);
        if (row[colMap.MANAGER]) toList.push(String(row[colMap.MANAGER]).trim());

        const toAddr = toList.join(", ");
        const ccAddr = [ccList, ccExtra].filter(Boolean).join(",");

        if (!toAddr) {
          Logger.log(`Row ${i + 1}: No recipients available. Skipping send.`);
          return false;
        }

        MailApp.sendEmail({ to: toAddr, cc: ccAddr, subject, htmlBody, inlineImages: inlineImagesBase });
        return true;
      };

      // 1) Asset Return Reminder
      if (
        hasLwd &&
        (assetReturned === "no" || assetReturned === "n") &&
        !isSent(row[colMap.ASSET_RETURN])
      ) {
        const subject = "Employee Exit: Asset Handover and Email Deactivation";
        let body =
          BANNER_HTML +            // wide banner
          BODY_OPEN + 
          `<p>Hi Team,</p>
          <p>I hope this message finds you well. I wanted to inform you about the termination of ${empName}, whose last working day is ${lwdFormatted}.</p>
          <h4>Asset Handover</h4>
          <p>Please note that ${empName} will need to return all company property by their final day, including but not limited to:</p>
          <ul>
            <li>Laptop</li>
            <li>Phone (if issued)</li>
            <li>ID Badge</li>
            <li>Other assigned assets</li>
          </ul>
          <p>Please ensure that the necessary arrangements are made for the collection of these items.</p>
          <h4>Email Account Deletion</h4>
          <p>Additionally, please prepare to disable and delete ${empName}'s company email account on ${lwdFormatted}. It is important that this is completed promptly to maintain our security protocols.</p>
          <h4>Finance</h4>
          <p>For finance, please ensure that ${empName}'s final paycheck is processed in accordance with company policies. This should include compensation for any unused vacation or paid time off, as well as any other applicable payments.</p>
          <p>If you have any questions or require further details, please let me know. Thank you for your cooperation in ensuring a smooth transition.</p>`;
        body += footer +
          `<div style="height:20px;"></div>` +
          TECH_HTML +
          `<div style="height:10px;"></div>` +
          DISC  +
          BODY_CLOSE ;

        try {
          const ok = sendSafely({ subject, htmlBody: body,ccExtra: "s2@corptive.com" });
          if (ok) {
            sheet.getRange(i + 1, colMap.ASSET_RETURN + 1).setValue('Sent');
          }
        } catch (err) {
          Logger.log(`Asset return mail failed (row ${i + 1}): ${err}`);
        }
      }

      // 2) FnF Mail (45+ calendar days after LWD AND assets returned)
      if (
        (exitType === "resignation" || exitType === "termination") &&
        parseSheetDate(row[colMap.EXIT_DATE_INIT]) &&
        isYes(assetReturned) &&
        !isSent(row[colMap.FNF])
      ) {
        const exitDateInit = parseSheetDate(row[colMap.EXIT_DATE_INIT]);
        const fnfDate = parseSheetDate(row[colMap.FNF_DATE]);

        let shouldSendFnF = false;

        // If FnF Date is given, send only on that date
        if (fnfDate) {
          shouldSendFnF = fmt(fnfDate) === fmt(today);
        } else {
          // Otherwise send after 45 days from Exit Date Initiation
          const daysSinceExitDateInit = Math.floor((today - exitDateInit) / (1000 * 60 * 60 * 24));
          shouldSendFnF = daysSinceExitDateInit >= 45;
        }

        if (shouldSendFnF) {
          const subject = "Full and Final Settlement Process";
          let body =
            BANNER_HTML +            // wide banner
            BODY_OPEN + 
            `<p>Dear ${empName || "Former Employee"},</p>
            <p>I hope this message finds you well.I am writing to confirm that we have completed your full and final settlement payment, which was processed on ${todayFormatted}.</p>
            <p>Please acknowledge receipt of this email.</p>
            <p>Thank you for your contributions during your time with us, and we wish you all the best in your future endeavors!</p>
            <p>If you have any questions, please feel free to reach out.</p>`;
          body += footer +
            `<div style="height:20px;"></div>` +
            TECH_HTML +
            `<div style="height:10px;"></div>` +
            DISC +
            BODY_CLOSE ;
          try {
            if (sendSafely({ subject, htmlBody: body })) {
              sheet.getRange(i + 1, colMap.FNF + 1).setValue("Sent");
            }
          } catch (err) {
            Logger.log(`FnF mail failed (row ${i + 1}): ${err}`);
          }
        }
      }

      // 3) Termination Mail
      if (exitType === "termination" && hasLwd && !isSent(row[colMap.TERMINATION])) {
        const subject = "Termination of Employment";
        let body =
          BANNER_HTML +            // wide banner
          BODY_OPEN +
          `<p>Dear ${empName || "Employee"},</p>
          <p>I hope this message finds you well.</p>
          <p>After careful consideration, we regret to inform you that your employment with ${COMPANY_NAME} will be terminated effective immediately. This decision is based on ongoing concerns regarding your performance, including repeated mistakes and a lack of adherence to our established work standards.</p>
          <p>Despite previous discussions regarding your performance and efforts to support you in your role, we have not seen the necessary improvements. Specifically, we had a meeting where we outlined our concerns and provided you with a verbal warning regarding the need for improvement.</p>
          <p>Please arrange to return any company property by today itself, and we will ensure that you receive your final paycheck and any remaining benefits as per company policy.</p>
          <p>If you have any questions or need assistance during this transition, please feel free to reach out.</p>
          <p>Thank you for your time with us, and we wish you the best in your future endeavors.</p>`;
        body += footer +
          `<div style="height:20px;"></div>` +
          TECH_HTML +
          `<div style="height:10px;"></div>` +
          DISC  +
          BODY_CLOSE ;
          
        try {
          if (sendSafely({ subject, htmlBody: body })) {
            sheet.getRange(i + 1, colMap.TERMINATION + 1).setValue("Sent");
          }
        } catch (err) {
          Logger.log(`Termination mail failed (row ${i + 1}): ${err}`);
        }
      }

      // 4) Absconding Mail
      if (
        exitType === "absconding" && hasLwd &&
        parseInt(contactAttempt, 10) === 3 &&
        !isSent(row[colMap.ABSCONDING])
      ) {
        const absDate = parseSheetDate(row[colMap.ABSCONDING_DATE]);
        const absDateText = absDate ? fmt(absDate) : String(row[colMap.ABSCONDING_DATE]);
        const subject = `Absconding Notification – ${(empName || "").trim()}`;
        let body =
          BANNER_HTML +            // wide banner
          BODY_OPEN +
          `<p>Dear ${empName || "Employee"},</p>
          <p>We are writing to inform you that due to your continued unexplained absence from work since ${lwdFormatted} and the lack of response to our previous communications, we are officially revoking your employment with ${COMPANY_NAME}, effective ${todayFormatted}.</p>
          <p>This decision is final, and your employment with us is considered concluded as of today.</p>
          <p>Please arrange to return any company property you may have.</p>
          <p>If you have any questions or need further information, please feel free to contact us.</p>
          <p>We wish you the best in your future endeavors.</p>`;
        body += footer +
          `<div style="height:20px;"></div>` +
          TECH_HTML +
          `<div style="height:10px;"></div>` +
          DISC  +
          BODY_CLOSE ;
        try {
          if (sendSafely({ subject, htmlBody: body })) {
            sheet.getRange(i + 1, colMap.ABSCONDING + 1).setValue("Sent");
          }
        } catch (err) {
          Logger.log(`Absconding mail failed (row ${i + 1}): ${err}`);
        }
      }

      // 5) Exit Interview Mail (to Manager from HR)
      if (
        (exitType === "resignation" || exitType === "termination") &&
        (row[colMap.EXIT_REASON] || "") !== "" &&
        (row[colMap.EXIT_INTERVIEW_FEEDBACK] || "") !== "" &&
        !isSent(row[colMap.EXIT_INTERVIEW])
      ) {
        const reason = (row[colMap.EXIT_REASON] || "").toString().trim();
        const summary = (row[colMap.EXIT_INTERVIEW_SUMMARY] || "").toString().trim() || "neutral";
        const feedback = (row[colMap.EXIT_INTERVIEW_FEEDBACK] || "").toString().trim();
        const subject = `Exit Interview Summary – ${empName || ""}`;
        let body =
          BANNER_HTML +            // wide banner
          BODY_OPEN +
          `<p>Dear Sir,</p>
          <p>The exit interview with ${empName || "the employee"} was completed on ${lwdFormatted}. They shared that their reason for leaving is ${reason}.</p>
          <p>Overall, their experience at Nlink Tech was ${summary}, and they expressed appreciation for the support from the team. They did mention ${feedback}, which could be helpful for future consideration.</p>
          <p>Let me know if you’d like to discuss any of this in more detail.</p>`;
        body += footer +
          `<div style="height:20px;"></div>` +
          TECH_HTML +
          `<div style="height:10px;"></div>` +
          DISC  +
          BODY_CLOSE ;

        try {
          MailApp.sendEmail({
            to: String(row[colMap.MANAGER] || "").trim(),
            cc: "",
            subject,
            htmlBody: body,
            inlineImages: inlineImagesBase
          });
          sheet.getRange(i + 1, colMap.EXIT_INTERVIEW + 1).setValue("Sent");
        } catch (err) {
          Logger.log(`Exit Interview mail failed (row ${i + 1}): ${err}`);
        }
      }
    } // end for rows
  } catch (err) {
    Logger.log("Error: " + err);
    throw err;
  } finally {
    lock.releaseLock();
  }
}