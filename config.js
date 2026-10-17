/*
 * Legal Toolkit: firm-wide settings.
 *
 * This is the ONE file the firm edits before hosting the add-in.
 * Everyone who uses the add-in gets these defaults. Individual users can
 * override fonts, domains and the matter-number rule from the add-in's
 * Settings tab (their overrides are saved in their own mailbox).
 */
window.LT_CONFIG = {
  firmName: "Our Firm",

  // Email domains that count as "internal". Leave empty to use the domain
  // of the signed-in user's own email address.
  firmDomains: [],

  // Text styles used by the Format menu.
  style: {
    fontFamily: "Calibri, Arial, Helvetica, sans-serif",
    bodySizePt: 11,
    heading1SizePt: 14,
    heading2SizePt: 12,
    headingColor: "#1F3864",
    textColor: "#000000"
  },

  // Matter numbers. {matter} is replaced with the number the user types.
  matter: {
    subjectFormat: "[{matter}]",
    // Regular expression a matter number must match (used by the pre-send
    // check to recognise a matter number already in the subject).
    pattern: "\\b\\d{3,}[-./]\\d{2,}\\b",
    requireInSubject: false
  },

  // Privilege and confidentiality markings (Mark menu). The ids must stay
  // the same because the ribbon buttons refer to them; the wording can change.
  markings: {
    privileged: {
      label: "Privileged & Confidential",
      banner: "PRIVILEGED & CONFIDENTIAL — ATTORNEY-CLIENT COMMUNICATION",
      subjectTag: "[PRIVILEGED & CONFIDENTIAL]"
    },
    workProduct: {
      label: "Attorney Work Product",
      banner: "PRIVILEGED & CONFIDENTIAL — ATTORNEY WORK PRODUCT — PREPARED IN ANTICIPATION OF LITIGATION",
      subjectTag: "[ATTORNEY WORK PRODUCT]"
    },
    confidential: {
      label: "Confidential",
      banner: "CONFIDENTIAL",
      subjectTag: "[CONFIDENTIAL]"
    },
    settlement: {
      label: "Settlement Communication",
      banner: "CONFIDENTIAL SETTLEMENT COMMUNICATION — FOR SETTLEMENT PURPOSES ONLY",
      subjectTag: "[SETTLEMENT COMMUNICATION]"
    },
    withoutPrejudice: {
      label: "Without Prejudice",
      banner: "WITHOUT PREJUDICE",
      subjectTag: "[WITHOUT PREJUDICE]"
    },
    draft: {
      label: "Draft — For Discussion",
      banner: "DRAFT — FOR DISCUSSION PURPOSES ONLY — SUBJECT TO FURTHER REVIEW",
      subjectTag: "[DRAFT]"
    }
  },
  markingColor: "#C00000",

  // Pre-send check.
  check: {
    maxTotalAttachmentMB: 20,
    largeDistribution: 15,
    // Attachment file names containing these words are flagged for a second look.
    attachmentWarnWords: ["draft", "internal", "privileged", "notes", "comments", "redline", "do not send"]
  },

  // Standard paragraphs (Clauses tab). {{Field name}} placeholders are
  // filled in by the user before the text is inserted.
  // Have your risk / general counsel team review this wording before rollout.
  clauses: [
    { group: "Sending documents", title: "Please find attached",
      text: "Please find attached {{document description}} for your review. Please let me know if you have any questions or comments." },
    { group: "Sending documents", title: "Draft for review",
      text: "Attached is a draft of {{document}} for your review. The draft remains subject to further review and comment, including by our client." },
    { group: "Sending documents", title: "Comparison enclosed",
      text: "For convenience, we have also attached a comparison showing the changes made since the draft dated {{date}}." },
    { group: "Sending documents", title: "Execution version",
      text: "Attached is the execution version of {{document}}. Please arrange for signature where indicated and return a signed copy to us by {{date}}." },

    { group: "Privilege & confidentiality", title: "Do not forward (privilege)",
      text: "This email contains legal advice and is protected by legal professional privilege. Please do not forward it outside your organization without speaking to us first, as doing so may waive privilege." },
    { group: "Privilege & confidentiality", title: "No waiver",
      text: "Nothing in this communication is intended to waive, and nothing should be construed as a waiver of, any privilege or protection, all of which are expressly reserved." },
    { group: "Privilege & confidentiality", title: "Reservation of rights",
      text: "This communication is without prejudice to, and with full reservation of, all of our client's rights and remedies, whether at law, in equity or otherwise, none of which are waived." },
    { group: "Privilege & confidentiality", title: "Confidentiality notice",
      text: "This email and any attachments are confidential and may be privileged. If you are not the intended recipient, please notify the sender immediately, delete this email and do not copy, use or disclose its contents." },

    { group: "Litigation", title: "Litigation hold",
      text: "Please preserve all documents and electronically stored information that may be relevant to {{matter}}, including emails, text and instant messages, and files stored on personal devices. Do not delete, discard or alter any such material until we confirm in writing that this hold has been lifted." },
    { group: "Litigation", title: "For settlement purposes only",
      text: "This communication is made for settlement purposes only and may not be used as evidence of liability or for any other purpose." },
    { group: "Litigation", title: "Deadline reminder",
      text: "As a reminder, the deadline for {{step}} is {{date}}. Please send us any comments or instructions by {{internal date}} so that we can meet this deadline." },

    { group: "Client service", title: "Scope of advice",
      text: "Our advice is limited to the matters set out above and is based on the facts as you have described them to us. Please let us know if any of those facts are inaccurate or incomplete, as this may affect our advice." },
    { group: "Client service", title: "Acknowledge receipt",
      text: "Thank you for your email of {{date}}. We are reviewing it and will respond by {{response date}}." },
    { group: "Client service", title: "Happy to discuss",
      text: "I would be happy to discuss this at your convenience. Please let me know a time that suits you." }
  ]
};
