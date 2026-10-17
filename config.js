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
  }
};
