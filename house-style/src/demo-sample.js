/* Sample reply used only by the browser preview. It mimics a new-Outlook
   reply: default compose font, unstyled paragraphs ready for house styles,
   text pasted from a draft with stray fonts and highlighting, a quoted
   thread, and a few attachments. All names and details are fictional. */
(function (root) {
  "use strict";
  root.DEMO_SAMPLE = {
    subject: "RE: Project Cedar – SPA comments",
    body: [
      "<div>Dear Rania,</div>",
      "<div><br></div>",
      "<div>Thank you for your comments on the draft share purchase agreement. Our responses are set out below.</div>",
      "<div><br></div>",
      "<div>Completion</div>",
      "<div>The buyer has asked for three changes:</div>",
      "<div>completion to take place within ten business days of the last condition being satisfied</div>",
      "<div>the long-stop date to move to 31 January 2027</div>",
      "<div>the seller to deliver signed resignation letters from the outgoing directors</div>",
      "<div><br></div>",
      "<div>Conditions</div>",
      "<div>Clause 4.1 currently reads:</div>",
      "<div><span style=\"font-family:'Times New Roman',serif;font-size:13.5pt;color:#333366\">“Completion is conditional on the Buyer receiving written confirmation from the Companies Control Department that the transfer of the Sale Shares has been registered.”</span></div>",
      "<div><span style=\"font-family:Calibri,sans-serif;font-size:15px;background-color:yellow\">We suggest removing “written”, so that confirmation by email is enough.</span></div>",
      "<div><br></div>",
      "<div>Please let me know if you would like to discuss any of this before Thursday.</div>",
      "<div><br></div>",
      "<div>Kind regards</div>",
      "<div id=\"appendonsend\"></div>",
      "<hr>",
      "<div id=\"divRplyFwdMsg\" style=\"font-family:Calibri,sans-serif;font-size:11pt;color:#000000\"><b>From:</b> Rania Haddad &lt;r.haddad@client.example&gt;<br><b>Sent:</b> 16 October 2026 18:05<br><b>Subject:</b> Project Cedar – SPA comments</div>",
      "<div><br></div>",
      "<div style=\"font-family:Calibri,sans-serif;font-size:11pt;color:#444444\">Hello, please find our comments on the SPA attached. The buyer would like to sign next week if possible. Best, Rania</div>",
    ].join(""),
    attachments: [
      { name: "SPA - draft v4 (clean).docx", size: 325632, type: "file" },
      { name: "SPA - draft v4 (compare v3).pdf", size: 1468006, type: "file" },
      { name: "Disclosure letter.docx", size: 98304, type: "file" },
      { name: "Board resolution - seller.pdf", size: 188416, type: "file" },
      { name: "RE: Completion mechanics", size: 55296, type: "item" },
    ],
  };
})(typeof self !== "undefined" ? self : this);
