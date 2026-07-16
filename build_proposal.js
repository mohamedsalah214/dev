const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  LevelFormat, PositionalTab, PositionalTabAlignment, PositionalTabLeader
} = require('docx');
const fs = require('fs');

const NAVY = "1F3864";
const ACCENT = "2E74B5";
const GREY = "595959";
const LIGHT = "EAF0F7";

// ---------- helpers ----------
const H1 = (t) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 280, after: 120 },
  children: [new TextRun({ text: t, bold: true, color: NAVY, size: 30 })],
});

const H2 = (t) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 180, after: 80 },
  children: [new TextRun({ text: t, bold: true, color: ACCENT, size: 24 })],
});

const P = (t, opts = {}) => new Paragraph({
  spacing: { after: 120, line: 276 },
  children: [new TextRun({ text: t, size: 21, color: opts.color || "222222", italics: !!opts.italics, bold: !!opts.bold })],
});

const bullet = (t, boldLead) => new Paragraph({
  numbering: { reference: "checks", level: 0 },
  spacing: { after: 70, line: 264 },
  children: boldLead
    ? [new TextRun({ text: boldLead + ": ", bold: true, size: 21, color: "222222" }), new TextRun({ text: t, size: 21, color: "222222" })]
    : [new TextRun({ text: t, size: 21, color: "222222" })],
});

const cell = (text, { bold = false, header = false, align = AlignmentType.LEFT, width } = {}) =>
  new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: header ? { type: ShadingType.CLEAR, fill: NAVY, color: "auto" } : { type: ShadingType.CLEAR, fill: "FFFFFF", color: "auto" },
    margins: { top: 60, bottom: 60, left: 110, right: 110 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, bold: bold || header, size: 20, color: header ? "FFFFFF" : "222222" })],
    })],
  });

const shadedCell = (text, { bold = false, align = AlignmentType.LEFT, width, fill = LIGHT } = {}) =>
  new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill, color: "auto" },
    margins: { top: 60, bottom: 60, left: 110, right: 110 },
    children: [new Paragraph({ alignment: align, children: [new TextRun({ text, bold, size: 20, color: "222222" })] })],
  });

const divider = () => new Paragraph({
  spacing: { before: 60, after: 120 },
  border: { bottom: { color: "C8D6E5", style: BorderStyle.SINGLE, size: 8 } },
  children: [new TextRun({ text: "" })],
});

// ---------- document body ----------
const children = [];

// Header / title block
children.push(new Paragraph({
  spacing: { after: 40 },
  children: [new TextRun({ text: "PARKSIDE", bold: true, size: 40, color: NAVY, characterSpacing: 40 })],
}));
children.push(new Paragraph({
  spacing: { after: 60 },
  children: [new TextRun({ text: "Website & Technical Management — Proposal", size: 24, color: ACCENT, bold: true })],
}));
children.push(new Paragraph({
  border: { bottom: { color: NAVY, style: BorderStyle.SINGLE, size: 12 } },
  spacing: { after: 160 },
  children: [new TextRun({ text: "" })],
}));

// Meta table
children.push(new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [1600, 3080, 1600, 3080],
  borders: {
    top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
    left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
    insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE },
  },
  rows: [
    new TableRow({ children: [
      shadedCell("Prepared for", { bold: true, width: 1600, fill: LIGHT }),
      shadedCell("Eng. Mohamed Ibrahim — CEO, Parkside", { width: 3080, fill: "FFFFFF" }),
      shadedCell("Date", { bold: true, width: 1600, fill: LIGHT }),
      shadedCell("July 16, 2026", { width: 3080, fill: "FFFFFF" }),
    ]}),
    new TableRow({ children: [
      shadedCell("Prepared by", { bold: true, width: 1600, fill: LIGHT }),
      shadedCell("Mohamed Mahmoud", { width: 3080, fill: "FFFFFF" }),
      shadedCell("Domain", { bold: true, width: 1600, fill: LIGHT }),
      shadedCell("parkside-dev.com", { width: 3080, fill: "FFFFFF" }),
    ]}),
    new TableRow({ children: [
      shadedCell("Role", { bold: true, width: 1600, fill: LIGHT }),
      shadedCell("Technical & Digital Solutions Manager", { width: 3080, fill: "FFFFFF" }),
      shadedCell("Contact", { bold: true, width: 1600, fill: LIGHT }),
      shadedCell("01124290934", { width: 3080, fill: "FFFFFF" }),
    ]}),
  ],
}));

children.push(new Paragraph({ spacing: { after: 100 }, children: [] }));

// Intro
children.push(P("This document summarizes the technical work already delivered for Parkside over the past year, outlines the ongoing management plan going forward, and sets out clear, transparent pricing so we can align and continue building on a solid, professional foundation."));

children.push(divider());

// ========== PART 1: DELIVERED ==========
children.push(H1("Part 1 — Work Delivered (Past 12 Months)"));
children.push(P("The following work has already been completed and is in place. All access, credentials, and control of Parkside's digital assets are held and managed securely on the company's behalf."));

children.push(bullet("Domain registered, secured, and activated under Parkside's account (parkside-dev.com).", "Domain Registration"));
children.push(bullet("Hosting plan booked and configured, with the website environment set up and ready.", "Hosting Setup"));
children.push(bullet("Around 10 official company email accounts created under the domain (e.g. name@parkside-dev.com), configured for webmail and mobile/desktop use.", "Corporate Emails"));
children.push(bullet("A professional branded email signature designed (name, title, phone, logo, and social links) ready to be rolled out across all accounts.", "Email Signature Design"));
children.push(bullet("Centralized, secure management of all hosting, domain, and email credentials — protecting Parkside's assets and ensuring continuity.", "Access & Security"));

children.push(new Paragraph({
  spacing: { before: 60, after: 120 },
  shading: { type: ShadingType.CLEAR, fill: LIGHT, color: "auto" },
  border: {
    left: { color: ACCENT, style: BorderStyle.SINGLE, size: 24 },
    top: { color: LIGHT, style: BorderStyle.SINGLE, size: 2 },
    bottom: { color: LIGHT, style: BorderStyle.SINGLE, size: 2 },
    right: { color: LIGHT, style: BorderStyle.SINGLE, size: 2 },
  },
  children: [new TextRun({ text: "Over the past year this work has kept Parkside's domain, hosting, and email running reliably and under secure control — the essential digital foundation the business depends on daily.", italics: true, size: 21, color: GREY })],
}));

children.push(divider());

// ========== PART 2: GOING FORWARD ==========
children.push(H1("Part 2 — Going Forward: Ongoing Management"));
children.push(P("The immediate next step is deploying the website, followed by the ongoing scope that keeps Parkside's digital ecosystem functional, secure, and up to date."));

children.push(H2("Website Deployment & Launch (next step)"));
children.push(bullet("Taking the developed website and uploading it to Parkside's hosting, then connecting it to the domain."));
children.push(bullet("Verifying all pages, links, and forms work correctly across browsers and devices."));
children.push(bullet("Optimizing performance and loading speed, then taking the site fully live."));
children.push(bullet("Handover safeguard: a full copy of the site files (and database, if any) is secured under Parkside's control."));

children.push(H2("Technical Maintenance & Monitoring"));
children.push(bullet("Regular monitoring of website performance, server health, and domain status."));
children.push(bullet("Monthly minor updates, backups, and security patches."));

children.push(H2("Website Management"));
children.push(bullet("Ongoing optimization of speed, structure, and links."));
children.push(bullet("Adding or adjusting small content updates as needed (up to 10 per month)."));
children.push(bullet("Ensuring compatibility across browsers and mobile devices."));

children.push(H2("Email & Signature Support"));
children.push(bullet("Maintenance of all corporate emails and quick help with delivery, quota, or login issues."));
children.push(bullet("Rollout and installation of the branded email signature across all accounts, on web and mobile."));

children.push(H2("Social Media & Security"));
children.push(bullet("Keeping official pages active, secure, and properly branded."));
children.push(bullet("Basic security configuration (SSL, spam protection, uptime monitoring)."));

children.push(H2("Backup & Recovery"));
children.push(bullet("Monthly full backup of website and email data, with rapid restoration if issues occur."));

children.push(divider());

// ========== PART 3: INVESTMENT ==========
children.push(H1("Part 3 — Investment & Pricing"));

children.push(H2("Services"));
children.push(new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [4400, 2960, 2000],
  rows: [
    new TableRow({ tableHeader: true, children: [
      cell("Item", { header: true, width: 4400 }),
      cell("Details", { header: true, width: 2960 }),
      cell("Amount", { header: true, width: 2000, align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      cell("Setup Phase (one-time)", { bold: true, width: 4400 }),
      cell("Setup incl. website deployment", { width: 2960 }),
      cell("10,000 EGP", { width: 2000, align: AlignmentType.RIGHT, bold: true }),
    ]}),
    new TableRow({ children: [
      cell("Past-year management (goodwill)", { bold: true, width: 4400 }),
      cell("Discounted lump sum for the past 12 months", { width: 2960 }),
      cell("12,000 EGP", { width: 2000, align: AlignmentType.RIGHT, bold: true }),
    ]}),
    new TableRow({ children: [
      cell("Ongoing monthly management", { bold: true, width: 4400 }),
      cell("Starting from this month", { width: 2960 }),
      cell("2,000–3,000 EGP / mo", { width: 2000, align: AlignmentType.RIGHT, bold: true }),
    ]}),
  ],
}));

children.push(new Paragraph({ spacing: { after: 60 }, children: [] }));
children.push(P("The monthly fee starts at 2,000 EGP, covering all essential maintenance, monitoring, backups, and up to 10 content updates per month. It may rise to a maximum of 3,000 EGP only if additional requests, extended support, or more than 10–15 updates are required in a given month.", { italics: true }));

children.push(H2("Hosting & Domain (annual, at cost)"));
children.push(new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [3400, 3480, 2480],
  rows: [
    new TableRow({ tableHeader: true, children: [
      cell("Item", { header: true, width: 3400 }),
      cell("Details", { header: true, width: 3480 }),
      cell("Price", { header: true, width: 2480, align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      cell("Hosting — Stellar", { bold: true, width: 3400 }),
      cell("Up to 30 email accounts", { width: 3480 }),
      cell("$27 / yr ($59 renewal)", { width: 2480, align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      cell("Hosting — Plus", { bold: true, width: 3400 }),
      cell("Unlimited email accounts", { width: 3480 }),
      cell("$35 / yr ($88 renewal)", { width: 2480, align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      cell("Domain", { bold: true, width: 3400 }),
      cell("Already secured & active", { width: 3480 }),
      cell("$12 / yr ($17 renewal)", { width: 2480, align: AlignmentType.RIGHT }),
    ]}),
  ],
}));

children.push(new Paragraph({ spacing: { after: 100 }, children: [] }));
children.push(H2("Payment Terms"));
children.push(bullet("Setup Phase: 50% down payment, 50% on completion — the setup fee includes deploying and launching the developed website, and the remaining balance falls due once the site is live."));
children.push(bullet("Past-year management: one-time settlement, payable on agreement."));
children.push(bullet("Ongoing management: billed monthly."));

children.push(divider());

// Sign-off
children.push(new Paragraph({ spacing: { before: 100, after: 40 }, children: [new TextRun({ text: "Best regards,", size: 21, color: "222222" })] }));
children.push(new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: "Mohamed Mahmoud", bold: true, size: 22, color: NAVY })] }));
children.push(new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: "Technical & Digital Solutions Manager", size: 20, color: GREY })] }));
children.push(new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: "mohamedmahmoud214@gmail.com  •  01124290934", size: 20, color: GREY })] }));

// ---------- assemble ----------
const doc = new Document({
  numbering: {
    config: [{
      reference: "checks",
         levels: [{
        level: 0, format: LevelFormat.BULLET, text: "✓", alignment: AlignmentType.LEFT,
        style: { run: { color: ACCENT, bold: true }, paragraph: { indent: { left: 360, hanging: 240 } } },
      }],
    }],
  },
  styles: { default: { document: { run: { font: "Calibri" } } } },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, bottom: 1080, left: 1440, right: 1440 } } },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("/sessions/inspiring-vibrant-brahmagupta/mnt/outputs/Parkside_Proposal.docx", buf);
  console.log("written");
});
