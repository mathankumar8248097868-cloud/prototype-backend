const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  AlignmentType,
  PageBreak,
  UnderlineType,
  Footer,
  TabStopType,
  TabStopPosition,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  ShadingType,
  VerticalAlign,
} = require("docx");

const axios = require("axios");
const fs    = require("fs");
const path  = require("path");
const db    = require("../config/db");

// ── Default logo paths (place logo-left.png / logo-right.png in backend/protected/) ──
const LOGO_LEFT_PATH  = path.join(__dirname, "../protected/logo-left.png");
const LOGO_RIGHT_PATH = path.join(__dirname, "../protected/logo-right.png");

// ── 3D-style bar chart via QuickChart ────────────────────────────────────────
async function generateChart(labels, data, xLabel, yLabel) {
  const palette = [
    "rgba(37,99,235,0.88)",  "rgba(14,165,233,0.88)",
    "rgba(16,185,129,0.88)", "rgba(245,158,11,0.88)",
    "rgba(239,68,68,0.88)",  "rgba(168,85,247,0.88)",
    "rgba(236,72,153,0.88)", "rgba(20,184,166,0.88)",
  ];
  const borders = palette.map((c) => c.replace("0.88", "1"));

  const config = {
    type: "bar",
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: labels.map((_, i) => palette[i % palette.length]),
        borderColor:      labels.map((_, i) => borders[i % borders.length]),
        borderWidth: 2,
        borderSkipped: false,
      }],
    },
    options: {
      legend: { display: false },
      scales: {
        xAxes: [{
          scaleLabel: { display: true, labelString: xLabel, fontSize: 14, fontStyle: "bold", fontColor: "#1e2b3c" },
          gridLines: { display: false },
          ticks: { fontColor: "#374151", fontSize: 12 },
        }],
        yAxes: [{
          scaleLabel: { display: true, labelString: yLabel, fontSize: 14, fontStyle: "bold", fontColor: "#1e2b3c" },
          ticks: { beginAtZero: true, fontColor: "#374151", fontSize: 12 },
          gridLines: { color: "rgba(0,0,0,0.07)", drawBorder: false },
        }],
      },
      plugins: {
        datalabels: {
          anchor: "end", align: "top",
          color: "#1e2b3c",
          font: { weight: "bold", size: 13 },
          formatter: (v) => v,
        },
      },
    },
  };

  const resp = await axios.post(
    "https://quickchart.io/chart",
    { chart: config, width: 700, height: 450, backgroundColor: "white", version: 2 },
    { responseType: "arraybuffer" }
  );
  return Buffer.from(resp.data);
}

// ── Pie chart ─────────────────────────────────────────────────────────────────
async function generatePieChart(labels, data) {
  const config = {
    type: "pie",
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: ["rgba(37,99,235,0.88)", "rgba(236,72,153,0.88)"],
        borderColor:      ["rgba(37,99,235,1)",    "rgba(236,72,153,1)"],
        borderWidth: 2,
      }],
    },
    options: {
      legend: { display: true, position: "bottom", labels: { fontSize: 13, fontColor: "#1e2b3c" } },
      plugins: {
        datalabels: {
          color: "#fff",
          font: { weight: "bold", size: 14 },
          formatter: (value, ctx) => {
            const total = ctx.chart.data.datasets[0].data.reduce((a, b) => a + b, 0);
            return total > 0 ? Math.round((value / total) * 100) + "%" : String(value);
          },
        },
      },
    },
  };

  const resp = await axios.post(
    "https://quickchart.io/chart",
    { chart: config, width: 600, height: 420, backgroundColor: "white", version: 2 },
    { responseType: "arraybuffer" }
  );
  return Buffer.from(resp.data);
}

// ── Border helpers ────────────────────────────────────────────────────────────
const thinBorder = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
const cellBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
const noBorder    = { style: BorderStyle.NONE,   size: 0, color: "FFFFFF" };
const noBorders   = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

exports.generateReport = async (req, res) => {
  try {
    const d      = req.body;
    const photos = req.files?.photos ?? [];

    // ── Read default logos from disk (graceful fallback if missing) ───────────
    const logoLeftBuf  = fs.existsSync(LOGO_LEFT_PATH)  ? fs.readFileSync(LOGO_LEFT_PATH)  : null;
    const logoRightBuf = fs.existsSync(LOGO_RIGHT_PATH) ? fs.readFileSync(LOGO_RIGHT_PATH) : null;
    const LOGO_SIZE = 95; // px

    const children = [];

    // ── Paragraph helpers ─────────────────────────────────────────────────────
    const heading = (text) =>
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: 360 },
        children: [new TextRun({
          text: (text || "").toUpperCase(),
          font: "Times New Roman", size: 28, bold: true,
          underline: { type: UnderlineType.SINGLE },
        })],
      });

    const normalText = (text, center = false) =>
      new Paragraph({
        alignment: center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
        spacing: { line: 480 },
        children: [new TextRun({ text: String(text ?? ""), font: "Times New Roman", size: 24 })],
      });

    const boldLabel = (text) =>
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 360 },
        children: [new TextRun({ text: String(text ?? ""), font: "Times New Roman", size: 24, bold: true })],
      });

    const blank = () => new Paragraph({ text: "", spacing: { line: 360 } });

    const centeredImage = (buf, w, h) =>
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120 },
        children: [new ImageRun({ data: buf, transformation: { width: w, height: h }, type: "png" })],
      });

    // ── Stats cell ────────────────────────────────────────────────────────────
    const statsCell = (text, w, isHeader = false) =>
      new TableCell({
        borders: cellBorders,
        width: { size: w, type: WidthType.DXA },
        shading: isHeader ? { fill: "D0D8E8", type: ShadingType.CLEAR } : undefined,
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: String(text ?? ""), font: "Times New Roman", size: 24, bold: isHeader })],
        })],
      });

    // ── Signature cells ───────────────────────────────────────────────────────
    const sigCell = (text, w) =>
      new TableCell({
        borders: cellBorders,
        width: { size: w, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: String(text ?? ""), font: "Times New Roman", size: 22 })],
        })],
      });

    const sigHeaderCell = (text, w) =>
      new TableCell({
        borders: cellBorders,
        width: { size: w, type: WidthType.DXA },
        shading: { fill: "D0D8E8", type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: String(text ?? ""), font: "Times New Roman", size: 22, bold: true })],
        })],
      });

    // ── Parse people lists ────────────────────────────────────────────────────
    const staffList  = d.staffList  ? JSON.parse(d.staffList)  : [];
    const pgList     = d.pgList     ? JSON.parse(d.pgList)     : [];
    const internList = d.internList ? JSON.parse(d.internList) : [];

    // ── Logo header table (left logo | centre text | right logo) ─────────────
    const logoCell = (buf, w) =>
      new TableCell({
        borders: noBorders,
        width: { size: w, type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: buf
            ? [new ImageRun({ data: buf, transformation: { width: LOGO_SIZE, height: LOGO_SIZE }, type: "png" })]
            : [new TextRun("")],
        })],
      });

    const makeLogoHeader = (lines) => {
      const centreChildren = lines.map((line) =>
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { line: 300 },
          children: [new TextRun({
            text: line.text || "",
            font: "Times New Roman",
            size: line.size || 22,
            bold: !!line.bold,
            underline: line.underline ? { type: UnderlineType.SINGLE } : undefined,
          })],
        })
      );

      // Left logo cell — flush to left corner
      const leftCornerCell = new TableCell({
        borders: noBorders,
        width: { size: 1440, type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
        children: [new Paragraph({
          alignment: AlignmentType.LEFT,
          children: logoLeftBuf
            ? [new ImageRun({ data: logoLeftBuf, transformation: { width: LOGO_SIZE, height: LOGO_SIZE }, type: "png" })]
            : [new TextRun("")],
        })],
      });

      // Right logo cell — flush to right corner
      const rightCornerCell = new TableCell({
        borders: noBorders,
        width: { size: 1440, type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: logoRightBuf
            ? [new ImageRun({ data: logoRightBuf, transformation: { width: LOGO_SIZE, height: LOGO_SIZE }, type: "png" })]
            : [new TextRun("")],
        })],
      });

      return new Table({
        alignment: AlignmentType.CENTER,
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1440, 6480, 1440],
        borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideH: noBorder, insideV: noBorder },
        rows: [new TableRow({
          children: [
            leftCornerCell,
            new TableCell({
              borders: noBorders,
              width: { size: 6480, type: WidthType.DXA },
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 0, bottom: 0, left: 60, right: 60 },
              children: centreChildren,
            }),
            rightCornerCell,
          ],
        })],
      });
    };

    // ── Signature table ───────────────────────────────────────────────────────
    const makeSignatureTable = (label, nameList) => {
      const nameW = 5580, sigW = 3780;
      const rows = [
        new TableRow({ children: [sigHeaderCell(label, nameW), sigHeaderCell("SIGNATURE", sigW)] }),
        ...nameList.map((n) => new TableRow({ children: [sigCell(n, nameW), sigCell("", sigW)] })),
      ];
      if (!nameList.length)
        rows.push(new TableRow({ children: [sigCell("", nameW), sigCell("", sigW)] }));
      return new Table({
        alignment: AlignmentType.CENTER,
        width: { size: nameW + sigW, type: WidthType.DXA },
        columnWidths: [nameW, sigW],
        rows,
      });
    };

    // ═══════════════════════════════════════════════════════════════════════════
    // PAGE 1 — CAMP CIRCULAR
    // ═══════════════════════════════════════════════════════════════════════════
    children.push(new Paragraph({
      spacing: { line: 360 },
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      children: [
        new TextRun({ text: `Ref No: ${d.refNo || ""}`, font: "Times New Roman", size: 22 }),
        new TextRun({ text: "\t" }),
        new TextRun({ text: `Date: ${d.reportDateShort || ""}`, font: "Times New Roman", size: 22 }),
      ],
    }));
    children.push(blank());

    children.push(new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { line: 360 },
      children: [new TextRun({ text: (d.departmentName || "").toUpperCase(), font: "Times New Roman", size: 28, bold: true, underline: { type: UnderlineType.SINGLE } })],
    }));
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { line: 360 },
      children: [new TextRun({ text: "CAMP CIRCULAR", font: "Times New Roman", size: 26, bold: true })],
    }));
    children.push(blank());

    // Clean body paragraph — no stray words
    children.push(new Paragraph({
      alignment: AlignmentType.JUSTIFIED, spacing: { line: 480 },
      children: [new TextRun({
        text: `A dental screening and treatment camp will be conducted at ${d.campLocation || ""} on ${d.reportDateLong || ""} from ${d.startTime || ""} to ${d.endTime || ""}. The bus will depart from the campus at ${d.busTime || d.startTime || ""}.`,
        font: "Times New Roman", size: 24,
      })],
    }));
    children.push(new Paragraph({
      alignment: AlignmentType.JUSTIFIED, spacing: { line: 480 },
      children: [new TextRun({
        text: "The following Staff, PG students, and interns are posted for the above camp.",
        font: "Times New Roman", size: 24,
      })],
    }));
    children.push(blank());

    children.push(boldLabel("STAFF:"));
    staffList.forEach((n) => children.push(normalText(n)));
    if (!staffList.length) children.push(normalText(""));
    children.push(blank());

    children.push(boldLabel("POSTGRADUATE:"));
    pgList.forEach((n) => children.push(normalText(n)));
    if (!pgList.length) children.push(normalText(""));
    children.push(blank());

    children.push(boldLabel("INTERNS:"));
    internList.forEach((n) => children.push(normalText(n)));
    if (!internList.length) children.push(normalText(""));

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ═══════════════════════════════════════════════════════════════════════════
    // PAGE 2 — ATTENDANCE SHEET
    // ═══════════════════════════════════════════════════════════════════════════
    children.push(makeLogoHeader([
      { text: (d.collegeName || "").toUpperCase(),    size: 26, bold: true, underline: true },
      { text: d.collegeAddress || "",                 size: 20 },
      { text: (d.departmentName || "").toUpperCase(), size: 22, bold: true, underline: true },
    ]));
    children.push(blank());

    children.push(new Paragraph({
      spacing: { line: 360 },
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      children: [
        new TextRun({ text: `CAMP SITE: ${d.campLocation || ""}`, font: "Times New Roman", size: 22, bold: true }),
        new TextRun({ text: "\t" }),
        new TextRun({ text: `DATE: ${d.reportDateShort || ""}`,   font: "Times New Roman", size: 22, bold: true }),
      ],
    }));
    children.push(blank());

    children.push(makeSignatureTable("STAFFS",        staffList));
    children.push(blank());
    children.push(makeSignatureTable("POST GRADUATE",  pgList));
    children.push(blank());
    children.push(makeSignatureTable("INTERNS",        internList));

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ═══════════════════════════════════════════════════════════════════════════
    // PAGE 3 — CAMP REPORT SUMMARY
    // ═══════════════════════════════════════════════════════════════════════════
    children.push(makeLogoHeader([
      { text: (d.collegeName || "").toUpperCase(),    size: 26, bold: true, underline: true },
      { text: d.collegeAddress || "",                 size: 20 },
      { text: (d.departmentName || "").toUpperCase(), size: 22, bold: true, underline: true },
    ]));
    children.push(blank());
    children.push(heading(d.collegeName));
    children.push(heading(d.departmentName));
    children.push(heading(`Camp Report – ${d.campLocation}`));
    children.push(heading(`Date: ${d.reportDateShort}`));
    children.push(blank());

    // Clean, natural paragraph — no stray ".state" or odd fragments
    children.push(normalText(
      `Department of Public Health Dentistry, ${d.collegeName}, Madurai in association with ${d.associationName} and ${d.projectName} conducted a dental screening and treatment camp at ${d.campLocation} on ${d.reportDateLong}.`
    ));
    children.push(normalText(
      `The camp was organised by Dr. R. Palanivel Pandian. It started at ${d.startTime} and concluded at ${d.endTime}. A team comprising ${d.staffCount} staff, ${d.postgraduateCount} postgraduate student${pgList.length !== 1 ? "s" : ""}, and ${d.internCount} intern${internList.length !== 1 ? "s" : ""} provided oral health care to the community.`
    ));
    children.push(normalText(
      `A total of ${d.totalPatients} people attended the dental camp, of whom ${d.treatmentCount} received treatment. All attendees were given oral health education and oral hygiene instructions.`
    ));

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ═══════════════════════════════════════════════════════════════════════════
    // PAGE 4 – PHOTOS  (two per row, spaced via a borderless 2-col table)
    // ═══════════════════════════════════════════════════════════════════════════
    children.push(makeLogoHeader([
      { text: (d.collegeName || "").toUpperCase(),    size: 26, bold: true, underline: true },
      { text: d.collegeAddress || "",                 size: 20 },
      { text: (d.departmentName || "").toUpperCase(), size: 22, bold: true, underline: true },
    ]));
    children.push(blank());
    children.push(heading("Photos"));
    children.push(blank());

    // Photo cell — no border, centred content, generous padding between photos
    const photoCell = (buf) =>
      new TableCell({
        borders: noBorders,
        width: { size: 4500, type: WidthType.DXA },
        margins: { top: 0, bottom: 0, left: 200, right: 200 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: buf
            ? [new ImageRun({ data: buf, transformation: { width: 270, height: 185 }, type: "png" })]
            : [new TextRun("")],
        })],
      });

    for (let i = 0; i < photos.length; i += 2) {
      const buf1 = photos[i]     ? fs.readFileSync(photos[i].path)     : null;
      const buf2 = photos[i + 1] ? fs.readFileSync(photos[i + 1].path) : null;

      children.push(new Table({
        alignment: AlignmentType.CENTER,
        width: { size: 9000, type: WidthType.DXA },
        columnWidths: [4500, 4500],
        borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideH: noBorder, insideV: noBorder },
        rows: [new TableRow({ children: [photoCell(buf1), photoCell(buf2)] })],
      }));

      // Spacer paragraph between photo rows
      children.push(new Paragraph({ text: "", spacing: { before: 200, after: 200 } }));
    }

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ═══════════════════════════════════════════════════════════════════════════
    // PAGE 5 – CAMP STATISTICS
    // ═══════════════════════════════════════════════════════════════════════════
    children.push(makeLogoHeader([
      { text: (d.collegeName || "").toUpperCase(),    size: 26, bold: true, underline: true },
      { text: d.collegeAddress || "",                 size: 20 },
      { text: (d.departmentName || "").toUpperCase(), size: 22, bold: true, underline: true },
    ]));
    children.push(blank());
    const cW = 2800;
    const campTable = new Table({
      alignment: AlignmentType.CENTER,
      width: { size: cW * 2, type: WidthType.DXA },
      columnWidths: [cW, cW],
      rows: [
        new TableRow({ children: [statsCell("Gender",  cW, true),  statsCell("No of Patients", cW, true)]  }),
        new TableRow({ children: [statsCell("Male",    cW, false), statsCell(d.maleCount,      cW, false)] }),
        new TableRow({ children: [statsCell("Female",  cW, false), statsCell(d.femaleCount,    cW, false)] }),
      ],
    });
    const campChart = await generatePieChart(
      ["Male", "Female"],
      [parseInt(d.maleCount) || 0, parseInt(d.femaleCount) || 0]
    );
    children.push(heading("Camp Statistics"));
    children.push(campTable);
    children.push(blank());
    children.push(centeredImage(campChart, 520, 360));
    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ═══════════════════════════════════════════════════════════════════════════
    // PAGE 6 – SCREENING STATISTICS
    // ═══════════════════════════════════════════════════════════════════════════
    children.push(makeLogoHeader([
      { text: (d.collegeName || "").toUpperCase(),    size: 26, bold: true, underline: true },
      { text: d.collegeAddress || "",                 size: 20 },
      { text: (d.departmentName || "").toUpperCase(), size: 22, bold: true, underline: true },
    ]));
    children.push(blank());
    let screeningRows = [
      ["Dental Caries", d.dentalCaries],
      ["Gingivitis",    d.gingivitis],
    ];
    if (d.extraScreening)
      JSON.parse(d.extraScreening).forEach((item) => screeningRows.push([item.name, item.value]));

    const sW = 3500;
    const screeningTable = new Table({
      alignment: AlignmentType.CENTER,
      width: { size: sW * 2, type: WidthType.DXA },
      columnWidths: [sW, sW],
      rows: [
        new TableRow({ children: [statsCell("Diagnosis", sW, true), statsCell("No of Patients", sW, true)] }),
        ...screeningRows.map((r) =>
          new TableRow({ children: [statsCell(r[0], sW, false), statsCell(r[1], sW, false)] })
        ),
      ],
    });
    const screeningChart = await generateChart(
      screeningRows.map((r) => r[0]),
      screeningRows.map((r) => parseInt(r[1]) || 0),
      "Diagnosis", "No of Patients"
    );
    children.push(heading("Screening Statistics"));
    children.push(screeningTable);
    children.push(blank());
    children.push(centeredImage(screeningChart, 560, 380));
    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ═══════════════════════════════════════════════════════════════════════════
    // PAGE 7 – TREATMENT STATISTICS
    // ═══════════════════════════════════════════════════════════════════════════
    children.push(makeLogoHeader([
      { text: (d.collegeName || "").toUpperCase(),    size: 26, bold: true, underline: true },
      { text: d.collegeAddress || "",                 size: 20 },
      { text: (d.departmentName || "").toUpperCase(), size: 22, bold: true, underline: true },
    ]));
    children.push(blank());
    let treatmentRows = [["Scaling", d.scaling || 0]];
    if (d.extraTreatment)
      JSON.parse(d.extraTreatment).forEach((item) => treatmentRows.push([item.name, item.value]));

    const tW = 3200;
    const treatmentTable = new Table({
      alignment: AlignmentType.CENTER,
      width: { size: tW * 2, type: WidthType.DXA },
      columnWidths: [tW, tW],
      rows: [
        new TableRow({ children: [statsCell("Treatment", tW, true), statsCell("No of Patients", tW, true)] }),
        ...treatmentRows.map((r) =>
          new TableRow({ children: [statsCell(r[0], tW, false), statsCell(r[1], tW, false)] })
        ),
      ],
    });
    const treatmentChart = await generateChart(
      treatmentRows.map((r) => r[0]),
      treatmentRows.map((r) => parseInt(r[1]) || 0),
      "Treatment", "No of Patients"
    );
    children.push(heading("Treatment Statistics"));
    children.push(treatmentTable);
    children.push(blank());
    children.push(centeredImage(treatmentChart, 560, 380));

    // ── FOOTER (pages 2-7 only — page 1 is plain circular) ───────────────────
    const footer = new Footer({
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({
          text: "HEAD OF THE DEPARTMENT                                      PRINCIPAL",
          font: "Times New Roman", size: 28, bold: true,
        })],
      })],
    });

    const doc = new Document({
      sections: [{
        footers: { default: footer },
        properties: {
          page: {
            size:   { width: 12240, height: 15840 },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
        },
        children,
      }],
    });

    const buffer   = await Packer.toBuffer(doc);
    const filename = "Camp_Report_" + Date.now() + ".docx";

    await db.query(
      "INSERT INTO reports(username, filename, file_data, created_date, created_time) VALUES($1,$2,$3,CURRENT_DATE,CURRENT_TIME)",
      [req.session.user, filename, buffer]
    );

    // Clean up uploaded photos
    photos.forEach((f) => { try { fs.unlinkSync(f.path); } catch (_) {} });

    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).send("Error generating report");
  }
};
