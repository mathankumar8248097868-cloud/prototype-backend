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
const fs = require("fs");
const path = require("path");
const db = require("../config/db");

// ── Generate chart using QuickChart (Chart.js v2 syntax) ──────────────────────
async function generateChart(labels, data, xLabel, yLabel) {
  const config = {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          data,
          backgroundColor: "lightblue",
          borderColor: "steelblue",
          borderWidth: 1,
        },
      ],
    },
    options: {
      legend: { display: false },
      scales: {
        xAxes: [{ scaleLabel: { display: true, labelString: xLabel, fontSize: 14, fontStyle: "bold" } }],
        yAxes: [{ scaleLabel: { display: true, labelString: yLabel, fontSize: 14, fontStyle: "bold" }, ticks: { beginAtZero: true } }],
      },
    },
  };

  const response = await axios.post(
    "https://quickchart.io/chart",
    { chart: config, width: 600, height: 400, backgroundColor: "white" },
    { responseType: "arraybuffer" }
  );
  return Buffer.from(response.data);
}

async function generatePieChart(labels, data) {
  const config = {
    type: "pie",
    data: {
      labels,
      datasets: [
        {
          data,
          backgroundColor: ["lightblue", "lightpink"],
          borderColor: ["steelblue", "hotpink"],
          borderWidth: 1,
        },
      ],
    },
    options: {
      legend: { display: true },
    },
  };

  const response = await axios.post(
    "https://quickchart.io/chart",
    { chart: config, width: 600, height: 400, backgroundColor: "white" },
    { responseType: "arraybuffer" }
  );
  return Buffer.from(response.data);
}

// ── Helper: thin border object ────────────────────────────────────────────────
const thinBorder = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
const cellBorders = {
  top: thinBorder,
  bottom: thinBorder,
  left: thinBorder,
  right: thinBorder,
};

exports.generateReport = async (req, res) => {
  try {
    const d = req.body;
    const photos = [];
    let logoFile = null;

    // Separate logo from camp photos
    if (req.files) {
      for (const f of req.files) {
        if (f.fieldname === "logo") {
          logoFile = f;
        } else {
          photos.push(f);
        }
      }
    }

    const children = [];

    // ── Paragraph helpers ─────────────────────────────────────────────────────
    const heading = (text) =>
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: 360 },
        children: [
          new TextRun({
            text: (text || "").toUpperCase(),
            font: "Times New Roman",
            size: 28,
            bold: true,
            underline: { type: UnderlineType.SINGLE },
          }),
        ],
      });

    const normalText = (text, center = false) =>
      new Paragraph({
        alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
        spacing: { line: 480 },
        children: [
          new TextRun({ text: String(text ?? ""), font: "Times New Roman", size: 24 }),
        ],
      });

    const boldText = (text, center = false) =>
      new Paragraph({
        alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
        spacing: { line: 360 },
        children: [
          new TextRun({ text: String(text ?? ""), font: "Times New Roman", size: 24, bold: true }),
        ],
      });

    const blank = () => new Paragraph({ text: "", spacing: { line: 480 } });

    // ── Signature table cell helper ───────────────────────────────────────────
    const sigCell = (text, widthDxa) =>
      new TableCell({
        borders: cellBorders,
        width: { size: widthDxa, type: WidthType.DXA },
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: String(text ?? ""), font: "Times New Roman", size: 22 })],
          }),
        ],
      });

    const sigHeaderCell = (text, widthDxa) =>
      new TableCell({
        borders: cellBorders,
        width: { size: widthDxa, type: WidthType.DXA },
        shading: { fill: "E8E8E8", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: String(text ?? ""), font: "Times New Roman", size: 22, bold: true })],
          }),
        ],
      });

    // Parse staff/pg/intern lists
    const staffList   = d.staffList   ? JSON.parse(d.staffList)   : [];
    const pgList      = d.pgList      ? JSON.parse(d.pgList)      : [];
    const internList  = d.internList  ? JSON.parse(d.internList)  : [];

    // ── Logo image buffer ─────────────────────────────────────────────────────
    let logoBuffer = null;
    if (logoFile) {
      logoBuffer = fs.readFileSync(logoFile.path);
    }

    // ── Helper: logo row paragraph (left logo | center heading | right logo) ──
    const makeLogoHeaderRows = (centerLines) => {
      const logoRun = (buf) =>
        buf
          ? new ImageRun({ data: buf, transformation: { width: 70, height: 70 }, type: "png" })
          : new TextRun("");

      // Build a 3-column table: left logo | center text | right logo
      const centerChildren = centerLines.map(
        (line) =>
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { line: 280 },
            children: [
              new TextRun({
                text: line.text || "",
                font: "Times New Roman",
                size: line.size || 22,
                bold: !!line.bold,
                underline: line.underline ? { type: UnderlineType.SINGLE } : undefined,
              }),
            ],
          })
      );

      return new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1200, 6960, 1200],
        borders: {
          top: { style: BorderStyle.NONE },
          bottom: { style: BorderStyle.NONE },
          left: { style: BorderStyle.NONE },
          right: { style: BorderStyle.NONE },
          insideH: { style: BorderStyle.NONE },
          insideV: { style: BorderStyle.NONE },
        },
        rows: [
          new TableRow({
            children: [
              // Left logo
              new TableCell({
                borders: {
                  top: { style: BorderStyle.NONE },
                  bottom: { style: BorderStyle.NONE },
                  left: { style: BorderStyle.NONE },
                  right: { style: BorderStyle.NONE },
                },
                width: { size: 1200, type: WidthType.DXA },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [logoBuffer ? logoRun(logoBuffer) : new TextRun("")],
                  }),
                ],
              }),
              // Center text
              new TableCell({
                borders: {
                  top: { style: BorderStyle.NONE },
                  bottom: { style: BorderStyle.NONE },
                  left: { style: BorderStyle.NONE },
                  right: { style: BorderStyle.NONE },
                },
                width: { size: 6960, type: WidthType.DXA },
                verticalAlign: VerticalAlign.CENTER,
                children: centerChildren,
              }),
              // Right logo
              new TableCell({
                borders: {
                  top: { style: BorderStyle.NONE },
                  bottom: { style: BorderStyle.NONE },
                  left: { style: BorderStyle.NONE },
                  right: { style: BorderStyle.NONE },
                },
                width: { size: 1200, type: WidthType.DXA },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [logoBuffer ? logoRun(logoBuffer) : new TextRun("")],
                  }),
                ],
              }),
            ],
          }),
        ],
      });
    };

    // ═══════════════════════════════════════════════════════════════════════════
    // PAGE 1 — CAMP CIRCULAR (like Ref No document)
    // ═══════════════════════════════════════════════════════════════════════════

    // Ref No & Date line
    children.push(
      new Paragraph({
        spacing: { line: 360 },
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        children: [
          new TextRun({ text: `Ref No: ${d.refNo || ""}`, font: "Times New Roman", size: 22 }),
          new TextRun({ text: "\t" }),
          new TextRun({ text: `Date: ${d.reportDateShort || ""}`, font: "Times New Roman", size: 22 }),
        ],
      })
    );
    children.push(blank());

    // Department heading
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: 360 },
        children: [
          new TextRun({ text: d.departmentName || "", font: "Times New Roman", size: 28, bold: true, underline: { type: UnderlineType.SINGLE } }),
        ],
      })
    );

    // Camp Circular heading
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: 360 },
        children: [
          new TextRun({ text: "CAMP CIRCULAR", font: "Times New Roman", size: 26, bold: true }),
        ],
      })
    );

    children.push(blank());

    // Description paragraph
    children.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 480 },
        children: [
          new TextRun({
            text: `A dental screening and treatment camp will be conducted at ${d.campLocation || ""} on ${d.reportDateLong || ""} from ${d.startTime || ""} to ${d.endTime || ""}. The bus will depart from the campus at ${d.busTime || d.startTime || ""} on ${d.reportDateLong || ""}.`,
            font: "Times New Roman",
            size: 24,
          }),
        ],
      })
    );

    children.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 480 },
        children: [
          new TextRun({
            text: "The following Staff, PG students, and interns are posted for the above camp.",
            font: "Times New Roman",
            size: 24,
          }),
        ],
      })
    );

    children.push(blank());

    // STAFF section
    children.push(boldText("STAFF:"));
    if (staffList.length > 0) {
      staffList.forEach((name) => children.push(normalText(name)));
    } else {
      children.push(normalText(""));
    }

    children.push(blank());

    // POSTGRADUATE section
    children.push(boldText("POSTGRADUATE:"));
    if (pgList.length > 0) {
      pgList.forEach((name) => children.push(normalText(name)));
    } else {
      children.push(normalText(""));
    }

    children.push(blank());

    // INTERNS section
    children.push(boldText("INTERNS:"));
    if (internList.length > 0) {
      internList.forEach((name) => children.push(normalText(name)));
    } else {
      children.push(normalText(""));
    }

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ═══════════════════════════════════════════════════════════════════════════
    // PAGE 2 — ATTENDANCE SHEET (like atteddd document)
    // ═══════════════════════════════════════════════════════════════════════════

    // Logo header with college name + department
    children.push(
      makeLogoHeaderRows([
        { text: (d.collegeName || "").toUpperCase(), size: 26, bold: true, underline: true },
        { text: d.collegeAddress || "", size: 20 },
        { text: (d.departmentName || "").toUpperCase(), size: 22, bold: true, underline: true },
      ])
    );

    children.push(blank());

    // Camp Site & Date
    children.push(
      new Paragraph({
        spacing: { line: 360 },
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        children: [
          new TextRun({ text: `CAMP SITE: ${d.campLocation || ""}`, font: "Times New Roman", size: 22, bold: true }),
          new TextRun({ text: "\t" }),
          new TextRun({ text: `DATE: ${d.reportDateShort || ""}`, font: "Times New Roman", size: 22, bold: true }),
        ],
      })
    );

    children.push(blank());

    // STAFFS signature table
    const makeSignatureTable = (headerLabel, nameList) => {
      const nameWidth = 5580;
      const sigWidth = 3780;
      const rows = [
        new TableRow({
          children: [
            sigHeaderCell(headerLabel, nameWidth),
            sigHeaderCell("SIGNATURE", sigWidth),
          ],
        }),
        ...nameList.map(
          (name) =>
            new TableRow({
              children: [
                sigCell(name, nameWidth),
                sigCell("", sigWidth),
              ],
            })
        ),
      ];
      // Always at least one empty row if no names
      if (nameList.length === 0) {
        rows.push(
          new TableRow({
            children: [sigCell("", nameWidth), sigCell("", sigWidth)],
          })
        );
      }
      return new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [nameWidth, sigWidth],
        rows,
      });
    };

    children.push(makeSignatureTable("STAFFS", staffList));
    children.push(blank());
    children.push(makeSignatureTable("POST GRADUATE", pgList));
    children.push(blank());
    children.push(makeSignatureTable("INTERNS", internList));

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ═══════════════════════════════════════════════════════════════════════════
    // PAGE 3 — ORIGINAL PAGE 1 (Camp Report summary)
    // ═══════════════════════════════════════════════════════════════════════════
    children.push(heading(d.collegeName));
    children.push(heading(d.departmentName));
    children.push(heading(`Camp Report – ${d.campLocation}`));
    children.push(heading(`Date: ${d.reportDateShort}`));
    children.push(blank());
    children.push(normalText(`Department of Public Health Dentistry, ${d.collegeName}, Madurai in association with ${d.associationName} and with ${d.projectName} conducted a dental screening and treatment camp at ${d.campLocation} on ${d.reportDateLong}.`));
    children.push(normalText(`Dr R. Palanivel Pandian organised this program. The Camp started at ${d.startTime} and ended at ${d.endTime}. A team of dentists including ${d.staffCount} staff member, ${d.postgraduateCount} postgraduate member and ${d.internCount} interns member provided oral health care to the people.`));
    children.push(normalText(`A total of ${d.totalPatients} people attended the dental camp and ${d.treatmentCount} people were treated along with oral health education and oral hygiene instructions.`));
    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ═══════════════════════════════════════════════════════════════════════════
    // PAGE 4 – PHOTOS
    // ═══════════════════════════════════════════════════════════════════════════
    children.push(heading("Photos"));
    for (let i = 0; i < photos.length; i += 2) {
      const img1 = photos[i]     ? fs.readFileSync(photos[i].path)     : null;
      const img2 = photos[i + 1] ? fs.readFileSync(photos[i + 1].path) : null;
      children.push(
        new Paragraph({
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          spacing: { line: 360 },
          children: [
            img1 ? new ImageRun({ data: img1, transformation: { width: 250, height: 170 }, type: "png" }) : new TextRun(""),
            new TextRun("\t"),
            img2 ? new ImageRun({ data: img2, transformation: { width: 250, height: 170 }, type: "png" }) : new TextRun(""),
          ],
        })
      );
      children.push(blank());
    }
    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ═══════════════════════════════════════════════════════════════════════════
    // PAGE 5 – CAMP STATISTICS
    // ═══════════════════════════════════════════════════════════════════════════
    const campTable = new Table({
      width: { size: 4680, type: WidthType.DXA },
      columnWidths: [2340, 2340],
      rows: [
        new TableRow({ children: [new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA }, children: [normalText("Gender", true)] }), new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA }, children: [normalText("No of Patients", true)] })] }),
        new TableRow({ children: [new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA }, children: [normalText("Male", true)] }), new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA }, children: [normalText(d.maleCount, true)] })] }),
        new TableRow({ children: [new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA }, children: [normalText("Female", true)] }), new TableCell({ borders: cellBorders, width: { size: 2340, type: WidthType.DXA }, children: [normalText(d.femaleCount, true)] })] }),
      ],
    });
    const campChart = await generatePieChart(["Male", "Female"], [parseInt(d.maleCount) || 0, parseInt(d.femaleCount) || 0]);
    children.push(heading("Camp Statistics"));
    children.push(campTable);
    children.push(blank());
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: campChart, transformation: { width: 500, height: 320 }, type: "png" })] }));
    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ═══════════════════════════════════════════════════════════════════════════
    // PAGE 6 – SCREENING STATISTICS
    // ═══════════════════════════════════════════════════════════════════════════
    let screeningRows = [["Dental Caries", d.dentalCaries], ["Gingivitis", d.gingivitis], ["Missing", d.missing]];
    if (d.extraScreening) JSON.parse(d.extraScreening).forEach((item) => screeningRows.push([item.name, item.value]));
    const col1 = 3500, col2 = 3500;
    const screeningTable = new Table({
      width: { size: col1 + col2, type: WidthType.DXA },
      columnWidths: [col1, col2],
      rows: [
        new TableRow({ children: [new TableCell({ borders: cellBorders, width: { size: col1, type: WidthType.DXA }, children: [normalText("Diagnosis", true)] }), new TableCell({ borders: cellBorders, width: { size: col2, type: WidthType.DXA }, children: [normalText("No of Patients", true)] })] }),
        ...screeningRows.map((row) => new TableRow({ children: row.map((val) => new TableCell({ borders: cellBorders, width: { size: col1, type: WidthType.DXA }, children: [normalText(val, true)] })) })),
      ],
    });
    const screeningChart = await generateChart(screeningRows.map((r) => r[0]), screeningRows.map((r) => parseInt(r[1]) || 0), "Diagnosis", "No of Patients");
    children.push(heading("Screening Statistics"));
    children.push(screeningTable);
    children.push(blank());
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: screeningChart, transformation: { width: 500, height: 320 }, type: "png" })] }));
    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ═══════════════════════════════════════════════════════════════════════════
    // PAGE 7 – TREATMENT STATISTICS
    // ═══════════════════════════════════════════════════════════════════════════
    let treatmentRows = [["Scaling", d.scaling || 0]];
    if (d.extraTreatment) JSON.parse(d.extraTreatment).forEach((item) => treatmentRows.push([item.name, item.value]));
    const t1 = 3000, t2 = 3000;
    const treatmentTable = new Table({
      width: { size: t1 + t2, type: WidthType.DXA },
      columnWidths: [t1, t2],
      rows: [
        new TableRow({ children: [new TableCell({ borders: cellBorders, width: { size: t1, type: WidthType.DXA }, children: [normalText("Treatment", true)] }), new TableCell({ borders: cellBorders, width: { size: t2, type: WidthType.DXA }, children: [normalText("No of Patients", true)] })] }),
        ...treatmentRows.map((row) => new TableRow({ children: row.map((val) => new TableCell({ borders: cellBorders, width: { size: t1, type: WidthType.DXA }, children: [normalText(val, true)] })) })),
      ],
    });
    const treatmentChart = await generateChart(treatmentRows.map((r) => r[0]), treatmentRows.map((r) => parseInt(r[1]) || 0), "Treatment", "No of Patients");
    children.push(heading("Treatment Statistics"));
    children.push(treatmentTable);
    children.push(blank());
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: treatmentChart, transformation: { width: 500, height: 320 }, type: "png" })] }));

    // ── FOOTER ────────────────────────────────────────────────────────────────
    const footer = new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "HEAD OF THE DEPARTMENT                                      PRINCIPAL", font: "Times New Roman", size: 28, bold: true })],
        }),
      ],
    });

    const doc = new Document({
      sections: [{
        footers: { default: footer },
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
        },
        children,
      }],
    });
    const buffer = await Packer.toBuffer(doc);
    const filename = "Camp_Report_" + Date.now() + ".docx";

    // ── SAVE TO DATABASE ──────────────────────────────────────────────────────
    await db.query(
      "INSERT INTO reports(username, filename, file_data, created_date, created_time) VALUES($1, $2, $3, CURRENT_DATE, CURRENT_TIME)",
      [req.session.user, filename, buffer]
    );

    // ── CLEAN UP UPLOADED FILES ───────────────────────────────────────────────
    [...photos, logoFile].filter(Boolean).forEach((f) => {
      try { fs.unlinkSync(f.path); } catch (e) {}
    });

    res.setHeader("Content-Disposition", "attachment; filename=" + filename);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.send(buffer);

  } catch (err) {
    console.log(err);
    res.status(500).send("Error generating report");
  }
};