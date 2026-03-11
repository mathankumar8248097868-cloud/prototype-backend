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

exports.generateReport = async (req, res) => {
  try {
    const d = req.body;
    const photos = req.files || [];
    const children = [];

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
          new TextRun({ text: String(text), font: "Times New Roman", size: 24 }),
        ],
      });

    const blank = () => new Paragraph({ text: "", spacing: { line: 480 } });

    // ── PAGE 1 ────────────────────────────────────────────────────────────────
    children.push(heading(d.collegeName));
    children.push(heading(d.departmentName));
    children.push(heading(`Camp Report – ${d.campLocation}`));
    children.push(heading(`Date: ${d.reportDateShort}`));
    children.push(blank());
    children.push(normalText(`Department of Public Health Dentistry, ${d.collegeName}, Madurai in association with ${d.associationName} and with ${d.projectName} conducted a dental screening and treatment camp at ${d.campLocation} on ${d.reportDateLong}.`));
    children.push(normalText(`Dr R. Palanivel Pandian organised this program. The Camp started at ${d.startTime} and ended at ${d.endTime}. A team of dentists including ${d.staffCount} staff member, ${d.postgraduateCount} postgraduate member and ${d.internCount} interns member provided oral health care to the people.`));
    children.push(normalText(`A total of ${d.totalPatients} people attended the dental camp and ${d.treatmentCount} people were treated along with oral health education and oral hygiene instructions.`));
    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ── PAGE 2 – PHOTOS ───────────────────────────────────────────────────────
    children.push(heading("Photos"));
    for (let i = 0; i < photos.length; i += 2) {
      const img1 = photos[i]     ? fs.readFileSync(photos[i].path)     : null;
      const img2 = photos[i + 1] ? fs.readFileSync(photos[i + 1].path) : null;
      children.push(
        new Paragraph({
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          spacing: { line: 360 },
          children: [
            img1 ? new ImageRun({ data: img1, transformation: { width: 250, height: 170 } }) : new TextRun(""),
            new TextRun("\t"),
            img2 ? new ImageRun({ data: img2, transformation: { width: 250, height: 170 } }) : new TextRun(""),
          ],
        })
      );
      children.push(blank());
    }
    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ── PAGE 3 – CAMP STATISTICS ──────────────────────────────────────────────
    const campTable = new Table({
      alignment: AlignmentType.CENTER,
      width: { size: 60, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [new TableCell({ children: [normalText("Gender", true)] }), new TableCell({ children: [normalText("No of Patients", true)] })] }),
        new TableRow({ children: [new TableCell({ children: [normalText("Male", true)] }), new TableCell({ children: [normalText(d.maleCount, true)] })] }),
        new TableRow({ children: [new TableCell({ children: [normalText("Female", true)] }), new TableCell({ children: [normalText(d.femaleCount, true)] })] }),
      ],
    });
    const campChart = await generateChart(["Male", "Female"], [parseInt(d.maleCount) || 0, parseInt(d.femaleCount) || 0], "Gender", "No of Patients");
    children.push(heading("Camp Statistics"));
    children.push(campTable);
    children.push(blank());
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: campChart, transformation: { width: 500, height: 320 } })] }));
    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ── PAGE 4 – SCREENING STATISTICS ────────────────────────────────────────
    let screeningRows = [["Dental Caries", d.dentalCaries], ["Gingivitis", d.gingivitis], ["Missing", d.missing]];
    if (d.extraScreening) JSON.parse(d.extraScreening).forEach((item) => screeningRows.push([item.name, item.value]));
    const screeningTable = new Table({
      alignment: AlignmentType.CENTER,
      width: { size: 70, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [new TableCell({ children: [normalText("Diagnosis", true)] }), new TableCell({ children: [normalText("No of Patients", true)] })] }),
        ...screeningRows.map((row) => new TableRow({ children: row.map((val) => new TableCell({ children: [normalText(val, true)] })) })),
      ],
    });
    const screeningChart = await generateChart(screeningRows.map((r) => r[0]), screeningRows.map((r) => parseInt(r[1]) || 0), "Diagnosis", "No of Patients");
    children.push(heading("Screening Statistics"));
    children.push(screeningTable);
    children.push(blank());
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: screeningChart, transformation: { width: 500, height: 320 } })] }));
    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ── PAGE 5 – TREATMENT STATISTICS ────────────────────────────────────────
    let treatmentRows = [["Scaling", d.scaling || 0]];
    if (d.extraTreatment) JSON.parse(d.extraTreatment).forEach((item) => treatmentRows.push([item.name, item.value]));
    const treatmentTable = new Table({
      alignment: AlignmentType.CENTER,
      width: { size: 60, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [new TableCell({ children: [normalText("Treatment", true)] }), new TableCell({ children: [normalText("No of Patients", true)] })] }),
        ...treatmentRows.map((row) => new TableRow({ children: row.map((val) => new TableCell({ children: [normalText(val, true)] })) })),
      ],
    });
    const treatmentChart = await generateChart(treatmentRows.map((r) => r[0]), treatmentRows.map((r) => parseInt(r[1]) || 0), "Treatment", "No of Patients");
    children.push(heading("Treatment Statistics"));
    children.push(treatmentTable);
    children.push(blank());
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: treatmentChart, transformation: { width: 500, height: 320 } })] }));

    // ── FOOTER ────────────────────────────────────────────────────────────────
    const footer = new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "HEAD OF THE DEPARTMENT                                      PRINCIPAL", font: "Times New Roman", size: 28, bold: true })],
        }),
      ],
    });

    const doc = new Document({ sections: [{ footers: { default: footer }, children }] });
    const buffer = await Packer.toBuffer(doc);
    const filename = "Camp_Report_" + Date.now() + ".docx";

    // ── SAVE BUFFER TO DATABASE (persists across restarts) ────────────────────
    await db.query(
      "INSERT INTO reports(username, filename, file_data, created_date, created_time) VALUES($1, $2, $3, CURRENT_DATE, CURRENT_TIME)",
      [req.session.user, filename, buffer]
    );

    // ── CLEAN UP UPLOADED PHOTOS ──────────────────────────────────────────────
    photos.forEach((photo) => {
      try { fs.unlinkSync(photo.path); } catch (e) {}
    });

    res.setHeader("Content-Disposition", "attachment; filename=" + filename);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.send(buffer);

  } catch (err) {
    console.log(err);
    res.status(500).send("Error generating report");
  }
};
