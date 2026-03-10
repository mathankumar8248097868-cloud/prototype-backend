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

// QuickChart generator
async function generateChart(config) {
  const response = await axios.post(
    "https://quickchart.io/chart",
    { chart: config },
    { responseType: "arraybuffer" }
  );
  return Buffer.from(response.data);
}

exports.generateReport = async (req, res) => {
  try {

    const d = req.body;
    const photos = req.files || [];
    const children = [];

    // ===== Helper Functions =====

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
          new TextRun({
            text: String(text || ""),
            font: "Times New Roman",
            size: 24,
          }),
        ],
      });

    const blank = () =>
      new Paragraph({
        text: "",
        spacing: { line: 480 },
      });

    // ===== PAGE 1 =====

    children.push(heading(d.collegeName));
    children.push(heading(d.departmentName));
    children.push(heading(`Camp Report – ${d.campLocation}`));
    children.push(heading(`Date: ${d.reportDateShort}`));
    children.push(blank());

    children.push(
      normalText(
        `Department of Public Health Dentistry, ${d.collegeName}, Madurai in association with ${d.associationName} and with ${d.projectName} conducted a dental screening and treatment camp at ${d.campLocation} on ${d.reportDateLong}.`
      )
    );

    children.push(
      normalText(
        `Dr R. Palanivel Pandian organised this program. The Camp started at ${d.startTime} and ended at ${d.endTime}. A team of dentists including ${d.staffCount} staff member, ${d.postgraduateCount} postgraduate member and ${d.internCount} interns member provided oral health care to the people.`
      )
    );

    children.push(
      normalText(
        `A total of ${d.totalPatients} people attended the dental camp and ${d.treatmentCount} people were treated along with oral health education and oral hygiene instructions.`
      )
    );

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ===== PAGE 2 PHOTOS =====

    children.push(heading("Photos"));

    for (let i = 0; i < photos.length; i += 2) {

      const img1 = photos[i] ? fs.readFileSync(photos[i].path) : null;
      const img2 = photos[i + 1] ? fs.readFileSync(photos[i + 1].path) : null;

      children.push(
        new Paragraph({
          tabStops: [
            {
              type: TabStopType.RIGHT,
              position: TabStopPosition.MAX,
            },
          ],
          spacing: { line: 360 },
          children: [
            img1
              ? new ImageRun({
                  data: img1,
                  transformation: { width: 250, height: 170 },
                })
              : new TextRun(""),

            new TextRun({ text: "\t" }),

            img2
              ? new ImageRun({
                  data: img2,
                  transformation: { width: 250, height: 170 },
                })
              : new TextRun(""),
          ],
        })
      );

      children.push(blank());
    }

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ===== PAGE 3 CAMP STATISTICS =====

    const campChart = await generateChart({
      type: "bar",
      data: {
        labels: ["Male", "Female"],
        datasets: [
          {
            data: [
              parseInt(d.maleCount || 0),
              parseInt(d.femaleCount || 0),
            ],
            backgroundColor: "lightblue",
          },
        ],
      },
      options: {
        plugins: { legend: { display: false } },
        scales: {
          x: { title: { display: true, text: "Gender" } },
          y: { title: { display: true, text: "No of Patients" } },
        },
      },
    });

    children.push(heading("Camp Statistics"));

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new ImageRun({
            data: campChart,
            transformation: { width: 500, height: 300 },
          }),
        ],
      })
    );

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ===== PAGE 4 SCREENING =====

    let screeningDataRows = [
      ["Dental Caries", d.dentalCaries],
      ["Gingivitis", d.gingivitis],
      ["Missing", d.missing],
    ];

    if (d.extraScreening) {
      const extra = JSON.parse(d.extraScreening);
      extra.forEach((item) =>
        screeningDataRows.push([item.name, item.value])
      );
    }

    const screeningChart = await generateChart({
      type: "bar",
      data: {
        labels: screeningDataRows.map((r) => r[0]),
        datasets: [
          {
            data: screeningDataRows.map((r) =>
              parseInt(r[1] || 0)
            ),
            backgroundColor: "lightblue",
          },
        ],
      },
      options: {
        plugins: { legend: { display: false } },
        scales: {
          x: { title: { display: true, text: "Diagnosis" } },
          y: { title: { display: true, text: "No of Patients" } },
        },
      },
    });

    children.push(heading("Screening Statistics"));

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new ImageRun({
            data: screeningChart,
            transformation: { width: 500, height: 300 },
          }),
        ],
      })
    );

    children.push(new Paragraph({ children: [new PageBreak()] }));

    // ===== PAGE 5 TREATMENT =====

    let treatmentRows = [["Scaling", d.scaling || 0]];

    if (d.extraTreatment) {
      const extra = JSON.parse(d.extraTreatment);
      extra.forEach((item) =>
        treatmentRows.push([item.name, item.value])
      );
    }

    const treatmentChart = await generateChart({
      type: "bar",
      data: {
        labels: treatmentRows.map((r) => r[0]),
        datasets: [
          {
            data: treatmentRows.map((r) =>
              parseInt(r[1] || 0)
            ),
            backgroundColor: "lightblue",
          },
        ],
      },
      options: {
        plugins: { legend: { display: false } },
        scales: {
          x: { title: { display: true, text: "Treatment" } },
          y: { title: { display: true, text: "No of Patients" } },
        },
      },
    });

    children.push(heading("Treatment Statistics"));

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new ImageRun({
            data: treatmentChart,
            transformation: { width: 500, height: 300 },
          }),
        ],
      })
    );

    // ===== FOOTER =====

    const footer = new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text:
                "HEAD OF THE DEPARTMENT                                      PRINCIPAL",
              font: "Times New Roman",
              size: 28,
              bold: true,
            }),
          ],
        }),
      ],
    });

    // ===== DOCUMENT BUILD =====

    const doc = new Document({
      sections: [
        {
          footers: { default: footer },
          children,
        },
      ],
    });

    const buffer = await Packer.toBuffer(doc);

    const filename = "Camp_Report_" + Date.now() + ".docx";
    const reportPath = path.join("/tmp/", filename);

    fs.writeFileSync(reportPath, buffer);

    await db.query(
      "INSERT INTO reports(username, filename, created_date, created_time) VALUES($1,$2,CURRENT_DATE,CURRENT_TIME)",
      [req.session.user, filename]
    );

    res.setHeader(
      "Content-Disposition",
      "attachment; filename=" + filename
    );

    res.send(buffer);

  } catch (err) {

    console.log(err);
    res.status(500).send("Error generating report");

  }
};
