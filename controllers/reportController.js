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
WidthType
} = require("docx")

const axios = require("axios")
const fs = require("fs")
const path = require("path")
const db = require("../config/db")

// ================= CHART GENERATOR =================

async function generateChart(config){

const url = "https://quickchart.io/chart"

const response = await axios.post(
url,
{ chart: config },
{ responseType:"arraybuffer" }
)

return Buffer.from(response.data)

}

// ================= MAIN REPORT FUNCTION =================

exports.generateReport = async (req,res)=>{

try{

// check login session
if(!req.session.user){
return res.status(401).json({success:false})
}

const username = req.session.user
const d = req.body
const photos = req.files || []

const children = []

// ================= TEXT FUNCTIONS =================

const heading = (text)=>

new Paragraph({

alignment:AlignmentType.CENTER,

spacing:{ line:360 },

children:[

new TextRun({

text:(text||"").toUpperCase(),

font:"Times New Roman",

size:28,

bold:true,

underline:{ type:UnderlineType.SINGLE }

})

]

})

const normalText = (text,center=false)=>

new Paragraph({

alignment:center ? AlignmentType.CENTER : AlignmentType.LEFT,

spacing:{ line:480 },

children:[

new TextRun({

text:String(text),

font:"Times New Roman",

size:24

})

]

})

const blank = ()=> new Paragraph({ text:"", spacing:{ line:480 } })

// ================= PAGE 1 =================

children.push(heading(d.collegeName))
children.push(heading(d.departmentName))
children.push(heading(`Camp Report – ${d.campLocation}`))
children.push(heading(`Date: ${d.reportDateShort}`))
children.push(blank())

children.push(

normalText(

`Department of Public Health Dentistry, ${d.collegeName}, Madurai in association with ${d.associationName} and with ${d.projectName} conducted a dental screening and treatment camp at ${d.campLocation} on ${d.reportDateLong}.`

)

)

children.push(

normalText(

`Dr R. Palanivel Pandian organised this program. The Camp started at ${d.startTime} and ended at ${d.endTime}. A team of dentists including ${d.staffCount} staff member, ${d.postgraduateCount} postgraduate member and ${d.internCount} interns member provided oral health care to the people.`

)

)

children.push(

normalText(

`A total of ${d.totalPatients} people attended the dental camp and ${d.treatmentCount} people were treated along with oral health education and oral hygiene instructions.`

)

)

children.push(new Paragraph({ children:[new PageBreak()] }))

// ================= PAGE 2 PHOTOS =================

children.push(heading("Photos"))

for(let i=0;i<photos.length;i+=2){

const img1 = photos[i] ? fs.readFileSync(photos[i].path) : null
const img2 = photos[i+1] ? fs.readFileSync(photos[i+1].path) : null

children.push(

new Paragraph({

tabStops:[{ type:TabStopType.RIGHT, position:TabStopPosition.MAX }],

spacing:{ line:360 },

children:[

img1 ? new ImageRun({ data:img1, transformation:{ width:250,height:170 } }) : new TextRun(""),

new TextRun({ text:"\t" }),

img2 ? new ImageRun({ data:img2, transformation:{ width:250,height:170 } }) : new TextRun("")

]

})

)

children.push(blank())

}

children.push(new Paragraph({ children:[new PageBreak()] }))

// ================= CAMP STATISTICS =================

const campTable = new Table({

alignment:AlignmentType.CENTER,

width:{ size:60, type:WidthType.PERCENTAGE },

rows:[

new TableRow({

children:[

new TableCell({ children:[normalText("Gender",true)] }),

new TableCell({ children:[normalText("No of Patients",true)] })

]

}),

new TableRow({

children:[

new TableCell({ children:[normalText("Male",true)] }),

new TableCell({ children:[normalText(d.maleCount,true)] })

]

}),

new TableRow({

children:[

new TableCell({ children:[normalText("Female",true)] }),

new TableCell({ children:[normalText(d.femaleCount,true)] })

]

})

]

})

const campChart = await generateChart({

type:"bar",

data:{

labels:["Male","Female"],

datasets:[{

label:"Patients",

data:[parseInt(d.maleCount),parseInt(d.femaleCount)],

backgroundColor:"lightblue"

}]

}

})

children.push(heading("Camp Statistics"))
children.push(campTable)
children.push(blank())

children.push(

new Paragraph({

alignment:AlignmentType.CENTER,

children:[

new ImageRun({ data:campChart, transformation:{ width:500,height:300 } })

]

})

)

// ================= FOOTER =================

const footer = new Footer({

children:[

new Paragraph({

alignment:AlignmentType.CENTER,

children:[

new TextRun({

text:"HEAD OF THE DEPARTMENT                                      PRINCIPAL",

font:"Times New Roman",

size:28,

bold:true

})

]

})

]

})

// ================= CREATE DOCUMENT =================

const doc = new Document({

sections:[

{

footers:{ default:footer },

children

}

]

})

const buffer = await Packer.toBuffer(doc)

const filename = "Camp_Report_" + Date.now() + ".docx"

const reportPath = path.join("/tmp/", filename)

fs.writeFileSync(reportPath, buffer)

// ================= SAVE REPORT TO DATABASE =================

db.query(

"INSERT INTO reports(username,filename,created_date,created_time) VALUES(?,?,CURDATE(),CURTIME())",

[username,filename],

(err)=>{

if(err){

console.log("Database insert error:",err)

}

}

)

// ================= SEND FILE =================

res.setHeader("Content-Disposition","attachment; filename="+filename)

res.send(buffer)

}catch(err){

console.log(err)

res.status(500).send("Error generating report")

}

}
