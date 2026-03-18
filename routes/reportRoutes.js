const express = require("express")
const router = express.Router()
const multer = require("multer")
const path = require("path")
const fs = require("fs")
const reportController = require("../controllers/reportController")

if (!fs.existsSync("uploads")) fs.mkdirSync("uploads")

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, "uploads/"),
  filename:    (req, file, cb) => cb(null, Date.now() + "_" + file.fieldname + path.extname(file.originalname))
})

const upload = multer({ storage })

// Only accept camp photos — logos come from server-side default files
const uploadFields = upload.fields([
  { name: "photos", maxCount: 30 },
])

router.post("/generate", (req, res, next) => {
  if (!req.session.user) return res.status(403).send("Not allowed")
  next()
}, uploadFields, reportController.generateReport)

module.exports = router