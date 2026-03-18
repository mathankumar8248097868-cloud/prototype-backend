const express = require("express")
const router = express.Router()
const multer = require("multer")
const path = require("path")
const fs = require("fs")
const reportController = require("../controllers/reportController")

// ── Create uploads folder if it doesn't exist ────────────────────────────────
if (!fs.existsSync("uploads")) {
  fs.mkdirSync("uploads")
}

// ── Multer storage config ─────────────────────────────────────────────────────
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, "uploads/")
  },
  filename: function (req, file, cb) {
    cb(null, Date.now() + "_" + file.fieldname + path.extname(file.originalname))
  }
})

const upload = multer({ storage: storage })

// Accept: logo (single) + photos (multiple)
const uploadFields = upload.fields([
  { name: "logo", maxCount: 1 },
  { name: "photos", maxCount: 30 },
])

// ── Protect route — only logged-in users can generate reports ─────────────────
router.post("/generate", (req, res, next) => {
  if (!req.session.user) return res.status(403).send("Not allowed")
  next()
}, uploadFields, reportController.generateReport)

module.exports = router