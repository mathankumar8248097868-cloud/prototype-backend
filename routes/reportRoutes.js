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
    cb(null, Date.now() + path.extname(file.originalname))
  }
})

const upload = multer({ storage: storage })

// ── Protect route — only logged-in users can generate reports ─────────────────
router.post("/generate", (req, res, next) => {
  if (!req.session.user) return res.status(403).send("Not allowed")
  next()
}, upload.array("photos"), reportController.generateReport)

module.exports = router