const express = require("express")
const router = express.Router()

const multer = require("multer")
const path = require("path")

const reportController = require("../controllers/reportController")

// ================= FILE STORAGE =================

const storage = multer.diskStorage({

destination: function (req, file, cb) {

cb(null, "uploads/")

},

filename: function (req, file, cb) {

cb(null, Date.now() + path.extname(file.originalname))

}

})

const upload = multer({ storage })

// ================= GENERATE REPORT =================

router.post("/generate", upload.array("photos"), (req,res,next)=>{

// check login
if(!req.session.user){
return res.status(401).json({success:false, message:"Not logged in"})
}

// attach username to request
req.username = req.session.user

next()

}, reportController.generateReport)

module.exports = router
