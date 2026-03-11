const express = require("express")
const cors = require("cors")
const path = require("path")
const session = require("express-session")
const fs = require("fs")

const reportRoutes = require("./routes/reportRoutes")
const db = require("./config/db")

const app = express()

app.set("trust proxy", 1)

app.use(cors({
  origin: ["https://mathankumar8248097868-cloud.github.io"],
  credentials: true
}))

app.use(express.json())

app.use(session({
  secret: "secretkey",
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: true,
    sameSite: "none",
    maxAge: 1000 * 60 * 60 * 24
  }
}))

// ================= USER LOGIN =================

app.post("/api/login", async (req, res) => {
  const { username, password } = req.body
  try {
    const result = await db.query(
      "SELECT * FROM users WHERE username=$1 AND password=$2",
      [username, password]
    )
    if (result.rows.length > 0) {
      req.session.user = username
      res.json({ success: true })
    } else {
      res.json({ success: false })
    }
  } catch (err) {
    console.log(err)
    res.json({ success: false })
  }
})

// ================= ADMIN LOGIN =================

app.post("/api/admin/login", (req, res) => {
  const { username, password } = req.body
  if (username === "ultra" && password === "ultra@123") {
    req.session.admin = true
    res.json({ success: true })
  } else {
    res.json({ success: false })
  }
})

// ================= SESSION CHECK — USER =================

app.get("/api/auth/check", (req, res) => {
  res.json({ loggedIn: !!req.session.user })
})

// ================= SESSION CHECK — ADMIN =================

app.get("/api/auth/admin/check", (req, res) => {
  res.json({ loggedIn: !!req.session.admin })
})

// ================= LOGOUT — USER =================

app.post("/api/logout", (req, res) => {
  req.session.destroy()
  res.json({ success: true })
})

// ================= LOGOUT — ADMIN =================

app.post("/api/admin/logout", (req, res) => {
  req.session.destroy()
  res.json({ success: true })
})

// ================= ADMIN CREATE USER =================

app.post("/api/admin/adduser", async (req, res) => {
  if (!req.session.admin) return res.status(403).json({ success: false })
  const { username, password } = req.body
  try {
    await db.query(
      "INSERT INTO users(username,password,created_date,created_time) VALUES($1,$2,CURRENT_DATE,CURRENT_TIME)",
      [username, password]
    )
    res.json({ success: true })
  } catch (err) {
    console.log(err)
    res.json({ success: false })
  }
})

// ================= ADMIN USER LIST =================

app.get("/api/admin/users", async (req, res) => {
  if (!req.session.admin) return res.status(403).send("Not allowed")
  try {
    const result = await db.query("SELECT * FROM users")
    res.json(result.rows)
  } catch (err) {
    console.log(err)
    res.json([])
  }
})

// ================= DELETE USER =================

app.delete("/api/admin/delete/:id", async (req, res) => {
  if (!req.session.admin) return res.status(403).send("Not allowed")
  try {
    await db.query("DELETE FROM users WHERE id=$1", [req.params.id])
    res.json({ success: true })
  } catch (err) {
    console.log(err)
    res.json({ success: false })
  }
})

// ================= EDIT USER =================

app.put("/api/admin/edit/:id", async (req, res) => {
  if (!req.session.admin) return res.status(403).send("Not allowed")
  const { username, password } = req.body
  try {
    await db.query(
      "UPDATE users SET username=$1, password=$2 WHERE id=$3",
      [username, password, req.params.id]
    )
    res.json({ success: true })
  } catch (err) {
    console.log(err)
    res.json({ success: false })
  }
})

// ================= ADMIN REPORT LIST =================

app.get("/api/admin/reports", async (req, res) => {
  if (!req.session.admin) return res.status(403).send("Not allowed")
  try {
    const result = await db.query(
      "SELECT id, username, filename, created_date, created_time FROM reports"
    )
    res.json(result.rows)
  } catch (err) {
    console.log(err)
    res.json([])
  }
})

// ================= DOWNLOAD REPORT FROM DATABASE =================

app.get("/api/admin/reports/download", async (req, res) => {
  if (!req.session.admin) return res.status(403).send("Not allowed")
  const { filename } = req.query
  if (!filename) return res.status(400).send("Filename required")
  try {
    const result = await db.query(
      "SELECT filename, file_data FROM reports WHERE filename=$1",
      [filename]
    )
    if (result.rows.length === 0 || !result.rows[0].file_data) {
      return res.status(404).send("File not found in database.")
    }
    const { file_data } = result.rows[0]
    const safeName = path.basename(filename)
    res.setHeader("Content-Disposition", `attachment; filename="${safeName}"`)
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    res.send(file_data)
  } catch (err) {
    console.log(err)
    res.status(500).send("Error retrieving file")
  }
})

// ================= DELETE REPORT =================

app.delete("/api/admin/reports/delete", async (req, res) => {
  if (!req.session.admin) return res.status(403).send("Not allowed")
  const { filename } = req.query
  if (!filename) return res.status(400).json({ success: false, message: "Filename required" })
  try {
    await db.query("DELETE FROM reports WHERE filename=$1", [filename])
    res.json({ success: true })
  } catch (err) {
    console.log(err)
    res.json({ success: false })
  }
})

// ================= REPORT API =================

app.use("/api/report", reportRoutes)

// ================= PROTECTED PAGES (served by backend only) =================
// The adminpage.html and report.html files must be placed in backend/protected/
// They are NEVER served as static files — only through these guarded routes.
// So even if someone knows the URL, the server returns a redirect, not the page.

app.get("/page/admin", (req, res) => {
  if (!req.session.admin) {
    // Not logged in — redirect to login page on GitHub Pages
    return res.redirect("https://mathankumar8248097868-cloud.github.io/prototype-frontend/adminlogin.html")
  }
  res.sendFile(path.join(__dirname, "protected/adminpage.html"))
})

app.get("/page/report", (req, res) => {
  if (!req.session.user) {
    // Not logged in — redirect to login page on GitHub Pages
    return res.redirect("https://mathankumar8248097868-cloud.github.io/prototype-frontend/userlogin.html")
  }
  res.sendFile(path.join(__dirname, "protected/report.html"))
})

// ================= STATIC FILES (public pages only) =================

app.use(express.static(path.join(__dirname, "../frontend"), { index: false }))

// ================= START SERVER =================

const PORT = process.env.PORT || 5000

app.listen(PORT, () => {
  console.log("Server running on port " + PORT)
})
