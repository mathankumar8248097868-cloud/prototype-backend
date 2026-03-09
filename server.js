const express = require("express")
const cors = require("cors")
const path = require("path")
const session = require("express-session")

const reportRoutes = require("./routes/reportRoutes")
const db = require("./config/db")

const app = express();

app.set("trust proxy", 1);

app.use(cors({
  origin:["https://mathankumar8248097868-cloud.github.io"],
  credentials:true
}));

app.use(express.json());

app.use(session({
  secret:"secretkey",
  resave:false,
  saveUninitialized:false,
  cookie:{
    secure:true,
    sameSite:"none",
    maxAge:1000*60*60*24
  }
}));

// ================= USER LOGIN =================

// ================= USER LOGIN =================

app.post("/api/login",(req,res)=>{

const {username,password} = req.body

const sql = "SELECT * FROM users WHERE username=? AND password=?"

db.query(sql,[username,password],(err,result)=>{

if(err){
console.log(err)
return res.json({success:false})
}

if(result.length > 0){

req.session.user = username

res.json({success:true})

}else{

res.json({success:false})

}

})

})

// ================= ADMIN LOGIN =================

app.post("/api/admin/login",(req,res)=>{

const {username,password} = req.body

if(username==="ultra" && password==="ultra@123"){

req.session.admin = true
res.json({success:true})

}else{

res.json({success:false})

}

})


// ================= ADMIN CREATE USER =================

app.post("/api/admin/adduser",(req,res)=>{

if(!req.session.admin){
return res.status(403).json({success:false})
}

const {username,password} = req.body

const sql = `
INSERT INTO users(username,password,created_date,created_time)
VALUES(?,?,CURDATE(),CURTIME())
`

db.query(sql,[username,password],(err,result)=>{

if(err){
console.log(err)
return res.json({success:false})
}

res.json({success:true})

})

})


// ================= ADMIN USER LIST =================

app.get("/api/admin/users",(req,res)=>{

if(!req.session.admin){
return res.status(403).send("Not allowed")
}

db.query("SELECT * FROM users",(err,result)=>{

if(err){
console.log(err)
return res.json([])
}

res.json(result)

})

})


// ================= DELETE USER =================

app.delete("/api/admin/delete/:id",(req,res)=>{

if(!req.session.admin){
return res.status(403).send("Not allowed")
}

db.query("DELETE FROM users WHERE id=?",[req.params.id],(err)=>{

if(err){
console.log(err)
return res.json({success:false})
}

res.json({success:true})

})

})


// ================= EDIT USER =================

app.put("/api/admin/edit/:id",(req,res)=>{

if(!req.session.admin){
return res.status(403).send("Not allowed")
}

const {username,password} = req.body

const sql = "UPDATE users SET username=?, password=? WHERE id=?"

db.query(sql,[username,password,req.params.id],(err)=>{

if(err){
console.log(err)
return res.json({success:false})
}

res.json({success:true})

})

})


// ================= ADMIN REPORT LIST =================

app.get("/api/admin/reports",(req,res)=>{

if(!req.session.admin){
return res.status(403).send("Not allowed")
}

db.query("SELECT * FROM reports",(err,result)=>{

if(err){
console.log(err)
return res.json([])
}

res.json(result)

})

})


// ================= PROTECT REPORT PAGE =================

app.get("/report",(req,res)=>{

if(!req.session.user){
return res.redirect("/userlogin.html")
}

res.sendFile(path.join(__dirname,"../frontend/report.html"))

})


// ================= PROTECT ADMIN PAGE =================

app.get("/admin",(req,res)=>{

if(!req.session.admin){
return res.redirect("/adminlogin.html")
}

res.sendFile(path.join(__dirname,"../frontend/adminpage.html"))

})


// ================= REPORT API =================

app.use("/api/report", reportRoutes)


// ================= HOME PAGE =================

//app.get("/", (req,res)=>{

//res.sendFile(path.join(__dirname,"../frontend/index.html"))

//})


// ================= BLOCK DIRECT ACCESS =================

app.use((req,res,next)=>{

if(req.path === "/report.html"){
return res.redirect("/userlogin.html")
}

if(req.path === "/adminpage.html"){
return res.redirect("/adminlogin.html")
}

next()

})


// ================= STATIC FILES =================

app.use(express.static(path.join(__dirname,"../frontend"),{
index:false
}))

app.use("/reports", express.static(path.join(__dirname,"reports")))

// ================= START SERVER =================

const PORT = process.env.PORT || 5000

app.listen(PORT,()=>{

console.log("Server running on port " + PORT)

})
