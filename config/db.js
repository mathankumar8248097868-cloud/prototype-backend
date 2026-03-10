const mysql = require("mysql2")

const db = mysql.createConnection({
  host: process.env.MYSQLHOST,
  user: process.env.MYSQLUSER,
  password: process.env.MYSQLPASSWORD,
  database: process.env.MYSQLDATABASE,
  port: process.env.MYSQLPORT
})

db.connect(err => {
  if (err) {
    console.log("Database error", err)
  } else {

    console.log("MySQL Connected")

    // DROP tables
    db.query("DROP TABLE IF EXISTS reports")
    db.query("DROP TABLE IF EXISTS users")

    // CREATE USERS TABLE
    const usersTable = `
    CREATE TABLE users (
      id INT AUTO_INCREMENT PRIMARY KEY,
      username VARCHAR(50),
      password VARCHAR(50),
      created_date DATE,
      created_time TIME
    )
    `

    db.query(usersTable,(err)=>{
      if(err){
        console.log("Users table error",err)
      }else{
        console.log("Users table created")
      }
    })

    // CREATE REPORTS TABLE
    const reportsTable = `
    CREATE TABLE reports (
      id INT AUTO_INCREMENT PRIMARY KEY,
      username VARCHAR(50),
      filename VARCHAR(255),
      created_date DATE,
      created_time TIME
    )
    `

    db.query(reportsTable,(err)=>{
      if(err){
        console.log("Reports table error",err)
      }else{
        console.log("Reports table created")
      }
    })

  }
})

module.exports = db
