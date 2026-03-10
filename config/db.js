const { Client } = require("pg")

const client = new Client({
  host: process.env.PGHOST,
  user: process.env.PGUSER,
  password: process.env.PGPASSWORD,
  database: process.env.PGDATABASE,
  port: process.env.PGPORT,
  ssl: {
    rejectUnauthorized: false
  }
})

client.connect()
.then(async () => {

  console.log("PostgreSQL Connected")

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
