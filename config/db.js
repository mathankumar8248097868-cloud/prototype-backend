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
  await client.query("DROP TABLE IF EXISTS reports")
  await client.query("DROP TABLE IF EXISTS users")

  // CREATE USERS TABLE
  const usersTable = `
  CREATE TABLE IF NOT EXISTS users (
    id SERIAL PRIMARY KEY,
    username VARCHAR(50),
    password VARCHAR(50),
    created_date DATE,
    created_time TIME
  )
  `

  await client.query(usersTable)
  console.log("Users table created")

  // CREATE REPORTS TABLE
  const reportsTable = `
  CREATE TABLE IF NOT EXISTS reports (
    id SERIAL PRIMARY KEY,
    username VARCHAR(50),
    filename VARCHAR(255),
    created_date DATE,
    created_time TIME
  )
  `

  await client.query(reportsTable)
  console.log("Reports table created")

})
.catch(err => {
  console.log("Database connection error:", err)
})

module.exports = client
