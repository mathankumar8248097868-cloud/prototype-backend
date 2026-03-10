const { Pool } = require("pg")

const pool = new Pool({
  connectionString: "postgresql://postgres:YOUR_PASSWORD@db.sfnzubekxitfdcsjfhcp.supabase.co:5432/postgres",
  ssl: {
    rejectUnauthorized: false
  }
})

pool.connect((err, client, release) => {
  if (err) {
    console.log("Database error", err)
  } else {

    console.log("PostgreSQL Connected")

    // DROP tables
    client.query("DROP TABLE IF EXISTS reports")
    client.query("DROP TABLE IF EXISTS users")

    // CREATE USERS TABLE
    const usersTable = `
    CREATE TABLE users (
      id SERIAL PRIMARY KEY,
      username VARCHAR(50),
      password VARCHAR(50),
      created_date DATE,
      created_time TIME
    )
    `

    client.query(usersTable, (err) => {
      if (err) {
        console.log("Users table error", err)
      } else {
        console.log("Users table created")
      }
    })

    // CREATE REPORTS TABLE
    const reportsTable = `
    CREATE TABLE reports (
      id SERIAL PRIMARY KEY,
      username VARCHAR(50),
      filename VARCHAR(255),
      created_date DATE,
      created_time TIME
    )
    `

    client.query(reportsTable, (err) => {
      if (err) {
        console.log("Reports table error", err)
      } else {
        console.log("Reports table created")
      }
    })

    release()
  }
})

module.exports = pool
