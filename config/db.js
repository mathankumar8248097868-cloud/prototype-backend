const { Pool } = require("pg")

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: {
    rejectUnauthorized: false
  }
})

pool.connect(async (err, client, release) => {
  if (err) {
    console.log("Database error", err)
  } else {

    console.log("PostgreSQL Connected")

    try {

      await client.query("DROP TABLE IF EXISTS reports")
      await client.query("DROP TABLE IF EXISTS users")

      await client.query(`
        CREATE TABLE users (
          id SERIAL PRIMARY KEY,
          username VARCHAR(50),
          password VARCHAR(50),
          created_date DATE,
          created_time TIME
        )
      `)

      console.log("Users table created")

      await client.query(`
        CREATE TABLE reports (
          id SERIAL PRIMARY KEY,
          username VARCHAR(50),
          filename VARCHAR(255),
          created_date DATE,
          created_time TIME
        )
      `)

      console.log("Reports table created")

    } catch (error) {
      console.log("Table creation error", error)
    }

    release()
  }
})

module.exports = pool
