const { Pool } = require("pg")

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: { rejectUnauthorized: false }
})

pool.connect()
  .then(async (client) => {
    console.log("PostgreSQL Connected")

    // Users table
    await client.query(`
      CREATE TABLE IF NOT EXISTS users (
        id SERIAL PRIMARY KEY,
        username VARCHAR(100),
        password VARCHAR(100),
        created_date DATE,
        created_time TIME
      )
    `)
    console.log("Users table created")

    // Reports table — includes file_data column for persistent storage
    await client.query(`
      CREATE TABLE IF NOT EXISTS reports (
        id SERIAL PRIMARY KEY,
        username VARCHAR(100),
        filename VARCHAR(255),
        file_data BYTEA,
        created_date DATE,
        created_time TIME
      )
    `)
    console.log("Reports table created")

    // If reports table already existed WITHOUT file_data, add the column safely
    await client.query(`
      ALTER TABLE reports ADD COLUMN IF NOT EXISTS file_data BYTEA
    `)
    console.log("file_data column verified")

    client.release()
  })
  .catch((err) => console.log("DB Error:", err))

module.exports = pool
