const sqlite3 = require('sqlite3').verbose();
const bcrypt = require('bcryptjs');

const db = new sqlite3.Database('./bhagwati_school.db', (err) => {
    if (err) return console.error('DB Error:', err.message);
    console.log('Database bhagwati_school.db created successfully.');
});

db.serialize(async () => {
    // 1. Admin accounts
    db.run(`CREATE TABLE IF NOT EXISTS admins (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL
    )`);

    // 2. School Notices
    db.run(`CREATE TABLE IF NOT EXISTS notices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        category TEXT DEFAULT 'General',
        content TEXT NOT NULL,
        published_date DATETIME DEFAULT CURRENT_TIMESTAMP
    )`);

    // 3. Exam Results
    db.run(`CREATE TABLE IF NOT EXISTS results (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        symbol_no TEXT NOT NULL,
        student_name TEXT NOT NULL,
        class_grade TEXT NOT NULL,
        exam_term TEXT NOT NULL,
        gpa TEXT,
        remarks TEXT
    )`);

    // Insert Default Admin (user: admin | pass: Password@123)
    const salt = await bcrypt.genSalt(10);
    const hash = await bcrypt.hash('Password@123', salt);

    db.run(
        `INSERT OR IGNORE INTO admins (id, username, password_hash) VALUES (1, 'admin', ?)`,
        [hash],
        () => console.log('Default admin initialized -> username: admin | password: Password@123')
    );
});

db.close();
