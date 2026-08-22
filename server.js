const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(express.json());

// Create uploads folder if missing
const uploadDir = path.join(__dirname, 'public', 'uploads');
if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir, { recursive: true });
}

// Serve static frontend and uploaded images
app.use(express.static('public'));
app.use('/uploads', express.static(uploadDir));

// Configure file storage
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, uploadDir),
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, uniqueSuffix + path.extname(file.originalname));
    }
});
const upload = multer({ storage });

const db = new sqlite3.Database('./bhagwati_school.db');

// Ensure tables exist with image support
db.serialize(() => {
    db.run(`CREATE TABLE IF NOT EXISTS notices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        file_url TEXT NOT NULL,
        published_date TEXT NOT NULL
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS results (
        class_grade TEXT PRIMARY KEY,
        file_url TEXT NOT NULL,
        published_date TEXT NOT NULL
    )`);
});

// --- API ENDPOINTS ---

// 1. Get all notices
app.get('/api/notices', (req, res) => {
    db.all('SELECT * FROM notices ORDER BY id DESC', [], (err, rows) => {
        if (err) return res.status(500).json({ error: err.message });
        res.json(rows);
    });
});

// 2. Get result by class
app.get('/api/results/:class_grade', (req, res) => {
    db.get('SELECT * FROM results WHERE class_grade = ?', [req.params.class_grade], (err, row) => {
        if (err) return res.status(500).json({ error: err.message });
        res.json(row || {});
    });
});

// 3. Upload notice or result
app.post('/api/upload', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).json({ message: 'No file uploaded' });

    const fileUrl = `/uploads/${req.file.filename}`;
    const dateStr = new Date().toISOString().split('T')[0];
    const { uploadType, targetClass } = req.body;

    if (uploadType === 'notice') {
        db.run('INSERT INTO notices (file_url, published_date) VALUES (?, ?)', [fileUrl, dateStr], function(err) {
            if (err) return res.status(500).json({ error: err.message });
            res.json({ success: true, url: fileUrl, date: dateStr, id: this.lastID });
        });
    } else {
        db.run(
            `INSERT INTO results (class_grade, file_url, published_date) VALUES (?, ?, ?)
             ON CONFLICT(class_grade) DO UPDATE SET file_url = excluded.file_url, published_date = excluded.published_date`,
            [targetClass, fileUrl, dateStr],
            function(err) {
                if (err) return res.status(500).json({ error: err.message });
                res.json({ success: true, url: fileUrl, date: dateStr, class_grade: targetClass });
            }
        );
    }
});

// 4. Delete notice
app.delete('/api/notices/:id', (req, res) => {
    db.run('DELETE FROM notices WHERE id = ?', [req.params.id], (err) => {
        if (err) return res.status(500).json({ error: err.message });
        res.json({ success: true });
    });
});

// 5. Delete result
app.delete('/api/results/:class_grade', (req, res) => {
    db.run('DELETE FROM results WHERE class_grade = ?', [req.params.class_grade], (err) => {
        if (err) return res.status(500).json({ error: err.message });
        res.json({ success: true });
    });
});

app.listen(5000, () => {
    console.log('Bhagwati School Backend running on http://localhost:5000');
});