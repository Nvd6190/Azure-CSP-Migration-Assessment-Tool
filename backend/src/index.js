const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const uploadRouter = require('./routes/upload');

const { version } = require('../package.json');

const app = express();
const PORT = process.env.PORT || 3000;

// Ensure uploads directory exists
const uploadsDir = path.join(__dirname, '..', 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

const allowedOrigins = [
  'http://localhost:4200',
  'https://csp-migration-frontend.azurewebsites.net'
];

app.use(cors({
  origin: function (origin, callback) {
    if (!origin || allowedOrigins.some(o => origin.startsWith(o))) {
      callback(null, true);
    } else {
      callback(null, true); // allow all for now, tighten later
    }
  }
}));
app.use(express.json());

app.use('/api', uploadRouter);

app.get('/api/health', (_req, res) => {
  res.json({ status: 'ok', version, timestamp: new Date().toISOString() });
});

app.listen(PORT, () => {
  console.log(`Resource Migration Assessment API running on port ${PORT}`);
});
