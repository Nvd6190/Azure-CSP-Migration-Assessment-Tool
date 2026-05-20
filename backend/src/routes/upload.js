const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const migrationService = require('../services/migrationService');
const { buildReport } = require('../services/excelReportBuilder');

const router = express.Router();

const storage = multer.diskStorage({
  destination: (_req, _file, cb) => {
    const uploadsDir = path.join(__dirname, '..', '..', 'uploads');
    if (!fs.existsSync(uploadsDir)) {
      fs.mkdirSync(uploadsDir, { recursive: true });
    }
    cb(null, uploadsDir);
  },
  filename: (_req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1e9);
    const ext = path.extname(file.originalname);
    cb(null, `upload-${uniqueSuffix}${ext}`);
  }
});

const fileFilter = (_req, file, cb) => {
  const allowedExtensions = ['.xlsx', '.xls', '.csv'];
  const ext = path.extname(file.originalname).toLowerCase();
  if (allowedExtensions.includes(ext)) {
    cb(null, true);
  } else {
    cb(new Error('Only .xlsx, .xls, and .csv files are allowed'), false);
  }
};

const upload = multer({
  storage,
  fileFilter,
  limits: { fileSize: 10 * 1024 * 1024 }
});

/**
 * POST /api/assess - Upload Excel, download assessed Excel directly
 */
router.post('/assess', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded. Please upload an .xlsx or .csv file.' });
    }

    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet);

    if (rawData.length === 0) {
      cleanupFile(req.file.path);
      return res.status(400).json({ error: 'The uploaded file contains no data rows.' });
    }

    const assessed = migrationService.assessResources(rawData);
    const summary = migrationService.getSummary(assessed);

    const outputFileName = `assessment-${Date.now()}.xlsx`;
    const outputPath = path.join(__dirname, '..', '..', 'uploads', outputFileName);

    await buildReport({
      assessed,
      summary,
      mode: 'subscription',
      sheetName: 'Migration Assessment',
      outputPath,
    });

    res.download(outputPath, outputFileName, () => {
      cleanupFile(req.file.path);
      cleanupFile(outputPath);
    });
  } catch (err) {
    if (req.file) cleanupFile(req.file.path);
    console.error('Assessment error:', err);
    res.status(500).json({ error: 'Failed to process the file. Ensure it is a valid Excel file with resource data.' });
  }
});

/**
 * POST /api/assess-json - Upload Excel, get JSON results (for the Angular frontend)
 */
router.post('/assess-json', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded. Please upload an .xlsx or .csv file.' });
    }

    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet);

    if (rawData.length === 0) {
      cleanupFile(req.file.path);
      return res.status(400).json({ error: 'The uploaded file contains no data rows.' });
    }

    const assessed = migrationService.assessResources(rawData);
    const summary = migrationService.getSummary(assessed);

    const downloadId = `assessment-${Date.now()}`;
    const outputPath = path.join(__dirname, '..', '..', 'uploads', `${downloadId}.xlsx`);

    await buildReport({
      assessed,
      summary,
      mode: 'subscription',
      sheetName: 'Migration Assessment',
      outputPath,
    });

    // Schedule cleanup after 10 minutes
    setTimeout(() => cleanupFile(outputPath), 10 * 60 * 1000);

    cleanupFile(req.file.path);

    res.json({ summary, resources: assessed, downloadId });
  } catch (err) {
    if (req.file) cleanupFile(req.file.path);
    console.error('Assessment error:', err);
    if (err.message && err.message.includes('Could not find')) {
      return res.status(400).json({ error: err.message });
    }
    res.status(500).json({ error: 'Failed to process the file. Ensure it is a valid Excel file with resource data.' });
  }
});

/**
 * GET /api/download/:id - Download a previously generated assessment file
 */
router.get('/download/:id', (req, res) => {
  const downloadId = req.params.id.replace(/[^a-zA-Z0-9-]/g, '');
  const filePath = path.join(__dirname, '..', '..', 'uploads', `${downloadId}.xlsx`);

  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'File not found or expired. Please run the assessment again.' });
  }

  res.download(filePath, `${downloadId}.xlsx`);
});

/**
 * POST /api/assess-region-json - Upload Excel, get JSON results for region move assessment
 */
router.post('/assess-region-json', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded. Please upload an .xlsx or .csv file.' });
    }

    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet);

    if (rawData.length === 0) {
      cleanupFile(req.file.path);
      return res.status(400).json({ error: 'The uploaded file contains no data rows.' });
    }

    const assessed = migrationService.assessRegionResources(rawData);
    const summary = migrationService.getRegionSummary(assessed);

    const downloadId = `region-assessment-${Date.now()}`;
    const outputPath = path.join(__dirname, '..', '..', 'uploads', `${downloadId}.xlsx`);

    await buildReport({
      assessed,
      summary,
      mode: 'region',
      sheetName: 'Region Move Assessment',
      outputPath,
    });

    setTimeout(() => cleanupFile(outputPath), 10 * 60 * 1000);
    cleanupFile(req.file.path);

    res.json({ summary, resources: assessed, downloadId });
  } catch (err) {
    if (req.file) cleanupFile(req.file.path);
    console.error('Region assessment error:', err);
    if (err.message && err.message.includes('Could not find')) {
      return res.status(400).json({ error: err.message });
    }
    res.status(500).json({ error: 'Failed to process the file. Ensure it is a valid Excel file with resource data.' });
  }
});

/**
 * GET /api/rules - Return the current move matrix rules
 */
router.get('/rules', (_req, res) => {
  res.json(migrationService.getRules());
});

/**
 * POST /api/assess-jio-json - Upload Excel, get JSON results for Jio region availability assessment
 */
router.post('/assess-jio-json', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded. Please upload an .xlsx or .csv file.' });
    }

    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet);

    if (rawData.length === 0) {
      cleanupFile(req.file.path);
      return res.status(400).json({ error: 'The uploaded file contains no data rows.' });
    }

    const assessed = migrationService.assessJioResources(rawData);
    const summary = migrationService.getJioSummary(assessed);

    const downloadId = `jio-assessment-${Date.now()}`;
    const outputPath = path.join(__dirname, '..', '..', 'uploads', `${downloadId}.xlsx`);

    await buildReport({
      assessed,
      summary,
      mode: 'jio',
      sheetName: 'Jio Availability Assessment',
      outputPath,
    });

    setTimeout(() => cleanupFile(outputPath), 10 * 60 * 1000);
    cleanupFile(req.file.path);

    res.json({ summary, resources: assessed, downloadId });
  } catch (err) {
    if (req.file) cleanupFile(req.file.path);
    console.error('Jio assessment error:', err);
    if (err.message && err.message.includes('Could not find')) {
      return res.status(400).json({ error: err.message });
    }
    res.status(500).json({ error: 'Failed to process the file. Ensure it is a valid Excel file with resource data.' });
  }
});

/**
 * POST /api/rules/refresh - Force re-fetch the dynamic rules from Microsoft's CSV
 */
router.post('/rules/refresh', async (_req, res) => {
  try {
    const result = await migrationService.refreshRules();
    res.json({ success: true, ...result });
  } catch (err) {
    console.error('Rules refresh failed:', err);
    res.status(500).json({ success: false, error: err.message, source: 'static' });
  }
});

/**
 * POST /api/assess-aws-json - Upload Excel, get JSON results for Azure-to-AWS comparison
 */
router.post('/assess-aws-json', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded. Please upload an .xlsx or .csv file.' });
    }

    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet);

    if (rawData.length === 0) {
      cleanupFile(req.file.path);
      return res.status(400).json({ error: 'The uploaded file contains no data rows.' });
    }

    const assessed = migrationService.assessAwsResources(rawData);
    const summary = migrationService.getAwsSummary(assessed);

    const downloadId = `aws-assessment-${Date.now()}`;
    const outputPath = path.join(__dirname, '..', '..', 'uploads', `${downloadId}.xlsx`);

    await buildReport({
      assessed,
      summary,
      mode: 'aws',
      sheetName: 'AWS to Azure Comparison',
      outputPath,
    });

    setTimeout(() => cleanupFile(outputPath), 10 * 60 * 1000);
    cleanupFile(req.file.path);

    res.json({ summary, resources: assessed, downloadId });
  } catch (err) {
    if (req.file) cleanupFile(req.file.path);
    console.error('AWS assessment error:', err);
    if (err.message && err.message.includes('Could not find')) {
      return res.status(400).json({ error: err.message });
    }
    res.status(500).json({ error: 'Failed to process the file. Ensure it is a valid Excel file with resource data.' });
  }
});

/**
 * POST /api/assess-gcp-json - Upload Excel, get JSON results for GCP-to-Azure comparison
 */
router.post('/assess-gcp-json', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded. Please upload an .xlsx or .csv file.' });
    }

    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet);

    if (rawData.length === 0) {
      cleanupFile(req.file.path);
      return res.status(400).json({ error: 'The uploaded file contains no data rows.' });
    }

    const assessed = migrationService.assessGcpResources(rawData);
    const summary = migrationService.getGcpSummary(assessed);

    const downloadId = `gcp-assessment-${Date.now()}`;
    const outputPath = path.join(__dirname, '..', '..', 'uploads', `${downloadId}.xlsx`);

    await buildReport({
      assessed,
      summary,
      mode: 'gcp',
      sheetName: 'GCP to Azure Comparison',
      outputPath,
    });

    setTimeout(() => cleanupFile(outputPath), 10 * 60 * 1000);
    cleanupFile(req.file.path);

    res.json({ summary, resources: assessed, downloadId });
  } catch (err) {
    if (req.file) cleanupFile(req.file.path);
    console.error('GCP assessment error:', err);
    if (err.message && err.message.includes('Could not find')) {
      return res.status(400).json({ error: err.message });
    }
    res.status(500).json({ error: 'Failed to process the file. Ensure it is a valid Excel file with resource data.' });
  }
});

/**
 * POST /api/jio/refresh - Upload a new Jio availability Excel to refresh the Jio data
 */
router.post('/jio/refresh', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded. Please upload the Jio availability .xlsx file.' });
    }

    const metadata = migrationService.refreshJioFromExcel(req.file.path);
    cleanupFile(req.file.path);

    res.json({ success: true, ...metadata });
  } catch (err) {
    if (req.file) cleanupFile(req.file.path);
    console.error('Jio refresh error:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

function cleanupFile(filePath) {
  try {
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
  } catch {
    // Ignore cleanup errors
  }
}

module.exports = router;
