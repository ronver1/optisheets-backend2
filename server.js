require('dotenv').config();
const express = require('express');

const validateKey        = require('./api/validate-key');
const getRecommendations = require('./api/get-recommendations');
const addCredits         = require('./api/add-credits');
const exportCsv          = require('./api/admin/export-csv');

const app = express();
app.use(express.json());

app.post('/api/validate-key',        validateKey);
app.post('/api/get-recommendations', getRecommendations);
app.post('/api/add-credits',         addCredits);
app.get('/api/admin/export-csv',     exportCsv);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`OptiSheets API running at http://localhost:${PORT}`);
  console.log('  POST /api/validate-key');
  console.log('  POST /api/get-recommendations');
  console.log('  POST /api/add-credits         (Bearer ADMIN_SECRET)');
  console.log('  GET  /api/admin/export-csv    (Bearer ADMIN_SECRET)');
});
