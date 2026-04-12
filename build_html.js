const fs = require('fs');
const path = require('path');

const SCRIPT_DIR = __dirname;
const TEMPLATE = path.join(SCRIPT_DIR, 'template.html');
const OUTPUT = path.join(SCRIPT_DIR, 'public', 'index.html');
const FREEE_DATA = '/tmp/freee_data/dashboard_data.json';
const SAMPLE_DATA = path.join(SCRIPT_DIR, 'sample_data.json');

console.log('=== BEYOND Holdings Dashboard Build ===');

// Determine data source
let dataPath;
if (process.env.USE_SAMPLE === '1') {
  dataPath = SAMPLE_DATA;
  console.log('Using sample data (USE_SAMPLE=1)');
} else if (fs.existsSync(FREEE_DATA)) {
  dataPath = FREEE_DATA;
  console.log('Using Freee API data');
} else {
  dataPath = SAMPLE_DATA;
  console.log('Freee data not found, falling back to sample data');
}

const dashboardData = JSON.parse(fs.readFileSync(dataPath, 'utf8'));
console.log(`  Companies: ${dashboardData.companies.length}`);
console.log(`  Fiscal Year: ${dashboardData.fiscalYear}`);

// Read template
let template = fs.readFileSync(TEMPLATE, 'utf8');
template = template.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

// Build data block - sanitize to prevent XSS
const jsonStr = JSON.stringify(dashboardData).replace(/<\/script>/gi, '<\\/script>');
const dataBlock = '<script>\nconst DASHBOARD_DATA = ' + jsonStr + ';\n</script>\n';

// Replace placeholder
const placeholder = '// DASHBOARD_DATA_PLACEHOLDER';
const endPlaceholder = '// END_DASHBOARD_DATA_PLACEHOLDER';
const startIdx = template.indexOf(placeholder);
const endIdx = template.indexOf(endPlaceholder);

if (startIdx === -1 || endIdx === -1) {
  console.error('ERROR: Placeholders not found in template!');
  process.exit(1);
}

const before = template.substring(0, startIdx);
const after = template.substring(endIdx + endPlaceholder.length);
const output = before + dataBlock + after;

// Write output
fs.mkdirSync(path.join(SCRIPT_DIR, 'public'), { recursive: true });
fs.writeFileSync(OUTPUT, output, 'utf8');

const lineCount = output.split('\n').length;
console.log(`=== Build complete ===`);
console.log(`Output: ${OUTPUT} (${lineCount} lines)`);
console.log(`  Contains DASHBOARD_DATA: ${output.includes('const DASHBOARD_DATA')}`);
