const fs = require('fs');
const path = require('path');

const SCRIPT_DIR = __dirname;
const TEMPLATE = path.join(SCRIPT_DIR, 'template.html');
const OUTPUT = path.join(SCRIPT_DIR, 'public', 'index.html');
const FREEE_DATA_LOCAL = path.join(SCRIPT_DIR, 'freee_data.json');
const FREEE_DATA = '/tmp/freee_data/dashboard_data.json';
const SAMPLE_DATA = path.join(SCRIPT_DIR, 'sample_data.json');

console.log('=== BEYOND Holdings Dashboard Build ===');

// Determine data source
let dataPath;
if (process.env.USE_SAMPLE === '1') {
  dataPath = SAMPLE_DATA;
  console.log('Using sample data (USE_SAMPLE=1)');
} else if (fs.existsSync(FREEE_DATA_LOCAL)) {
  dataPath = FREEE_DATA_LOCAL;
  console.log('Using local Freee data');
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

// Merge budget data if available
const BUDGET_FILE = path.join(SCRIPT_DIR, 'budget_data.json');
if (fs.existsSync(BUDGET_FILE)) {
  const budgetData = JSON.parse(fs.readFileSync(BUDGET_FILE, 'utf8'));
  console.log(`  Budget data: ${budgetData.length} companies`);
  for (const co of dashboardData.companies) {
    const budgetCo = budgetData.find(b => b.companyId === co.id);
    if (budgetCo) {
      co.budgets = budgetCo.budgets;
      console.log(`    ${co.name}: ${budgetCo.budgets.length} months budget loaded`);
    }
  }
} else {
  console.log('  No budget data found (budget_data.json)');
}

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

// Write output to both public/ and docs/ (GitHub Pages serves from /docs)
fs.mkdirSync(path.join(SCRIPT_DIR, 'public'), { recursive: true });
fs.mkdirSync(path.join(SCRIPT_DIR, 'docs'), { recursive: true });
fs.writeFileSync(OUTPUT, output, 'utf8');
fs.writeFileSync(path.join(SCRIPT_DIR, 'docs', 'index.html'), output, 'utf8');

const lineCount = output.split('\n').length;
console.log(`=== Build complete ===`);
console.log(`Output: ${OUTPUT} (${lineCount} lines)`);
console.log(`  Contains DASHBOARD_DATA: ${output.includes('const DASHBOARD_DATA')}`);
