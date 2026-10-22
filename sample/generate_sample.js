/**
 * Generates the sample Excel file: sample/ukazkovy_dokument.xlsx
 * Run with: node sample/generate_sample.js
 */

const XLSX = require('xlsx');
const path = require('path');
const SampleGenerator = require('../js/core/sample-generator');

const data = SampleGenerator.getSampleData();
const headers = SampleGenerator.getSampleHeaders();

// Create worksheet with explicit header order
const worksheet = XLSX.utils.json_to_sheet(data, { header: headers });

// Auto-adjust column widths for readability
const colWidths = headers.map(header => {
  let maxLen = header.length;
  data.forEach(row => {
    const val = row[header] ? String(row[header]) : '';
    if (val.length > maxLen) {
      maxLen = Math.min(val.length, 60); // Cap width at 60
    }
  });
  return { wch: Math.max(maxLen + 3, 12) };
});

worksheet['!cols'] = colWidths;

// Create workbook
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');

const outputPath = path.join(__dirname, 'ukazkovy_dokument.xlsx');
XLSX.writeFile(workbook, outputPath);

console.log(`✓ Vzorový soubor Excel byl úspěšně vygenerován: ${outputPath}`);
console.log(`  Počet článků: ${data.length}`);
console.log(`  Sloupců celkem: ${headers.length}`);
