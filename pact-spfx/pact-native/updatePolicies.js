const fs = require('fs');
const path = require('path');

const csvPath = path.join(__dirname, '..', '..', 'PACT Policy & Offence Library (1).csv');
const jsonPath = path.join(__dirname, 'src', 'webparts', 'pactApp', 'pact', 'data', 'policyLibrary.json');

const csvData = fs.readFileSync(csvPath, 'utf8');
const oldJsonData = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));

// Parse CSV (simple parser considering quotes)
const lines = csvData.split('\n').filter(l => l.trim().length > 0 && !l.startsWith(',,'));

const newPolicies = [];
for (let i = 1; i < lines.length; i++) {
  const line = lines[i];
  if (!line || line.trim() === '') continue;
  
  // Custom regex to handle comma inside quotes
  const parts = [];
  let current = '';
  let inQuotes = false;
  for (let char of line) {
    if (char === '"') inQuotes = !inQuotes;
    else if (char === ',' && !inQuotes) {
      parts.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  parts.push(current.trim());
  
  if (parts.length < 5) continue; // Skip malformed rows
  
  const offenceName = parts[0].replace(/^"|"$/g, '').trim();
  const category = parts[1].replace(/^"|"$/g, '').trim();
  const amountStr = parts[2].replace(/^"|"$/g, '').replace(/[^0-9.]/g, '');
  const amount = parseFloat(amountStr) || 0;
  const tier = parts[3].replace(/^"|"$/g, '').trim();
  const first = parts[4] ? parts[4].replace(/^"|"$/g, '').trim() : '';
  const second = parts[5] ? parts[5].replace(/^"|"$/g, '').trim() : '';
  const third = parts[6] ? parts[6].replace(/^"|"$/g, '').trim() : '';
  
  if (!offenceName) continue;
  
  const existing = oldJsonData.find(p => p.offenceName.toLowerCase() === offenceName.toLowerCase());
  
  newPolicies.push({
    id: existing ? existing.id : 'p' + (i + 100),
    offenceName: offenceName,
    tier: tier,
    category: category,
    description: existing ? existing.description : offenceName,
    defaultPenaltyAmount: amount,
    firstOffenceAction: first,
    secondOffenceAction: second,
    thirdOffenceAction: third,
    escalationTrigger: existing ? existing.escalationTrigger : true
  });
}

fs.writeFileSync(jsonPath, JSON.stringify(newPolicies, null, 2), 'utf8');
console.log('Successfully updated policyLibrary.json with ' + newPolicies.length + ' policies.');
