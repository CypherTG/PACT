const fs = require('fs');
const file = 'C:/Users/USER/Desktop/Business Support/PACT/pact-app/src/services/SharePointService.ts';
let content = fs.readFileSync(file, 'utf8');

// Remove the problematic ChargedPerson + 'Id' from the initial payload
content = content.replace(
  /\[COLUMNS\.CASES\.CHARGED_PERSON \+ 'Id'\]: parseInt\(newCase\.chargedPerson, 10\) \|\| null,\n\s*/,
  ''
);

// Fix the shock absorber so it doesn't strip important things
content = content.replace(
  /delete spData\[COLUMNS\.CASES\.CHARGED_PERSON \+ 'Id'\];\n\s*delete spData\[COLUMNS\.CASES\.OFFENCE_CATEGORY \+ 'Id'\];\n\s*delete spData\[COLUMNS\.CASES\.ISSUER_NAME\];\n\s*delete spData\[COLUMNS\.CASES\.SECONDARY_CONTACT\];/g,
  delete spData[COLUMNS.CASES.OFFENCE_CATEGORY + 'Id'];\n            spData[COLUMNS.CASES.OFFENCE_CATEGORY] = newCase.offenceCategoryName;
);

fs.writeFileSync(file, content);
console.log('Fixed pact-app');
