const fs = require('fs');
const files = [
  'C:/Users/USER/Desktop/Business Support/PACT/pact-spfx/pact-native/src/webparts/pactApp/pact/services/SharePointService.ts',
  'C:/Users/USER/Desktop/Business Support/PACT/pact-app/src/services/SharePointService.ts'
];
files.forEach(file => {
  let content = fs.readFileSync(file, 'utf8');
  content = content.replace(
    /\[COLUMNS\.CASES\.CHARGED_PERSON \+ 'Id'\]: parseInt\(newCase\.chargedPerson, 10\) \|\| null,/,
    // [COLUMNS.CASES.CHARGED_PERSON + 'Id']: parseInt(newCase.chargedPerson, 10) || null,
  );
  content = content.replace(
    /delete spData\[COLUMNS\.CASES\.CHARGED_PERSON \+ 'Id'\];/g,
    // delete spData[COLUMNS.CASES.CHARGED_PERSON + 'Id'];
  );
  content = content.replace(
    /delete spData\[COLUMNS\.CASES\.SECONDARY_CONTACT\];/g,
    // delete spData[COLUMNS.CASES.SECONDARY_CONTACT];
  );
  content = content.replace(
    /delete spData\[COLUMNS\.CASES\.OFFENCE_CATEGORY \+ 'Id'\];/g,
    delete spData[COLUMNS.CASES.OFFENCE_CATEGORY + 'Id'];\n            // restore string\n            spData[COLUMNS.CASES.OFFENCE_CATEGORY] = newCase.offenceCategoryName;
  );
  fs.writeFileSync(file, content);
});
console.log('Patch complete.');
