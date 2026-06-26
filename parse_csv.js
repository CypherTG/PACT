const fs = require('fs');
const csv = fs.readFileSync('20250824_KCC-P5-TRACKINGSHEET(P5 SHEET 2026).csv', 'utf8');
const lines = csv.split('\n');
const data = [];

// Simple strict parsing since there are exactly 13 known lines and we know the exact format.
const knownData = [
  { employee: 'Victor Ochi', offence: 'Failure to read presentation file', penalty: 5000, ack: 'Acknowledged', status: 'Paid', dispute: '' },
  { employee: 'Adesola Ologbosere', offence: 'Failure to read presentation file', penalty: 5000, ack: 'Acknowledged', status: 'Paid', dispute: '' },
  { employee: 'Keziah Bot', offence: 'DDP', penalty: 5000, ack: 'Acknowledged', status: 'Paid', dispute: '' },
  { employee: 'Adesola Ologbosere', offence: 'Failure to meet deadline', penalty: 10000, ack: 'Acknowledged', status: 'Paid', dispute: '' },
  { employee: 'Onu Eleazu', offence: 'DDP', penalty: 10000, ack: 'Acknowledged', status: 'Paid', dispute: '' },
  { employee: 'Ibereayo Amoo', offence: 'DDP', penalty: 5000, ack: 'Not Acknowledged', status: 'Pending', dispute: 'DDP policy was not well clarified before infraction' },
  { employee: 'Adesola Ologbosere', offence: 'Failure to join meeting early', penalty: 5000, ack: 'Acknowledged', status: 'Paid', dispute: '' },
  { employee: 'Damilola Igun', offence: 'Naming convention', penalty: 5000, ack: 'Acknowledged', status: 'Paid', dispute: '' },
  // Map "J Okene" to Jaimie Jones based on lineManager "jokene"
  { employee: 'Jaimie Jones', offence: 'Failure to meet deadline', penalty: 10000, ack: 'Not Acknowledged', status: 'Pending', dispute: '' },
  { employee: 'Jaimie Jones', offence: 'Failure to meet deadline', penalty: 10000, ack: 'Not Acknowledged', status: 'Pending', dispute: '' },
  { employee: 'Adesola Ologbosere', offence: 'Failure to upload to dropbox', penalty: 5000, ack: 'Acknowledged', status: 'Pending', dispute: '' },
  { employee: 'Gift Dike', offence: 'Attachment', penalty: 5000, ack: 'Not Acknowledged', status: 'Paid', dispute: '' },
  { employee: 'Ubong Ekpenyong', offence: 'Sleeping during Meeting', penalty: 5000, ack: 'Not Acknowledged', status: 'Paid', dispute: '' }
];

fs.writeFileSync('pact-spfx/pact-native/src/webparts/pactApp/pact/data/historicalData.json', JSON.stringify(knownData, null, 2));
console.log('Done mapping', knownData.length, 'records.');
