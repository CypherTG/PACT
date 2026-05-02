const fs = require('fs');
const path = require('path');

const filePath = path.join('src', 'services', 'SharePointService.ts');
const content = fs.readFileSync(filePath, 'utf8');

let braceCount = 0;
let inString = false;
let stringChar = '';
let inComment = false;
let inLineComment = false;

for (let i = 0; i < content.length; i++) {
    const char = content[i];
    const prevChar = content[i-1];
    const nextChar = content[i+1];

    if (inLineComment) {
        if (char === '\n') inLineComment = false;
        continue;
    }

    if (inComment) {
        if (char === '*' && nextChar === '/') {
            inComment = false;
            i++;
        }
        continue;
    }

    if (inString) {
        if (char === stringChar && prevChar !== '\\') {
            inString = false;
        }
        continue;
    }

    if (char === '/' && nextChar === '/') {
        inLineComment = true;
        i++;
        continue;
    }

    if (char === '/' && nextChar === '*') {
        inComment = true;
        i++;
        continue;
    }

    if (char === "'" || char === '"' || char === '`') {
        inString = true;
        stringChar = char;
        continue;
    }

    if (char === '{') {
        braceCount++;
    } else if (char === '}') {
        braceCount--;
    }

    if (braceCount < 0) {
        console.log(`Mismatch detected at character ${i}: "}" without "{"`);
        // Print context
        console.log(content.substring(i - 20, i + 20));
        braceCount = 0; // reset to keep looking
    }
}

console.log(`Final brace count balance: ${braceCount}`);
if (braceCount !== 0) {
    console.log("Unbalanced braces detected!");
}
