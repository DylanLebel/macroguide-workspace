const fs = require('fs');
const html = fs.readFileSync('C:\\AllMacros\\MacroGuide.html', 'utf8');
const re = /<script[^>]*>([\s\S]*?)<\/script>/g;
let m, i = 0, bad = 0;
while ((m = re.exec(html))) {
  i++;
  try { new Function(m[1]); }
  catch (e) {
    bad++;
    console.log('script #' + i + ' SYNTAX ERROR: ' + e.message);
    // find approx line of error
    const lines = html.slice(0, m.index + m[0].indexOf(m[1])).split('\n').length;
    console.log('  starts around line', lines);
  }
}
console.log('checked', i, 'script blocks,', bad, 'errors');
