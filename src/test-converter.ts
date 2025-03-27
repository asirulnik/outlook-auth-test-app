import * as fs from 'fs';
import * as path from 'path';
import { htmlToText } from './htmlToText';

/**
 * Simple script to test the HTML to text converter
 * Usage: npx ts-node src/test-converter.ts path/to/html/file.html
 */
function main() {
  // Get the file path from command line arguments
  const args = process.argv.slice(2);
  if (args.length < 1) {
    console.error('Please provide the path to an HTML file');
    console.error('Usage: npx ts-node src/test-converter.ts path/to/html/file.html');
    process.exit(1);
  }

  const filePath = args[0];
  
  try {
    // Read the HTML file
    const html = fs.readFileSync(filePath, 'utf8');
    
    // Convert to text using our converter
    const plainText = htmlToText(html, {
      wordwrap: 100,
      preserveNewlines: true,
      tables: true,
      preserveHrefLinks: true,
      headingStyle: 'linebreak'
    });
    
    // Print the result
    console.log('==== HTML to Text Conversion ====\n');
    console.log(plainText);
    console.log('\n==== End of Conversion ====');
    
    // Optionally write to a text file
    const outputPath = filePath.replace(/\.html$/, '') + '.txt';
    fs.writeFileSync(outputPath, plainText);
    console.log(`\nConversion saved to: ${outputPath}`);
    
  } catch (error) {
    console.error('Error processing the file:', error);
    process.exit(1);
  }
}

// Run the main function
main();
