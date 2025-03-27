/**
 * Enhanced HTML to plain text converter
 * This module provides better conversion of HTML content to readable plain text,
 * preserving whitespace, list formatting, and other structural elements.
 */

// Configuration options for the converter
interface HtmlToTextOptions {
  wordwrap?: number | false;      // Character limit before wrapping, or false to disable
  preserveNewlines?: boolean;     // Whether to keep existing newlines
  baseElement?: string | string[]; // HTML elements to extract (e.g., body, article)
  tables?: boolean;               // Whether to format tables
  uppercaseHeadings?: boolean;    // Whether to convert headings to uppercase
  preserveHrefLinks?: boolean;     // Whether to include the href links in brackets after the link text
  bulletIndent?: number;          // Indentation for bullets
  listIndent?: number;            // Indentation for lists
  headingStyle?: 'underline' | 'linebreak' | 'hashify'; // How to format headings
  maxLineLength?: number;         // Maximum line length
}

// Default options
const defaultOptions: HtmlToTextOptions = {
  wordwrap: 80,
  preserveNewlines: true,
  tables: true,
  uppercaseHeadings: false,
  preserveHrefLinks: true,
  bulletIndent: 2,
  listIndent: 2,
  headingStyle: 'linebreak',
  maxLineLength: 100
};

/**
 * Converts HTML to plain text while preserving basic formatting
 * @param html The HTML string to convert
 * @param options Configuration options
 * @returns Plain text representation of the HTML
 */
export function htmlToText(html: string, options: HtmlToTextOptions = {}): string {
  // Merge with default options
  const settings: HtmlToTextOptions = { ...defaultOptions, ...options };
  
  // Quick exit if the input is not HTML
  if (!html || !html.includes('<')) {
    return html || '';
  }
  
  // Convert common HTML entities
  let text = html
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (match, dec) => String.fromCharCode(dec))
    .replace(/&#x([0-9a-f]+);/gi, (match, hex) => String.fromCharCode(parseInt(hex, 16)));
  
  // Pre-processing for certain elements
  
  // Convert divs to paragraphs for easier processing
  text = text.replace(/<div\s[^>]*>/gi, '<div>');
  
  // Handle line breaks and paragraphs
  text = text
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<p\s[^>]*>/gi, '<p>')
    .replace(/<\/p>/gi, '\n');
  
  // Handle headings (h1-h6)
  if (settings.headingStyle === 'hashify') {
    // GitHub-style headings with #
    text = text
      .replace(/<h1\s[^>]*>/gi, '\n# ')
      .replace(/<h2\s[^>]*>/gi, '\n## ')
      .replace(/<h3\s[^>]*>/gi, '\n### ')
      .replace(/<h4\s[^>]*>/gi, '\n#### ')
      .replace(/<h5\s[^>]*>/gi, '\n##### ')
      .replace(/<h6\s[^>]*>/gi, '\n###### ');
  } else if (settings.headingStyle === 'underline') {
    // Underline style for headings
    text = text
      .replace(/<h1\s[^>]*>/gi, '\n')
      .replace(/<h2\s[^>]*>/gi, '\n')
      .replace(/<h3\s[^>]*>/gi, '\n')
      .replace(/<h4\s[^>]*>/gi, '\n')
      .replace(/<h5\s[^>]*>/gi, '\n')
      .replace(/<h6\s[^>]*>/gi, '\n');
      
    // We'll handle the underlines after removing tags
  } else {
    // Default: just use line breaks
    text = text
      .replace(/<h1\s[^>]*>/gi, '\n')
      .replace(/<h2\s[^>]*>/gi, '\n')
      .replace(/<h3\s[^>]*>/gi, '\n')
      .replace(/<h4\s[^>]*>/gi, '\n')
      .replace(/<h5\s[^>]*>/gi, '\n')
      .replace(/<h6\s[^>]*>/gi, '\n');
  }
  
  // Close heading tags with a newline
  text = text
    .replace(/<\/h[1-6]>/gi, '\n');
  
  // Handle lists
  // Convert list items to indented bullet points
  const bulletIndent = ' '.repeat(settings.bulletIndent || 2);
  const listIndent = ' '.repeat(settings.listIndent || 2);
  
  // Unordered lists - convert to bullet points
  text = text.replace(/<li\s[^>]*>/gi, '<li>');
  text = text.replace(/<li>/gi, `\n${bulletIndent}• `);
  
  // Ordered lists - convert to numbered items
  // This requires special handling to maintain numbering
  text = processOrderedLists(text, listIndent);
  
  // Handle tables if enabled
  if (settings.tables) {
    text = processTables(text);
  }
  
  // Handle links
  if (settings.preserveHrefLinks) {
    text = processLinks(text);
  }
  
  // Handle blockquotes
  text = text.replace(/<blockquote[^>]*>/gi, '\n> ');
  text = text.replace(/<\/blockquote>/gi, '\n');
  
  // Replace multiple blockquote patterns with deeper nesting
  text = text.replace(/>\s+>/g, '>>');
  
  // Handle pre-formatted text
  text = text
    .replace(/<pre[^>]*>/gi, '\n')
    .replace(/<\/pre>/gi, '\n');
  
  // Handle horizontal rules
  text = text.replace(/<hr[^>]*>/gi, '\n----------------------------\n');
  
  // Handle character styles - just remove them without adding formatting indicators
  text = text
    .replace(/<b>|<strong[^>]*>/gi, '')
    .replace(/<\/b>|<\/strong>/gi, '')
    .replace(/<i>|<em[^>]*>/gi, '')
    .replace(/<\/i>|<\/em>/gi, '')
    .replace(/<u[^>]*>/gi, '')
    .replace(/<\/u>/gi, '')
    .replace(/<s>|<strike>|<del[^>]*>/gi, '')
    .replace(/<\/s>|<\/strike>|<\/del>/gi, '')
    .replace(/<mark[^>]*>/gi, '')
    .replace(/<\/mark>/gi, '');
  
  // Strip remaining HTML tags
  text = text.replace(/<[^>]*>/g, '');
  
  // Post-processing cleanup
  
  // Decode URL-encoded characters
  text = text.replace(/%([0-9A-F]{2})/gi, (match, hex) => {
    try {
      return String.fromCharCode(parseInt(hex, 16));
    } catch (e) {
      return match;
    }
  });
  
  // Cleanup excessive whitespace
  text = text
    .replace(/\n\s+\n/g, '\n')              // Remove extra spaces between paragraphs
    .replace(/\n{2,}/g, '\n')               // No consecutive newlines
    .replace(/\t/g, '    ')                 // Convert tabs to spaces
    .replace(/[ \t]+\n/g, '\n')             // Remove trailing whitespace at end of lines
    .replace(/^\s+/, '')                    // Remove leading whitespace from start of document
    .replace(/\s+$/g, '')                   // Remove trailing whitespace from end of document
    .replace(/[ \t]+$/gm, '');              // Remove trailing whitespace from each line
  
  // Final processing
  
  // Word wrapping if enabled
  if (settings.wordwrap && typeof settings.wordwrap === 'number') {
    text = applyWordWrap(text, settings.wordwrap);
  }
  
  // Final pass to remove trailing whitespace from every line
  text = text.split('\n')
    .map(line => line.trimRight())
    .join('\n');
  
  return text;
}

/**
 * Process ordered lists, maintaining numbering
 */
function processOrderedLists(text: string, indent: string): string {
  // This is a simplistic approach - a more robust solution would use a parser
  const olRegex = /<ol[^>]*>([\s\S]*?)<\/ol>/gi;
  
  return text.replace(olRegex, (match) => {
    let listContent = match;
    let itemNumber = 1;
    
    // Replace each list item with a number
    listContent = listContent.replace(/<li[^>]*>([\s\S]*?)(?=<\/li>)/gi, () => {
      return `\n${indent}${itemNumber++}. `;
    });
    
    // Remove the list tags and closing li tags
    listContent = listContent
      .replace(/<ol[^>]*>/gi, '\n')
      .replace(/<\/ol>/gi, '\n')
      .replace(/<\/li>/gi, '');
    
    return listContent;
  });
}

/**
 * Process tables to convert them to plain text format
 */
function processTables(text: string): string {
  const tableRegex = /<table[^>]*>([\s\S]*?)<\/table>/gi;
  
  return text.replace(tableRegex, (match) => {
    // Extract rows from the table
    const rows: string[][] = [];
    let maxCols = 0;
    
    // Extract header rows
    const headerMatch = /<thead[^>]*>([\s\S]*?)<\/thead>/i.exec(match);
    if (headerMatch) {
      const headerRows = extractTableRows(headerMatch[1]);
      rows.push(...headerRows);
      maxCols = Math.max(maxCols, ...headerRows.map(row => row.length));
    }
    
    // Extract body rows
    const bodyMatch = /<tbody[^>]*>([\s\S]*?)<\/tbody>/i.exec(match);
    if (bodyMatch) {
      const bodyRows = extractTableRows(bodyMatch[1]);
      rows.push(...bodyRows);
      maxCols = Math.max(maxCols, ...bodyRows.map(row => row.length));
    }
    
    // If no thead or tbody, extract rows directly from table
    if (rows.length === 0) {
      const directRows = extractTableRows(match);
      rows.push(...directRows);
      maxCols = Math.max(maxCols, ...directRows.map(row => row.length));
    }
    
    // For email signatures, ensure we have at least 3 columns
    maxCols = Math.max(maxCols, 3);
    
    // Format the table as text
    let result = '\n';
    
    // Add table rows
    rows.forEach((row, rowIndex) => {
      // Start the row with pipe
      let rowText = '| ';
      
      // Add cells with proper content padding
      for (let i = 0; i < maxCols; i++) {
        if (i < row.length) {
          // Add the cell content
          rowText += row[i];
        }
        
        // Add the column separator
        rowText += ' | ';
      }
      
      // Trim the last space if there is one and add newline
      result += rowText.trimRight() + '\n';
    });
    
    // Add a blank line after the table to separate it from following content
    result += '\n';
    
    // Return the table
    return result;
  });
}

/**
 * Extract rows from a table section
 */
function extractTableRows(tableHtml: string): string[][] {
  const rows: string[][] = [];
  const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  
  let rowMatch;
  while ((rowMatch = rowRegex.exec(tableHtml)) !== null) {
    const rowContent = rowMatch[1];
    const cells: string[] = [];
    
    // Extract cells (td or th)
    const cellRegex = /<(td|th)[^>]*>([\s\S]*?)(?=<\/\1>)/gi;
    let cellMatch;
    
    while ((cellMatch = cellRegex.exec(rowContent)) !== null) {
      // Strip tags from cell content and clean whitespace
      let cellText = cellMatch[2]
        .replace(/<[^>]*>/g, '')
        .replace(/\s+/g, ' ')
        .trim();
      
      cells.push(cellText);
    }
    
    if (cells.length > 0) {
      rows.push(cells);
    }
  }
  
  return rows;
}

/**
 * Process links to include href URLs in the text
 */
function processLinks(text: string): string {
  return text.replace(/<a\s+(?:[^>]*?\s+)?href=(["'])(.*?)\1[^>]*>(.*?)<\/a>/gi, (match, quote, url, linkText) => {
    const cleanLinkText = linkText.replace(/<[^>]*>/g, '').trim();
    
    // Don't add the URL if it's the same as the link text
    if (url === cleanLinkText || url === `mailto:${cleanLinkText}`) {
      return cleanLinkText;
    }
    
    // Add the URL in brackets after the link text
    return `${cleanLinkText} [${url}]`;
  });
}

/**
 * Apply word wrapping to the text at the specified width
 */
function applyWordWrap(text: string, width: number): string {
  // Skip wrapping if width is too small
  if (width < 10) return text;
  
  const lines = text.split('\n');
  const wrappedLines: string[] = [];
  
  for (const line of lines) {
    // Skip wrapping for empty lines or lines with formatting characters
    if (line.trim() === '' || line.trim().startsWith('>') || line.trim().startsWith('|')) {
      wrappedLines.push(line);
      continue;
    }
    
    // Determine the indentation of the line
    const indentMatch = line.match(/^(\s+)/);
    const indent = indentMatch ? indentMatch[1] : '';
    const indentWidth = indent.length;
    
    // If there's no text after indentation, just push the line
    if (line.trim() === '') {
      wrappedLines.push(line);
      continue;
    }
    
    // Available width for text after accounting for indentation
    const contentWidth = width - indentWidth;
    
    // Skip wrapping if the effective width is too small
    if (contentWidth < 10) {
      wrappedLines.push(line);
      continue;
    }
    
    // Split the line content (after indentation) into words
    const words = line.substring(indentWidth).split(/\s+/);
    let currentLine = indent;
    
    for (const word of words) {
      // If adding this word would exceed the width, start a new line
      if (currentLine.length + word.length > width && currentLine.length > indentWidth) {
        wrappedLines.push(currentLine.trimRight()); // Trim any trailing whitespace
        currentLine = indent + word;
      } else {
        // Add the word with a space if it's not the first word on the line
        if (currentLine.length > indentWidth) {
          currentLine += ' ' + word;
        } else {
          currentLine += word;
        }
      }
    }
    
    // Add the last line if it has content, trimmed to remove trailing spaces
    if (currentLine.trim()) {
      wrappedLines.push(currentLine.trimRight());
    }
  }
  
  return wrappedLines.join('\n');
}
