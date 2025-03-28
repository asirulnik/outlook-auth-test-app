// Helper function to print email details
function printEmailDetails(email: EmailDetails) {
  console.log('\n==================================================');
  console.log(`Subject: ${email.subject}`);
  console.log(`From: ${email.from?.emailAddress.name || ''} <${email.from?.emailAddress.address || 'Unknown'}>`);
  
  if (email.toRecipients && email.toRecipients.length > 0) {
    console.log('To: ' + email.toRecipients.map(r => 
      `${r.emailAddress.name || ''} <${r.emailAddress.address}>`).join(', '));
  }
  
  if (email.ccRecipients && email.ccRecipients.length > 0) {
    console.log('CC: ' + email.ccRecipients.map(r => 
      `${r.emailAddress.name || ''} <${r.emailAddress.address}>`).join(', '));
  }
  
  if (email.receivedDateTime) {
    console.log(`Date: ${new Date(email.receivedDateTime).toLocaleString()}`);
  }
  
  if (email.attachments && email.attachments.length > 0) {
    console.log('\nAttachments:');
    email.attachments.forEach((attachment, i) => {
      const sizeInKB = Math.round(attachment.size / 1024);
      console.log(`${i + 1}. ${attachment.name} (${attachment.contentType}, ${sizeInKB} KB) - ID: ${attachment.id}`);
    });
  }
  
  console.log('\n--------------------------------------------------');
  if (email.body) {
    // Check if quoted content was removed
    const quotedContentRemoved = email.body.originalContent && email.body.originalContent !== email.body.content;
    
    if (quotedContentRemoved) {
      console.log('Note: Quoted content has been removed from this email.');
    }
    
    if (email.body.contentType === 'html') {
      console.log('Note: This is an HTML email. Plain text conversion shown:');
      
      // Use cached plain text content if available
      if (email.body.plainTextContent) {
        console.log(email.body.plainTextContent);
      } else {
        // Use our enhanced HTML to text converter with formatting preservation
        const textContent = htmlToText(email.body.content, {
          wordwrap: 100, // Adjust based on terminal width
          preserveNewlines: true,
          tables: true,
          preserveHrefLinks: true,
          headingStyle: 'linebreak'
        });
        console.log(textContent);
      }
    } else {
      console.log(email.body.content);
    }
  } else {
    console.log(email.bodyPreview || 'No content');
  }
  console.log('==================================================\n');
}