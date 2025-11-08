/**
 * Contact parsing utilities for extracting contact information from text
 */

/**
 * Extract email addresses from text
 * @param {string} text - Text to parse
 * @returns {Array<string>} - Array of email addresses
 */
function extractEmails(text) {
  if (!text) return [];

  // More comprehensive email regex
  const emailRegex = /[\w\.-]+@[\w\.-]+\.\w{2,}/gi;
  const matches = text.match(emailRegex) || [];

  // Filter out common false positives
  return matches
    .filter(email => {
      const lower = email.toLowerCase();
      const localPart = lower.split('@')[0];
      // Filter out image files, common domains to ignore
      return !localPart.endsWith('.png') &&
             !localPart.endsWith('.jpg') &&
             !localPart.endsWith('.gif') &&
             !lower.includes('noreply') &&
             !lower.includes('no-reply') &&
             !lower.includes('donotreply');
    })
    .map(email => email.toLowerCase().trim());
}

/**
 * Extract phone numbers from text (multiple formats)
 * @param {string} text - Text to parse
 * @returns {Array<string>} - Array of phone numbers
 */
function extractPhoneNumbers(text) {
  if (!text) return [];

  const phonePatterns = [
    // US formats: (123) 456-7890, 123-456-7890, 123.456.7890
    /\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}/g,
    // International: +1-234-567-8900, +44 20 1234 5678
    /\+\d{1,3}[\s.-]?\(?\d{1,4}\)?[\s.-]?\d{1,4}[\s.-]?\d{1,9}/g,
    // Simple: 1234567890 (10 digits)
    /\b\d{10}\b/g
  ];

  const numbers = new Set();

  phonePatterns.forEach(pattern => {
    const matches = text.match(pattern) || [];
    matches.forEach(num => {
      // Clean up the number
      const cleaned = num.replace(/[\s.-]/g, '');
      // Only include if it has at least 10 digits
      if (cleaned.length >= 10) {
        numbers.add(num.trim());
      }
    });
  });

  return Array.from(numbers);
}

/**
 * Extract LinkedIn URLs from text
 * @param {string} text - Text to parse
 * @returns {Array<string>} - Array of LinkedIn URLs
 */
function extractLinkedInUrls(text) {
  if (!text) return [];

  const linkedInRegex = /(?:https?:\/\/)?(?:www\.)?linkedin\.com\/in\/[\w-]+\/?/gi;
  const matches = text.match(linkedInRegex) || [];

  return matches.map(url => {
    // Normalize URLs
    if (!url.startsWith('http')) {
      return 'https://' + url;
    }
    return url;
  });
}

/**
 * Extract names from signature blocks
 * @param {string} signatureText - Signature text
 * @returns {string|null} - Extracted name or null
 */
function extractNameFromSignature(signatureText) {
  if (!signatureText) return null;

  // Common signature patterns
  const patterns = [
    // "Best regards,\nJohn Doe"
    /(?:best regards?|regards?|sincerely|thanks?|cheers),?\s*\n\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)/i,
    // "John Doe\nCEO"
    /^([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)\s*\n/m,
    // First line with capitalized words
    /^([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)$/m
  ];

  for (const pattern of patterns) {
    const match = signatureText.match(pattern);
    if (match && match[1]) {
      const name = match[1].trim();
      // Validate it's likely a name (2-4 words, reasonable length)
      const words = name.split(/\s+/);
      if (words.length >= 2 && words.length <= 4 && name.length <= 50) {
        return name;
      }
    }
  }

  return null;
}

/**
 * Extract company name from text
 * @param {string} text - Text to parse
 * @returns {string|null} - Company name or null
 */
function extractCompanyName(text) {
  if (!text) return null;

  const patterns = [
    // "CEO at Company Name"
    /(?:CEO|CTO|CFO|Manager|Director|President|Founder)\s+at\s+([A-Z][A-Za-z\s&,.-]+(?:Inc\.?|LLC\.?|Ltd\.?|Corp\.?|Corporation|Company)?)/i,
    // "Company Name\nJob Title"
    /^([A-Z][A-Za-z\s&,.-]+(?:Inc\.?|LLC\.?|Ltd\.?|Corp\.?|Corporation))\s*\n/m
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match && match[1]) {
      const company = match[1].trim();
      if (company.length >= 2 && company.length <= 100) {
        return company;
      }
    }
  }

  return null;
}

/**
 * Extract job title from text
 * @param {string} text - Text to parse
 * @returns {string|null} - Job title or null
 */
function extractJobTitle(text) {
  if (!text) return null;

  const commonTitles = [
    'CEO', 'CTO', 'CFO', 'COO', 'President', 'Vice President', 'VP',
    'Director', 'Manager', 'Senior Manager', 'Lead', 'Senior',
    'Engineer', 'Developer', 'Designer', 'Architect', 'Consultant',
    'Analyst', 'Specialist', 'Coordinator', 'Administrator',
    'Founder', 'Co-Founder', 'Partner', 'Associate', 'Assistant'
  ];

  // Create regex pattern from common titles
  const titlePattern = new RegExp(
    `\\b(${commonTitles.join('|')})(?:\\s+[A-Z][a-z]+)*\\b`,
    'i'
  );

  // Look for patterns like "John Doe\nSenior Software Engineer"
  const signaturePattern = /\n([A-Z][A-Za-z\s]+(?:Engineer|Developer|Manager|Director|Designer|Architect|Consultant|Analyst|Specialist))/i;

  let match = text.match(signaturePattern);
  if (match && match[1]) {
    const title = match[1].trim();
    if (title.length >= 3 && title.length <= 100) {
      return title;
    }
  }

  // Fallback to common title pattern
  match = text.match(titlePattern);
  if (match && match[0]) {
    const title = match[0].trim();
    if (title.length >= 3 && title.length <= 100) {
      return title;
    }
  }

  return null;
}

/**
 * Parse display name into first and last name
 * @param {string} displayName - Full display name
 * @returns {Object} - {firstName, lastName}
 */
function parseDisplayName(displayName) {
  if (!displayName) return { firstName: null, lastName: null };

  const parts = displayName.trim().split(/\s+/);

  if (parts.length === 0) {
    return { firstName: null, lastName: null };
  } else if (parts.length === 1) {
    return { firstName: parts[0], lastName: null };
  } else {
    // First word is first name, rest is last name
    return {
      firstName: parts[0],
      lastName: parts.slice(1).join(' ')
    };
  }
}

/**
 * Extract contact information from email body text
 * @param {string} bodyText - Email body text
 * @returns {Object} - Extracted contact data
 */
function extractContactsFromBody(bodyText) {
  if (!bodyText) return {};

  return {
    emails: extractEmails(bodyText),
    phoneNumbers: extractPhoneNumbers(bodyText),
    linkedInUrls: extractLinkedInUrls(bodyText),
    companyName: extractCompanyName(bodyText),
    jobTitle: extractJobTitle(bodyText)
  };
}

/**
 * Extract contact from email metadata (sender/recipients)
 * @param {Object} emailAddress - Graph API emailAddress object {name, address}
 * @returns {Object} - Contact data
 */
function extractContactFromEmailAddress(emailAddress) {
  if (!emailAddress || !emailAddress.address) return null;

  const { firstName, lastName } = parseDisplayName(emailAddress.name);

  return {
    email: emailAddress.address.toLowerCase().trim(),
    displayName: emailAddress.name || null,
    firstName,
    lastName,
    source: 'metadata'
  };
}

module.exports = {
  extractEmails,
  extractPhoneNumbers,
  extractLinkedInUrls,
  extractNameFromSignature,
  extractCompanyName,
  extractJobTitle,
  parseDisplayName,
  extractContactsFromBody,
  extractContactFromEmailAddress
};
