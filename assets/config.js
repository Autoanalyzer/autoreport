// Configure Google Drive upload here to avoid prompts.
// Fill in your OAuth Client ID (Web) and keep the folder ID.
// When set, clicking "บันทึกเป็น PDF" will upload directly to Drive.
window.APP_CONFIG = {
  // Example: '1234567890-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx.apps.googleusercontent.com'
  googleClientId: '',
  driveFolderId: '1ZAJn2vQJGfVugVchROfTNHdaxjMC6Zx7',
  // Dedicated folder for Conductivity module
  condFolderId: '17-fq-lvMJYi8u37dtwDNlAr0doRe8Tf6',
  // Dedicated folder for Consistency module
  consistencyFolderId: '19AuhoVwj3vXhwwFvx2t0JjabXlDmeoFu',
  uploadToDrive: true,
  // Optional: Use Apps Script Web App to save PDF to Drive without OAuth
  // Deploy GAS as Web App (Execute as: Me, Who has access: Anyone), then fill below
  gasUrl: 'https://script.google.com/macros/s/AKfycbw1GKVDrhrl43vj5IiBGrxInBtwIJxh3ftLGh0GKs73YU9sBL0a86_0KgXQFdMhkq1N9w/exec',
  gasSecret: 'aa123456789',
  // Optional: explicitly pass folderId to Apps Script (if your GAS uses it)
  gasFolderId: '1ZAJn2vQJGfVugVchROfTNHdaxjMC6Zx7'
};
