#!/usr/bin/env node

const readline = require('readline');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execSync } = require('child_process');

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

console.log('\n');
console.log('=============================================');
console.log('🚀 Outlook MCP Server Setup Assistant');
console.log('=============================================');
console.log('\n');
console.log('This script will help you set up your Outlook MCP Server.');
console.log('It will guide you through the necessary steps and configuration.');
console.log('\n');

async function askQuestion(question) {
  return new Promise(resolve => {
    rl.question(question, answer => {
      resolve(answer.trim());
    });
  });
}

async function setupServer() {
  console.log('Step 1: Creating .env file\n');
  
  const clientId = await askQuestion('Enter your Microsoft Azure client ID: ');
  
  if (!clientId) {
    console.log('\n❌ Client ID is required.');
    console.log('You can obtain it by registering an app in the Microsoft Entra admin center.');
    console.log('See the README.md file for detailed instructions.');
    process.exit(1);
  }
  
  const port = await askQuestion('Enter port number (default: 3000): ') || '3000';
  const sessionSecret = crypto.randomBytes(64).toString('hex');
  const redirectUri = `http://localhost:${port}/auth/callback`;
  
  const envContent = `# Microsoft Graph API Configuration
CLIENT_ID=${clientId}

# OAuth Configuration
REDIRECT_URI=${redirectUri}

# Server Configuration
PORT=${port}
SESSION_SECRET=${sessionSecret}

# Optional: Logging Configuration
LOG_LEVEL=info
`;

  fs.writeFileSync(path.join(__dirname, '.env'), envContent);
  console.log('\n✅ .env file created successfully.');
  
  console.log('\nStep 2: Installing dependencies\n');
  
  try {
    console.log('Installing npm packages...');
    execSync('npm install', { stdio: 'inherit' });
    console.log('\n✅ Dependencies installed successfully.');
  } catch (error) {
    console.log('\n❌ Failed to install dependencies.');
    console.log('Please run "npm install" manually.');
  }
  
  console.log('\nStep 3: Setting up public directory\n');
  
  const publicDir = path.join(__dirname, 'public');
  if (!fs.existsSync(publicDir)) {
    fs.mkdirSync(publicDir);
    console.log('Created public directory.');
  }
  
  // Ensure the index.html is in the public directory
  const indexSource = path.join(__dirname, 'public', 'index.html');
  if (!fs.existsSync(indexSource)) {
    console.log('\n⚠️ Warning: public/index.html not found.');
    console.log('You may need to create this file manually.');
  }
  
  console.log('\n=============================================');
  console.log('✨ Setup Complete!');
  console.log('=============================================');
  console.log('\nYou can now start the server with:');
  console.log('  npm start');
  console.log('\nThen open your browser and go to:');
  console.log(`  http://localhost:${port}`);
  console.log('\nDon\'t forget to configure the necessary permissions in Microsoft Azure!');
  console.log('See the README.md file for detailed instructions.');
  
  rl.close();
}

setupServer().catch(error => {
  console.error('Error during setup:', error);
  process.exit(1);
});