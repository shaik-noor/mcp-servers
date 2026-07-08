import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { spawnSync, execSync } from 'child_process';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, '..');

// Helper for output formatting
const style = {
  reset: '\x1b[0m',
  bright: '\x1b[1m',
  green: '\x1b[32m',
  yellow: '\x1b[33m',
  red: '\x1b[31m',
  cyan: '\x1b[36m',
  magenta: '\x1b[35m'
};

function log(message, type = 'info') {
  let prefix = '';
  switch (type) {
    case 'success':
      prefix = `${style.green}✔${style.reset} `;
      break;
    case 'warn':
      prefix = `${style.yellow}⚠ [WARN]${style.reset} `;
      break;
    case 'error':
      prefix = `${style.red}✖ [ERROR]${style.reset} `;
      break;
    case 'info':
    default:
      prefix = `${style.cyan}ℹ${style.reset} `;
      break;
  }
  console.log(`${prefix}${message}`);
}

function checkNode() {
  log('Checking Node.js version...');
  const currentVersion = process.versions.node;
  const major = parseInt(currentVersion.split('.')[0], 10);
  
  if (major >= 18) {
    log(`Node.js v${currentVersion} detected (v18+ recommended).`, 'success');
    return true;
  } else {
    log(`Node.js v${currentVersion} detected. You might experience compatibility issues. v18 or higher is recommended.`, 'warn');
    return false;
  }
}

function checkPython() {
  log('Checking Python installation...');
  
  // Try running "python --version"
  const pythonCheck = spawnSync('python', ['--version'], { encoding: 'utf-8' });
  if (pythonCheck.status === 0) {
    const version = (pythonCheck.stdout || pythonCheck.stderr || '').trim();
    log(`Python detected: ${version}`, 'success');
    return true;
  }
  
  // Try running "python3 --version" as a fallback
  const python3Check = spawnSync('python3', ['--version'], { encoding: 'utf-8' });
  if (python3Check.status === 0) {
    const version = (python3Check.stdout || python3Check.stderr || '').trim();
    log(`Python 3 detected (as 'python3'): ${version}`, 'success');
    log("Note: The MCP server spawns 'python' command. Ensure 'python' is aliased or mapped to Python 3 on your PATH.", 'warn');
    return true;
  }
  
  log('Python was not found on your PATH.', 'warn');
  log('Python is required for parsing local offline .one / .onepkg files. If you only plan to use online Microsoft Graph/OneNote API functionalities, you can proceed without it.', 'info');
  return false;
}

function setupEnv() {
  const envPath = path.join(projectRoot, '.env');
  const envExamplePath = path.join(projectRoot, '.env.example');
  
  if (fs.existsSync(envPath)) {
    log('.env file already exists. Skipping creation.', 'success');
    return;
  }
  
  if (!fs.existsSync(envExamplePath)) {
    log('Could not find .env.example template to copy.', 'error');
    return;
  }
  
  try {
    fs.copyFileSync(envExamplePath, envPath);
    log('Created .env from .env.example. Please configure your client credentials in it.', 'success');
  } catch (error) {
    log(`Failed to create .env file: ${error.message}`, 'error');
  }
}

function installDeps() {
  log('Installing Node.js dependencies...');
  
  try {
    execSync('npm install', { 
      cwd: projectRoot,
      stdio: 'inherit'
    });
    log('Dependencies installed successfully.', 'success');
  } catch (error) {
    log('npm install completed with warnings or errors.', 'warn');
  }
}

function buildMcp() {
  log('Building the OneNote MCP server...');
  
  try {
    execSync('npm run build', { 
      cwd: projectRoot,
      stdio: 'inherit'
    });
    log('MCP server built successfully (dist/ folder is ready).', 'success');
  } catch (error) {
    log('MCP server build failed.', 'error');
  }
}

function run() {
  console.log(`\n${style.bright}${style.magenta}=== OneNote MCP Server Environment Setup ===${style.reset}\n`);
  
  checkNode();
  checkPython();
  console.log('');
  setupEnv();
  console.log('');
  installDeps();
  console.log('');
  buildMcp();
  
  console.log(`\n${style.bright}${style.green}Environment setup check complete!${style.reset}`);
  console.log(`Refer to ${style.cyan}README.md${style.reset} to register your app on Microsoft Azure and authenticate.\n`);
}

run();
