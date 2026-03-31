// Get elements
const scanBtn = document.getElementById('scanBtn');
const resetBtn = document.getElementById('resetBtn');
const statusBox = document.getElementById('statusBox');
const loadingBox = document.getElementById('loadingBox');
const loadingText = document.getElementById('loadingText');

// Event listeners
scanBtn.addEventListener('click', startScan);
resetBtn.addEventListener('click', reset);

async function startScan() {
    scanBtn.disabled = true;
    hideStatus();
    
    // Directly show PowerShell instructions
    showPowerShellInstructions();
}

function showPowerShellInstructions() {
    const psCommand = `cd "$env:USERPROFILE\\Desktop\\ascm"; .\\dist\\"Scan Software.exe"`;
    
    const html = `
        <div style="background-color: #fff9e6; border: 2px solid #ff9800; border-radius: 6px; padding: 15px; margin-bottom: 15px;">
            <h3 style="margin: 0 0 10px; color: #ff6f00; font-size: 14px;">📋 Copy & Run This Command</h3>
            <div style="background-color: #2d2d2d; border-radius: 4px; padding: 12px; margin-bottom: 10px;">
                <code style="color: #f8f8f2; font-family: 'Courier New', monospace; font-size: 12px; word-break: break-all; display: block;">${psCommand}</code>
            </div>
            <button id="copyCmd" style="width: 100%; padding: 8px; background-color: #ff9800; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; margin-bottom: 10px;">📋 Copy Command</button>
            <p style="margin: 0; font-size: 12px; color: #333; line-height: 1.6;">
                <strong>Steps:</strong><br>
                1. Click "Copy Command" above<br>
                2. Open PowerShell as Administrator<br>
                3. Paste the command and press Enter<br>
                4. Excel report will open automatically
            </p>
        </div>
    `;
    
    statusBox.innerHTML = html;
    statusBox.classList.add('show');
    
    // Add copy button functionality
    const copyBtn = document.getElementById('copyCmd');
    if (copyBtn) {
        copyBtn.addEventListener('click', () => {
            navigator.clipboard.writeText(psCommand).then(() => {
                copyBtn.textContent = '✅ Copied!';
                copyBtn.style.backgroundColor = '#4caf50';
                setTimeout(() => {
                    copyBtn.textContent = '📋 Copy Command';
                    copyBtn.style.backgroundColor = '#ff9800';
                }, 2000);
            });
        });
    }
}

function showStatus(message, type = 'info') {
    statusBox.innerHTML = `<p>${message}</p>`;
    statusBox.classList.add('show');
    statusBox.classList.remove('success', 'error', 'warning');
    if (type !== 'info') {
        statusBox.classList.add(type);
    }
}

function hideStatus() {
    statusBox.classList.remove('show');
}

function reset() {
    scanBtn.disabled = false;
    hideStatus();
    if (loadingBox) {
        loadingBox.style.display = 'none';
    }
}

// Check if extension has proper permissions on load
document.addEventListener('DOMContentLoaded', () => {
    console.log('ASCM Compliance Checker extension loaded');
    chrome.permissions.getAll((permissions) => {
        console.log('Current permissions:', permissions);
    });
});
