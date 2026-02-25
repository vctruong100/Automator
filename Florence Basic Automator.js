
// ==UserScript==
// @name Florence Basic Automator
// @namespace vinh.activity.plan.state
// @version 1.1.0
// @description
// @match https://us.v2.researchbinders.com/*
// @updateURL    https://raw.githubusercontent.com/vctruong100/Automator/heads/main/Florence%20Basic%20Automator.js
// @downloadURL  https://raw.githubusercontent.com/vctruong100/Automator/heads/main/Florence%20Basic%20Automator.js
// @run-at document-idle
// @grant GM.openInTab
// @grant GM_openInTab
// @grant GM.xmlHttpRequest
// ==/UserScript==

(function () {

    //==========================
    // SHARED GUI AND PANEL FUNCTIONS
    //==========================
    // This section contains functions used by multiple features for panel management,
    // visibility control, hotkey handling, and UI interactions. These functions are
    // shared across all automation features and provide the common user interface.
    //==========================

    let guiVisible = false;
    let guiScale = 1;
    let isDragging = false;
    let dragOffsetX = 0;
    let dragOffsetY = 0;
    let logMessages = [];
    let lastScrollPosition = 0;
    
    const originalLog = console.log;
    const originalError = console.error;
    const originalWarn = console.warn;

    function addLogMessage(message, type = 'log') {
        const timestamp = new Date().toLocaleTimeString();
        logMessages.push({ timestamp, message, type });
        updateLogBox();
    }

    console.log = function(...args) {
        originalLog.apply(console, args);
        addLogMessage(args.join(' '), 'log');
    };

    console.error = function(...args) {
        originalError.apply(console, args);
        addLogMessage(args.join(' '), 'error');
    };

    console.warn = function(...args) {
        originalWarn.apply(console, args);
        addLogMessage(args.join(' '), 'warn');
    };

    function openSignaturesInputGUI() {
        addLogMessage('openSignaturesInputGUI: opening input modal', 'log');
        const modal = document.createElement('div');
        modal.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0, 0, 0, 0.7);
        z-index: 20000;
        display: flex;
        align-items: center;
        justify-content: center;
    `;

        const container = document.createElement('div');
        container.style.cssText = `
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 12px;
        padding: 24px;
        width: 450px;
        max-width: 90%;
        box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3);
        position: relative;
    `;

        const header = document.createElement('div');
        header.style.cssText = `
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 20px;
    `;

        const title = document.createElement('h3');
        title.textContent = 'Add Signatures';
        title.style.cssText = `
        margin: 0;
        color: white;
        font-size: 18px;
        font-weight: 600;
        letter-spacing: 0.2px;
    `;

        const closeButton = document.createElement('button');
        closeButton.innerHTML = '✕';
        closeButton.style.cssText = `
        background: rgba(255, 255, 255, 0.2);
        border: none;
        color: white;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        cursor: pointer;
        font-size: 16px;
        display: flex;
        align-items: center;
        justify-content: center;
        transition: all 0.3s ease;
    `;
        closeButton.onmouseover = () => {
            closeButton.style.background = 'rgba(255, 67, 54, 0.8)';
        };
        closeButton.onmouseout = () => {
            closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        };
        closeButton.onclick = () => {
            addLogMessage('openSignaturesInputGUI: modal closed by user', 'warn');
            document.body.removeChild(modal);
        };

        header.appendChild(title);
        header.appendChild(closeButton);

        const description = document.createElement('p');
        description.textContent = 'Paste or type the full list of signers. Separate names with commas or place each name on a new line.';
        description.style.cssText = `
        color: rgba(255, 255, 255, 0.9);
        margin: 0 0 12px 0;
        font-size: 14px;
        line-height: 1.4;
    `;

        const textarea = document.createElement('textarea');
        textarea.placeholder = 'Name1, Name2, Name3\nor\nName1\nName2\nName3';
        textarea.style.cssText = `
        width: 100%;
        height: 140px;
        padding: 12px 14px;
        border: 2px solid rgba(255, 255, 255, 0.35);
        border-radius: 10px;
        background: rgba(255, 255, 255, 0.95);
        color: #1e293b;
        font-size: 14px;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        resize: vertical;
        outline: none;
        transition: all 0.25s ease;
        box-shadow: 0 2px 0 rgba(0,0,0,0.04) inset;
    `;
        textarea.onfocus = () => {
            textarea.style.borderColor = '#8ea0ff';
            textarea.style.boxShadow = '0 0 0 4px rgba(102, 126, 234, 0.25)';
        };
        textarea.onblur = () => {
            textarea.style.borderColor = 'rgba(255, 255, 255, 0.35)';
            textarea.style.boxShadow = '0 2px 0 rgba(0,0,0,0.04) inset';
        };

        const buttonContainer = document.createElement('div');
        buttonContainer.style.cssText = `
        display: flex;
        gap: 12px;
        margin-top: 20px;
        justify-content: flex-end;
    `;

        const clearButton = document.createElement('button');
        clearButton.textContent = 'Clear All';
        clearButton.style.cssText = `
        background: rgba(255, 255, 255, 0.18);
        border: 2px solid rgba(255, 255, 255, 0.35);
        color: white;
        padding: 10px 20px;
        border-radius: 8px;
        cursor: pointer;
        font-size: 14px;
        font-weight: 500;
        transition: all 0.25s ease;
        backdrop-filter: blur(2px);
    `;
        clearButton.onmouseover = () => {
            clearButton.style.background = 'rgba(255, 255, 255, 0.28)';
        };
        clearButton.onmouseout = () => {
            clearButton.style.background = 'rgba(255, 255, 255, 0.18)';
        };
        clearButton.onclick = () => {
            addLogMessage('openSignaturesInputGUI: Clear All clicked', 'log');
            textarea.value = '';
        };

        const confirmButton = document.createElement('button');
        confirmButton.textContent = 'Confirm';
        confirmButton.style.cssText = `
        background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
        border: 2px solid rgba(255, 255, 255, 0.35);
        color: white;
        padding: 10px 20px;
        border-radius: 8px;
        cursor: pointer;
        font-size: 14px;
        font-weight: 600;
        letter-spacing: 0.2px;
        transition: all 0.25s ease;
    `;
        confirmButton.onmouseover = () => {
            confirmButton.style.background = 'linear-gradient(135deg, #218838 0%, #1ea085 100%)';
        };
        confirmButton.onmouseout = () => {
            confirmButton.style.background = 'linear-gradient(135deg, #28a745 0%, #20c997 100%)';
        };
        confirmButton.onclick = () => {
            addLogMessage('openSignaturesInputGUI: Confirm clicked', 'log');
            const names = parseNames(textarea.value);
            addLogMessage('openSignaturesInputGUI: parsed ' + names.length + ' name(s)', 'log');
            if (names.length === 0) {
                addLogMessage('openSignaturesInputGUI: no names entered, showing warning', 'warn');
                showWarning('Please enter at least one name.');
                return;
            }
            document.body.removeChild(modal);
            processSignatures(names);
        };

        buttonContainer.appendChild(clearButton);
        buttonContainer.appendChild(confirmButton);

        container.appendChild(header);
        container.appendChild(description);
        container.appendChild(textarea);
        container.appendChild(buttonContainer);

        modal.appendChild(container);

        document.body.appendChild(modal);

        textarea.focus();
    }

    
    function showWarning(message) {
        const modal = document.createElement('div');
        modal.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.7);
            z-index: 30000;
            display: flex;
            align-items: center;
            justify-content: center;
        `;

        const container = document.createElement('div');
        container.style.cssText = `
            background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
            border-radius: 12px;
            padding: 24px;
            width: 400px;
            max-width: 90%;
            box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3);
            position: relative;
        `;

        const header = document.createElement('div');
        header.style.cssText = `
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
        `;

        const title = document.createElement('h3');
        title.textContent = 'Warning';
        title.style.cssText = `
            margin: 0;
            color: white;
            font-size: 18px;
            font-weight: 600;
        `;

        const closeButton = document.createElement('button');
        closeButton.innerHTML = '✕';
        closeButton.style.cssText = `
            background: rgba(255, 255, 255, 0.2);
            border: none;
            color: white;
            width: 28px;
            height: 28px;
            border-radius: 50%;
            cursor: pointer;
            font-size: 16px;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.3s ease;
        `;
        closeButton.onmouseover = () => closeButton.style.background = 'rgba(255, 255, 255, 0.3)';
        closeButton.onmouseout = () => closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        closeButton.onclick = () => document.body.removeChild(modal);

        header.appendChild(title);
        header.appendChild(closeButton);

        const messageDiv = document.createElement('p');
        messageDiv.textContent = message;
        messageDiv.style.cssText = `
            color: rgba(255, 255, 255, 0.9);
            margin: 0;
            font-size: 14px;
            line-height: 1.5;
        `;

        const okButton = document.createElement('button');
        okButton.textContent = 'OK';
        okButton.style.cssText = `
            background: rgba(255, 255, 255, 0.2);
            border: 2px solid rgba(255, 255, 255, 0.3);
            color: white;
            padding: 10px 20px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            transition: all 0.3s ease;
            margin-top: 20px;
            width: 100%;
        `;
        okButton.onmouseover = () => okButton.style.background = 'rgba(255, 255, 255, 0.3)';
        okButton.onmouseout = () => okButton.style.background = 'rgba(255, 255, 255, 0.2)';
        okButton.onclick = () => document.body.removeChild(modal);

        container.appendChild(header);
        container.appendChild(messageDiv);
        container.appendChild(okButton);
        modal.appendChild(container);

        document.body.appendChild(modal);
    }
function makeDraggable(container, handle) {
    let isDraggingModal = false;
    let offsetX = 0;
    let offsetY = 0;
    
    handle.style.cursor = 'move';
    
    handle.addEventListener('mousedown', function(e) {
        if (e.target.tagName === 'BUTTON' || e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
        isDraggingModal = true;
        const rect = container.getBoundingClientRect();
        offsetX = e.clientX - rect.left;
        offsetY = e.clientY - rect.top;
        e.preventDefault();
    });
    
    document.addEventListener('mousemove', function(e) {
        if (!isDraggingModal) return;
        const newX = e.clientX - offsetX;
        const newY = e.clientY - offsetY;
        container.style.left = newX + 'px';
        container.style.top = newY + 'px';
        container.style.right = 'auto';
        container.style.transform = 'none';
    });
    
    document.addEventListener('mouseup', function() {
        isDraggingModal = false;
    });
}

        function createGUI() {
        const guiContainer = document.createElement('div');
        guiContainer.id = 'florence-gui';
        guiContainer.style.cssText = `
        position: fixed;
        top: 100px;
        right: 100px;
        width: 350px;
        min-height: 400px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 12px;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
        z-index: 10000;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        display: none;
        flex-direction: column;
        overflow: hidden;
        transform-origin: top right;
    `;

        const header = document.createElement('div');
        header.style.cssText = `
        background: rgba(255, 255, 255, 0.1);
        padding: 12px 16px;
        border-bottom: 1px solid rgba(255, 255, 255, 0.2);
        cursor: move;
        display: flex;
        justify-content: space-between;
        align-items: center;
    `;

        const title = document.createElement('h3');
        title.textContent = 'Florence Automator';
        title.style.cssText = `
        margin: 0;
        color: white;
        font-size: 16px;
        font-weight: 600;
    `;

        const closeButton = document.createElement('button');
        closeButton.innerHTML = '✕';
        closeButton.style.cssText = `
        background: rgba(255, 255, 255, 0.2);
        border: none;
        color: white;
        width: 24px;
        height: 24px;
        border-radius: 50%;
        cursor: pointer;
        font-size: 14px;
        display: flex;
        align-items: center;
        justify-content: center;
        transition: all 0.3s ease;
    `;
        closeButton.onmouseover = () => closeButton.style.background = 'rgba(255, 67, 54, 0.8)';
        closeButton.onmouseout = () => closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        closeButton.onclick = () => toggleGUI();

        header.appendChild(title);
        header.appendChild(closeButton);

        const buttonsContainer = document.createElement('div');
        buttonsContainer.style.cssText = `
        padding: 16px;
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 10px;
        background: rgba(255, 255, 255, 0.05);
    `;

        for (let i = 1; i <= 4; i++) {
            const button = document.createElement('button');
            if (i === 1) {
                button.textContent = 'Add Signatures';
            } else {
                button.textContent = `Placeholder ${i}`;
            }
            button.style.cssText = `
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border: 2px solid rgba(255, 255, 255, 0.3);
            color: white;
            padding: 12px 16px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            transition: all 0.3s ease;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
        `;
            button.onmouseover = () => {
                button.style.transform = 'translateY(-2px)';
                button.style.boxShadow = '0 6px 20px rgba(0, 0, 0, 0.3)';
            };
            button.onmouseout = () => {
                button.style.transform = 'translateY(0)';
                button.style.boxShadow = '0 4px 15px rgba(0, 0, 0, 0.2)';
            };

            if (i === 1) {
                button.onclick = () => {
                    console.log('Add Signatures button clicked');
                    startAddSignaturesFlow();
                };
            } else {
                button.onclick = () => console.log(`Placeholder ${i} clicked`);
            }

            buttonsContainer.appendChild(button);
        }

        const scaleContainer = document.createElement('div');
        scaleContainer.style.cssText = `
        padding: 12px 16px;
        background: rgba(255, 255, 255, 0.05);
        border-top: 1px solid rgba(255, 255, 255, 0.1);
    `;

        const scaleLabel = document.createElement('div');
        scaleLabel.textContent = `Scale: ${guiScale.toFixed(2)}x`;
        scaleLabel.style.cssText = `
        color: white;
        font-size: 12px;
        margin-bottom: 8px;
        font-weight: 500;
    `;

        const scaleSlider = document.createElement('input');
        scaleSlider.type = 'range';
        scaleSlider.min = '0.75';
        scaleSlider.max = '1';
        scaleSlider.step = '0.05';
        scaleSlider.value = guiScale;
        scaleSlider.style.cssText = `
        width: 100%;
        height: 6px;
        border-radius: 3px;
        background: rgba(255, 255, 255, 0.3);
        outline: none;
        -webkit-appearance: none;
    `;

        scaleSlider.oninput = (e) => {
            guiScale = parseFloat(e.target.value);
            scaleLabel.textContent = `Scale: ${guiScale.toFixed(2)}x`;
            updateGUIScale();
        };

        scaleContainer.appendChild(scaleLabel);
        scaleContainer.appendChild(scaleSlider);

        const logBox = document.createElement('div');
        logBox.id = 'florence-log-box';
        logBox.style.cssText = `
        flex: 1;
        padding: 12px;
        background: rgba(0, 0, 0, 0.3);
        overflow-y: auto;
        font-family: 'Consolas', 'Monaco', monospace;
        font-size: 11px;
        line-height: 1.4;
        max-height: 200px;
    `;

        guiContainer.appendChild(header);
        guiContainer.appendChild(buttonsContainer);
        guiContainer.appendChild(scaleContainer);
        guiContainer.appendChild(logBox);

        header.addEventListener('mousedown', startDragging);
        document.addEventListener('mousemove', drag);
        document.addEventListener('mouseup', stopDragging);

        function startDragging(e) {
            isDragging = true;
            dragOffsetX = e.clientX - guiContainer.offsetLeft;
            dragOffsetY = e.clientY - guiContainer.offsetTop;
            header.style.cursor = 'grabbing';
        }

        function drag(e) {
            if (isDragging) {
                const newX = e.clientX - dragOffsetX;
                const newY = e.clientY - dragOffsetY;
                const maxX = window.innerWidth - guiContainer.offsetWidth;
                const maxY = window.innerHeight - guiContainer.offsetHeight;
                guiContainer.style.left = Math.max(0, Math.min(newX, maxX)) + 'px';
                guiContainer.style.top = Math.max(0, Math.min(newY, maxY)) + 'px';
                guiContainer.style.right = 'auto';
            }
        }

        function stopDragging() {
            isDragging = false;
            header.style.cursor = 'move';
        }

        document.body.appendChild(guiContainer);
        addLogMessage('Florence Automator GUI initialized', 'log');
    }

    function updateLogBox() {
        const logBox = document.getElementById('florence-log-box');
        if (!logBox) return;

        logBox.innerHTML = logMessages.slice(-50).map(msg => {
            const color = msg.type === 'error' ? '#ff6b6b' :
            msg.type === 'warn' ? '#ffd93d' : '#6bcf7f';
            return `<div style="color: ${color}; margin-bottom: 4px;">
                <span style="opacity: 0.7;">[${msg.timestamp}]</span> ${msg.message}
            </div>`;
        }).join('');

        logBox.scrollTop = logBox.scrollHeight;
    }

    function updateGUIScale() {
        const gui = document.getElementById('florence-gui');
        if (!gui) return;

        gui.style.transform = `scale(${guiScale})`;
    }

    function toggleGUI() {
        const gui = document.getElementById('florence-gui');
        if (!gui) {
            createGUI();
            setTimeout(() => {
                document.getElementById('florence-gui').style.display = 'flex';
                guiVisible = true;
            }, 100);
        } else {
            guiVisible = !guiVisible;
            gui.style.display = guiVisible ? 'flex' : 'none';
        }
    }

    //==========================
    // ADD SIGNATURES FUNCTIONS
    //==========================
    // This section contains functions used by Add Signatures feature
    //==========================
    function startAddSignaturesFlow() {
        addLogMessage('Start Add Signatures flow clicked', 'log');
        addLogMessage('Validating if user is on the Request Signatures page', 'log');
        const isOnSignaturePage = validateSignaturePage();
        if (isOnSignaturePage) {
            addLogMessage('User is on the Request Signatures page, opening input GUI', 'log');
            openSignaturesInputGUI();
        } else {
            addLogMessage('User is not on the Request Signatures page, showing warning first', 'warn');
            showWarning('Please navigate to the "Request Signatures" page first.');
        }
    }
    function parseNames(input) {
        addLogMessage('parseNames: start', 'log');
        if (!input || !input.trim()) {
            addLogMessage('parseNames: empty input', 'warn');
            return [];
        }

        const names = [];
        const lines = input.split('\n');
        addLogMessage('parseNames: lines detected = ' + lines.length, 'log');

        for (const line of lines) {
            const lineNames = line.split(',').map(name => name.trim()).filter(name => name);
            addLogMessage('parseNames: line parsed -> ' + JSON.stringify(lineNames), 'log');
            names.push(...lineNames);
        }

        addLogMessage('parseNames: total names parsed = ' + names.length, 'log');
        return names;
    }


    function validateSignaturePage() {
        addLogMessage('validateSignaturePage: checking DOM for Request Signatures UI', 'log');
        const modalContainer = document.querySelector('modal-container');
        const signatureRequests = document.querySelector('documents-signature-requests');
        const headerTitle = document.querySelector('.c-modal-with-tabs_header_title');

        let conditionA = false;
        if (modalContainer && signatureRequests) {
            conditionA = true;
        }

        let conditionB = false;
        if (headerTitle && headerTitle.textContent && headerTitle.textContent.trim() === 'Request Signatures') {
            conditionB = true;
        }

        addLogMessage('validateSignaturePage: conditionA(modal + signatureRequests)=' + conditionA + ', conditionB(headerTitle)=' + conditionB, 'log');
        const result = conditionA || conditionB;
        addLogMessage('validateSignaturePage: result=' + result, 'log');
        return result;
    }

    function processSignatures(names) {
        addLogMessage('processSignatures: start', 'log');
        if (!validateSignaturePage()) {
            addLogMessage('processSignatures: not on Request Signatures page, showing warning', 'warn');
            showWarning('Please navigate to the "Request Signatures" page first.');
            return;
        }

        addLogMessage('processSignatures: proceeding with ' + names.length + ' name(s)', 'log');
        showLoadingGUI(names);
        setTimeout(() => {
            addLogMessage('processSignatures: calling executeSignatureSelection after delay', 'log');
            executeSignatureSelection(names);
        }, 100);
    }

    function showLoadingGUI(names) {
        const modal = document.createElement('div');
        modal.id = 'signatures-loading-modal';
        modal.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.8);
            z-index: 40000;
            display: flex;
            align-items: center;
            justify-content: center;
        `;

        const container = document.createElement('div');
        container.style.cssText = `
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 12px;
            padding: 24px;
            width: 500px;
            max-width: 90%;
            box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3);
            position: relative;
        `;

        const header = document.createElement('div');
        header.style.cssText = `
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
        `;

        const title = document.createElement('h3');
        title.textContent = 'Processing Signatures';
        title.style.cssText = `
            margin: 0;
            color: white;
            font-size: 18px;
            font-weight: 600;
        `;

        const closeButton = document.createElement('button');
        closeButton.innerHTML = '✕';
        closeButton.style.cssText = `
            background: rgba(255, 255, 255, 0.2);
            border: none;
            color: white;
            width: 28px;
            height: 28px;
            border-radius: 50%;
            cursor: pointer;
            font-size: 16px;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.3s ease;
        `;
        closeButton.onmouseover = () => closeButton.style.background = 'rgba(255, 67, 54, 0.8)';
        closeButton.onmouseout = () => closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        closeButton.onclick = () => {
            window.signatureProcessStopped = true;
            document.body.removeChild(modal);
        };

        header.appendChild(title);
        header.appendChild(closeButton);

        const spinner = document.createElement('div');
        spinner.style.cssText = `
            width: 40px;
            height: 40px;
            border: 4px solid rgba(255, 255, 255, 0.3);
            border-top: 4px solid white;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin: 20px auto;
        `;

        const style = document.createElement('style');
        style.textContent = `
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
        `;
        document.head.appendChild(style);

        const statusContainer = document.createElement('div');
        statusContainer.id = 'signature-status-container';
        statusContainer.style.cssText = `
            max-height: 300px;
            overflow-y: auto;
            background: rgba(0, 0, 0, 0.2);
            border-radius: 8px;
            padding: 12px;
        `;

        container.appendChild(header);
        container.appendChild(spinner);
        container.appendChild(statusContainer);
        modal.appendChild(container);

        document.body.appendChild(modal);

        names.forEach(name => {
            const statusDiv = document.createElement('div');
            statusDiv.id = `status-${name.replace(/\s+/g, '-')}`;
            statusDiv.style.cssText = `
                color: white;
                padding: 8px;
                margin: 4px 0;
                background: rgba(255, 255, 255, 0.1);
                border-radius: 4px;
                font-size: 14px;
            `;
            statusDiv.innerHTML = `<strong>${name}:</strong> <span style="color: #ffd93d;">Processing...</span>`;
            statusContainer.appendChild(statusDiv);
        });
    }

    function updateSignatureStatus(name, status, color = '#6bcf7f') {
        addLogMessage('updateSignatureStatus: name="' + name + '" status="' + status + '" color="' + color + '"', 'log');
        const statusDiv = document.getElementById('status-' + name.replace(/\s+/g, '-'));
        if (statusDiv) {
            statusDiv.innerHTML = '<strong>' + name + ':</strong> <span style="color: ' + color + ';">' + status + '</span>';
        } else {
            addLogMessage('updateSignatureStatus: statusDiv not found for "' + name + '"', 'warn');
        }
    }

    function executeSignatureSelection(names) {
        addLogMessage('executeSignatureSelection: start', 'log');
        window.signatureProcessStopped = false;

        try {
            const signersTab = document.querySelector('li.nav-item.active a[role="tab"]');
            if (signersTab && signersTab.textContent && signersTab.textContent.includes('Signers')) {
                addLogMessage('executeSignatureSelection: active Signers tab found, clicking', 'log');
                signersTab.click();
                setTimeout(() => {
                    addLogMessage('executeSignatureSelection: proceeding to selectSigners', 'log');
                    selectSigners(names);
                }, 500);
            } else {
                addLogMessage('executeSignatureSelection: active Signers tab not found, scanning all tabs', 'warn');
                const allTabs = document.querySelectorAll('li.nav-item a[role="tab"]');
                addLogMessage('executeSignatureSelection: total tabs found = ' + allTabs.length, 'log');
                let clicked = false;
                for (const tab of allTabs) {
                    const label = (tab.textContent || '').trim();
                    addLogMessage('executeSignatureSelection: inspecting tab label="' + label + '"', 'log');
                    if (label.includes('Signers')) {
                        addLogMessage('executeSignatureSelection: Signers tab found, clicking', 'log');
                        tab.click();
                        clicked = true;
                        setTimeout(() => {
                            addLogMessage('executeSignatureSelection: proceeding to selectSigners after tab click', 'log');
                            selectSigners(names);
                        }, 500);
                        return;
                    }
                }
                if (!clicked) {
                    addLogMessage('executeSignatureSelection: Could not find Signers tab', 'error');
                    updateSignatureStatus('System', 'Could not find Signers tab', '#ff6b6b');
                }
            }
        } catch (error) {
            addLogMessage('executeSignatureSelection: Error navigating to Signers tab: ' + error, 'error');
            updateSignatureStatus('System', 'Navigation failed', '#ff6b6b');
        }
    }

    function selectSigners(names) {
        addLogMessage('selectSigners: start', 'log');
        if (window.signatureProcessStopped) {
            addLogMessage('selectSigners: process stopped flag detected, aborting', 'warn');
            return;
        }

        try {
            const searchInput = document.getElementById('filtered-select-input');
            if (searchInput) {
                addLogMessage('selectSigners: search input found, clicking to open list', 'log');
                searchInput.click();
                searchInput.focus();
                setTimeout(() => {
                    addLogMessage('selectSigners: calling processSignerList', 'log');
                    processSignerList(names);
                }, 400);
            } else {
                addLogMessage('selectSigners: Search input not found with id="filtered-select-input"', 'error');
                updateSignatureStatus('System', 'Search input not found', '#ff6b6b');
            }
        } catch (error) {
            addLogMessage('selectSigners: Error clicking search input: ' + error, 'error');
            updateSignatureStatus('System', 'Failed to open signer list', '#ff6b6b');
        }
    }
    function closeLoadingGUI() {
        try {
            const modal = document.getElementById('signatures-loading-modal');
            if (modal) {
                document.body.removeChild(modal);
                addLogMessage('closeLoadingGUI: modal closed successfully', 'log');
            } else {
                addLogMessage('closeLoadingGUI: modal not found', 'warn');
            }
        } catch (e) {
            addLogMessage('closeLoadingGUI: error closing modal: ' + e, 'error');
        }
    }
    function processSignerList(names) {
        lastScrollPosition = 0;
        addLogMessage('processSignerList: start (sequential search+select via input)', 'log');
        if (window.signatureProcessStopped) {
            addLogMessage('processSignerList: process stopped flag detected, aborting', 'warn');
            return;
        }

        const searchInput = document.getElementById('filtered-select-input');
        if (!searchInput) {
            addLogMessage('processSignerList: search input not found, cannot proceed', 'error');
            updateSignatureStatus('System', 'Search input not found', '#ff6b6b');
            return;
        }

        let index = 0;

        function processNext() {
            if (window.signatureProcessStopped) {
                addLogMessage('processSignerList: stop flag detected, ending sequence', 'warn');
                return;
            }

            if (index >= names.length) {
                addLogMessage('processSignerList: all names processed, showing completion summary', 'log');
                setTimeout(() => {
                        closeLoadingGUI();
                }, 500);
                return;
            }

            const originalName = names[index];
            addLogMessage('processSignerList: processing "' + originalName + '" (index ' + (index + 1) + ' of ' + names.length + ')', 'log');

            const parts = splitNameParts(originalName);

            try {
                attemptSelectByScrolling(parts, function(success, matchType) {
                    // Clear the search input before processing the next name
                    const searchInput = document.getElementById('filtered-select-input');
                    if (searchInput) {
                        clearSearchInput(searchInput);
                        // Re-open the dropdown for the next search
                        setTimeout(() => {
                            searchInput.click();
                            searchInput.focus();
                            // Restore scroll position after reopening dropdown
                            setTimeout(() => {
                                const viewport = document.querySelector('cdk-virtual-scroll-viewport');
                                if (viewport) {
                                    viewport.scrollTop = lastScrollPosition;
                                    addLogMessage('processNext: restored scroll position to ' + lastScrollPosition, 'log');
                                }
                            }, 50);
                        }, 100);
                    }
                    
                    if (success) {
                        updateSignatureStatus(originalName, matchType, '#6bcf7f');
                        index = index + 1;
                        setTimeout(processNext, 500);
                        return;
                    } else {
                        addLogMessage('processSignerList: not found after scrolling search -> "' + originalName + '"', 'warn');
                        updateSignatureStatus(originalName, 'Not found', '#ff6b6b');
                        index = index + 1;
                        setTimeout(processNext, 500);
                        return;
                    }
                });
            } catch (err) {
                addLogMessage('processSignerList: error while processing "' + originalName + '": ' + err, 'error');
                updateSignatureStatus(originalName, 'Processing failed', '#ff6b6b');
                index = index + 1;
                setTimeout(processNext, 400);
                return;
            }
        }

        processNext();
    }
    function attemptSelectByQuery(searchInput, query, parts, callback) {
        addLogMessage('attemptSelectByQuery: searching for "' + query + '"', 'log');
        
        // Clear the input first
        clearSearchInput(searchInput);
        
        // Type the search query
        typeIntoSearchInput(searchInput, query, function() {
            // Wait for the list to update with search results
            waitForListUpdate(query, 3000, 200, function(items) {
                addLogMessage('attemptSelectByQuery: list updated, searching for candidate', 'log');
                const candidate = findCandidateItem(items, parts);
                
                if (candidate) {
                    addLogMessage('attemptSelectByQuery: found candidate for "' + query + '"', 'log');
                    const clicked = clickCheckboxForItem(candidate, parts.full);
                    if (clicked) {
                        setTimeout(function() {
                            clickAddButtonIfEnabled();
                            callback(true, 'Selected by search');
                        }, 80);
                        return;
                    }
                } else {
                    addLogMessage('attemptSelectByQuery: no candidate found for "' + query + '"', 'warn');
                }
                callback(false, '');
            }, function() {
                addLogMessage('attemptSelectByQuery: timeout waiting for list update', 'warn');
                callback(false, '');
            });
        });
    }

    function attemptSelectByScrolling(parts, callback) {
    addLogMessage('attemptSelectByScrolling: searching for "' + parts.full + '" from position ' + lastScrollPosition, 'log');
    
    const viewport = document.querySelector('cdk-virtual-scroll-viewport');
    if (!viewport) {
        addLogMessage('attemptSelectByScrolling: virtual scroll viewport not found, falling back to search', 'warn');
        // Fallback to search-based approach
        const searchInput = document.getElementById('filtered-select-input');
        if (searchInput && parts.first) {
            attemptSelectByQuery(searchInput, parts.first, parts, callback);
        } else {
            callback(false, '');
        }
        return;
    }

    // Start from the last scroll position instead of 0
    let scrollPosition = lastScrollPosition;
    const SCROLL_STEP = 200;
    const MAX_SCROLLS = 50;
    let scrollCount = 0;

    function scrollAndSearch() {
        if (scrollCount >= MAX_SCROLLS) {
            addLogMessage('attemptSelectByScrolling: max scrolls reached, item not found', 'warn');
            callback(false, '');
            return;
        }

        // Scroll to position
        viewport.scrollTop = scrollPosition;
        
        // Wait for items to render
        setTimeout(() => {
            const snapshot = getListItemsSnapshot();
            const candidate = findCandidateItem(snapshot.items, parts);
            
            if (candidate) {
                addLogMessage('attemptSelectByScrolling: found candidate at scroll position ' + scrollPosition, 'log');
                const clicked = clickCheckboxForItem(candidate, parts.full);
                if (clicked) {
                    setTimeout(function() {
                        clickAddButtonIfEnabled();
                        // Update the last scroll position to where we found the item
                        lastScrollPosition = scrollPosition;
                        addLogMessage('attemptSelectByScrolling: updated lastScrollPosition to ' + lastScrollPosition, 'log');
                        callback(true, 'Selected');
                    }, 80);
                    return;
                }
            }
            
            // Continue scrolling
            scrollPosition += SCROLL_STEP;
            scrollCount++;
            setTimeout(scrollAndSearch, 100);
        }, 200);
    }

    scrollAndSearch();
}

    function splitNameParts(name) {
        const result = { full: name, first: '', last: '' };

        if (!name) {
            addLogMessage('splitNameParts: empty name received', 'warn');
            return result;
        }

        const trimmed = name.trim();
        if (trimmed.indexOf(',') !== -1) {
            const parts = trimmed.split(',').map(x => x.trim()).filter(x => x);
            if (parts.length >= 2) {
                result.last = parts[0];
                result.first = parts[1].split(' ').filter(x => x)[0] || '';
            }
        } else {
            const tokens = trimmed.split(' ').filter(x => x);
            if (tokens.length >= 2) {
                result.first = tokens[0];
                result.last = tokens[tokens.length - 1];
            } else if (tokens.length === 1) {
                result.first = tokens[0];
                result.last = '';
            }
        }

        addLogMessage('splitNameParts: name="' + name + '" -> first="' + result.first + '" last="' + result.last + '"', 'log');
        return result;
    }
    function clearSearchInput(input) {
        try {
            input.focus();
            input.value = '';
            const ev1 = new InputEvent('input', { bubbles: true });
            input.dispatchEvent(ev1);
            const ev2 = new KeyboardEvent('keyup', { bubbles: true, key: 'Backspace' });
            input.dispatchEvent(ev2);
            const ev3 = new Event('change', { bubbles: true });
            input.dispatchEvent(ev3);
            addLogMessage('clearSearchInput: input cleared', 'log');
        } catch (e) {
            addLogMessage('clearSearchInput: error clearing input: ' + e, 'error');
        }
    }

    function waitForListUpdate(expectedFragment, timeoutMs, intervalMs, onSuccess, onTimeout) {
        const start = Date.now();

        function poll() {
            const snapshot = getListItemsSnapshot();
            const names = snapshot.itemsText;

            if (snapshot.items.length > 0) {
                let hasRelated = false;
                for (let i = 0; i < names.length; i++) {
                    const txt = names[i].toLowerCase();
                    if (txt.indexOf(expectedFragment.toLowerCase()) !== -1) {
                        hasRelated = true;
                        break;
                    }
                }

                if (hasRelated) {
                    addLogMessage('waitForListUpdate: related items detected for "' + expectedFragment + '"', 'log');
                    onSuccess(snapshot.items);
                    return;
                }
            }

            if ((Date.now() - start) >= timeoutMs) {
                onTimeout();
                return;
            }

            setTimeout(poll, intervalMs);
        }

        poll();
    }

    function typeIntoSearchInput(input, text, done) {
        try {
            input.focus();
            input.value = '';
            const evClear = new InputEvent('input', { bubbles: true });
            input.dispatchEvent(evClear);
            addLogMessage('typeIntoSearchInput: typing "' + text + '"', 'log');

            let i = 0;

            function typeNext() {
                if (i >= text.length) {
                    const evChange = new Event('change', { bubbles: true });
                    input.dispatchEvent(evChange);
                    addLogMessage('typeIntoSearchInput: finished typing "' + text + '"', 'log');
                    if (typeof done === 'function') {
                        done();
                    }
                    return;
                }

                const ch = text.charAt(i);
                const evKeyDown = new KeyboardEvent('keydown', { bubbles: true, key: ch });
                input.dispatchEvent(evKeyDown);

                input.value = input.value + ch;

                const evInput = new InputEvent('input', { bubbles: true, data: ch });
                input.dispatchEvent(evInput);

                const evKeyUp = new KeyboardEvent('keyup', { bubbles: true, key: ch });
                input.dispatchEvent(evKeyUp);

                i = i + 1;
                setTimeout(typeNext, 30);
            }

            typeNext();
        } catch (e) {
            addLogMessage('typeIntoSearchInput: error typing "' + text + '": ' + e, 'error');
            if (typeof done === 'function') {
                done();
            }
        }
    }
    function getListItemsSnapshot() {
        const container = document.querySelector('.filtered-select__list');
        let items = [];
        if (container) {
            items = Array.from(container.querySelectorAll('.filtered-select__list__item'));
        } else {
            items = Array.from(document.querySelectorAll('.filtered-select__list__item'));
        }
        const itemsText = items.map(getItemDisplayName);
        addLogMessage('getListItemsSnapshot: items now=' + items.length, 'log');
        return { items: items, itemsText: itemsText };
    }

    function getItemDisplayName(li) {
        if (!li) {
            return '';
        }

        const aria = li.getAttribute('aria-label') || '';
        if (aria && aria.trim().length > 0) {
            return aria.trim().replace(/\s+/g, ' ');
        }

        const span = li.querySelector('.filtered-select__list__item__text');
        if (span && span.textContent) {
            return span.textContent.trim().replace(/\s+/g, ' ');
        }

        return (li.textContent || '').trim().replace(/\s+/g, ' ');
    }
    function findCandidateItem(items, parts) {
        const firstLower = (parts.first || '').toLowerCase();
        const lastLower = (parts.last || '').toLowerCase();

        const candidates = [];
        for (let i = 0; i < items.length; i++) {
            const nameText = getItemDisplayName(items[i]);
            const lc = nameText.toLowerCase();

            if (firstLower && lc.indexOf(firstLower) !== -1) {
                candidates.push({ item: items[i], name: nameText });
            }
        }

        if (candidates.length === 1) {
            addLogMessage('findCandidateItem: single candidate on first name -> "' + candidates[0].name + '"', 'log');
            return candidates[0].item;
        }

        if (candidates.length > 1 && lastLower) {
            for (let j = 0; j < candidates.length; j++) {
                const n = candidates[j].name.toLowerCase();
                if (n.indexOf(lastLower) !== -1) {
                    addLogMessage('findCandidateItem: disambiguated by last name -> "' + candidates[j].name + '"', 'log');
                    return candidates[j].item;
                }
            }
        }

        if (!parts.first && lastLower) {
            for (let k = 0; k < items.length; k++) {
                const nt = getItemDisplayName(items[k]).toLowerCase();
                if (nt.indexOf(lastLower) !== -1) {
                    addLogMessage('findCandidateItem: fallback last name matched -> "' + getItemDisplayName(items[k]) + '"', 'log');
                    return items[k];
                }
            }
        }

        addLogMessage('findCandidateItem: no candidate matched with provided parts (first="' + parts.first + '" last="' + parts.last + '")', 'warn');
        return null;
    }

    function clickCheckboxForItem(li, displayTarget) {
        if (!li) {
            addLogMessage('clickCheckboxForItem: no list item provided', 'error');
            return false;
        }

        try {
            const checkbox = li.querySelector('[role="checkbox"]');
            if (!checkbox) {
                addLogMessage('clickCheckboxForItem: checkbox not found for "' + displayTarget + '"', 'error');
                return false;
            }

            const state = checkbox.getAttribute('aria-checked');
            addLogMessage('clickCheckboxForItem: current aria-checked="' + state + '" for "' + displayTarget + '"', 'log');

            if (state === 'false') {
                checkbox.click();
                addLogMessage('clickCheckboxForItem: checkbox clicked for "' + displayTarget + '"', 'log');
                return true;
            } else {
                addLogMessage('clickCheckboxForItem: already selected for "' + displayTarget + '"', 'warn');
                return true;
            }
        } catch (e) {
            addLogMessage('clickCheckboxForItem: error clicking checkbox for "' + displayTarget + '": ' + e, 'error');
            return false;
        }
    }

    function clickAddButtonIfEnabled() {
        try {
            const addBtn = document.querySelector('[data-test="add-button"]');
            if (!addBtn) {
                addLogMessage('clickAddButtonIfEnabled: Add button not found', 'warn');
                return false;
            }

            const disabledAttr = addBtn.getAttribute('disabled');
            if (disabledAttr === null) {
                addBtn.click();
                addLogMessage('clickAddButtonIfEnabled: Add button clicked', 'log');
                return true;
            } else {
                addLogMessage('clickAddButtonIfEnabled: Add button disabled', 'warn');
                return false;
            }
        } catch (e) {
            addLogMessage('clickAddButtonIfEnabled: error clicking Add button: ' + e, 'error');
            return false;
        }
    }
    function fuzzyMatch(name1, name2) {
        const n1 = name1.toLowerCase().replace(/\s+/g, '');
        const n2 = name2.toLowerCase().replace(/\s+/g, '');
        if (n1.includes(n2) || n2.includes(n1)) {
            addLogMessage('fuzzyMatch: substring match true ("' + n1 + '" vs "' + n2 + '")', 'log');
            return true;
        }
        const similarity = calculateSimilarity(n1, n2);
        const result = similarity > 0.7;
        addLogMessage('fuzzyMatch: similarity=' + similarity.toFixed(3) + ' threshold=0.7 result=' + result + ' ("' + n1 + '" vs "' + n2 + '")', 'log');
        return result;
    }

    function calculateSimilarity(str1, str2) {
        const longer = str1.length > str2.length ? str1 : str2;
        const shorter = str1.length > str2.length ? str2 : str1;

        if (longer.length === 0) return 1.0;

        const editDistance = levenshteinDistance(longer, shorter);
        return (longer.length - editDistance) / longer.length;
    }

    function levenshteinDistance(str1, str2) {
        const matrix = [];

        for (let i = 0; i <= str2.length; i++) {
            matrix[i] = [i];
        }

        for (let j = 0; j <= str1.length; j++) {
            matrix[0][j] = j;
        }

        for (let i = 1; i <= str2.length; i++) {
            for (let j = 1; j <= str1.length; j++) {
                if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
                    matrix[i][j] = matrix[i - 1][j - 1];
                } else {
                    matrix[i][j] = Math.min(
                        matrix[i - 1][j - 1] + 1,
                        matrix[i][j - 1] + 1,
                        matrix[i - 1][j] + 1
                    );
                }
            }
        }

        return matrix[str2.length][str1.length];
    }

    function swapNameFormat(name) {
        if (name.includes(',')) {
            const parts = name.split(',').map(part => part.trim());
            const swapped = parts[1] + ' ' + parts[0];
            addLogMessage('swapNameFormat: input="' + name + '" swapped="' + swapped + '"', 'log');
            return swapped;
        } else {
            const parts = name.split(' ');
            if (parts.length >= 2) {
                const swapped = parts[parts.length - 1] + ', ' + parts.slice(0, -1).join(' ');
                addLogMessage('swapNameFormat: input="' + name + '" swapped="' + swapped + '"', 'log');
                return swapped;
            }
        }
        addLogMessage('swapNameFormat: input="' + name + '" unchanged', 'log');
        return name;
    }

    function showCompletionSummary(names) {
        const modal = document.getElementById('signatures-loading-modal');
        if (modal) {
            document.body.removeChild(modal);
        }

        const summaryModal = document.createElement('div');
        summaryModal.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.7);
            z-index: 50000;
            display: flex;
            align-items: center;
            justify-content: center;
        `;

        const container = document.createElement('div');
        container.style.cssText = `
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
            border-radius: 12px;
            padding: 24px;
            width: 450px;
            max-width: 90%;
            box-shadow: 0 15px 35px rgba(0, 0, 0, 0.3);
            position: relative;
        `;

        const header = document.createElement('div');
        header.style.cssText = `
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
        `;

        const title = document.createElement('h3');
        title.textContent = 'Process Complete';
        title.style.cssText = `
            margin: 0;
            color: white;
            font-size: 18px;
            font-weight: 600;
        `;

        const closeButton = document.createElement('button');
        closeButton.innerHTML = '✕';
        closeButton.style.cssText = `
            background: rgba(255, 255, 255, 0.2);
            border: none;
            color: white;
            width: 28px;
            height: 28px;
            border-radius: 50%;
            cursor: pointer;
            font-size: 16px;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.3s ease;
        `;
        closeButton.onmouseover = () => closeButton.style.background = 'rgba(255, 255, 255, 0.3)';
        closeButton.onmouseout = () => closeButton.style.background = 'rgba(255, 255, 255, 0.2)';
        closeButton.onclick = () => document.body.removeChild(summaryModal);

        header.appendChild(title);
        header.appendChild(closeButton);

        const summaryDiv = document.createElement('div');
        summaryDiv.style.cssText = `
            background: rgba(0, 0, 0, 0.2);
            border-radius: 8px;
            padding: 16px;
            max-height: 300px;
            overflow-y: auto;
        `;

        names.forEach(name => {
            const statusDiv = document.createElement('div');
            statusDiv.style.cssText = `
                color: white;
                padding: 6px;
                margin: 2px 0;
                font-size: 14px;
            `;
            const originalStatus = document.getElementById(`status-${name.replace(/\s+/g, '-')}`);
            if (originalStatus) {
                statusDiv.innerHTML = originalStatus.innerHTML;
            }
            summaryDiv.appendChild(statusDiv);
        });

        const okButton = document.createElement('button');
        okButton.textContent = 'OK';
        okButton.style.cssText = `
            background: rgba(255, 255, 255, 0.2);
            border: 2px solid rgba(255, 255, 255, 0.3);
            color: white;
            padding: 10px 20px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            transition: all 0.3s ease;
            margin-top: 20px;
            width: 100%;
        `;
        okButton.onmouseover = () => okButton.style.background = 'rgba(255, 255, 255, 0.3)';
        okButton.onmouseout = () => okButton.style.background = 'rgba(255, 255, 255, 0.2)';
        okButton.onclick = () => document.body.removeChild(summaryModal);

        container.appendChild(header);
        container.appendChild(summaryDiv);
        container.appendChild(okButton);
        summaryModal.appendChild(container);

        document.body.appendChild(summaryModal);
    }

    function init() {
        document.addEventListener('keydown', (e) => {
            if (e.key === 'F2') {
                e.preventDefault();
                toggleGUI();
            }
        });

        console.log('Florence Basic Automator loaded. Press F2 to toggle GUI.');
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", init, { once: true });
    } else {
        init();
    }
})();