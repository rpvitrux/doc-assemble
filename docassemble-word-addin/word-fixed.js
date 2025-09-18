// Enhanced DocAssemble Word Add-in JavaScript
// Features: Dialog API Auth + Template Scanner + Clause Library

// Original DocAssemble variables
var daServer = null;
var daFullServer = null;
var attrs_showing = {};
var foundVariables = [];

// Legal clause library
var legalClauses = {
    liability: "LIMITATION OF LIABILITY: In no event shall either party be liable for any indirect, incidental, special, consequential or punitive damages, including without limitation, loss of profits, data, use, goodwill, or other intangible losses, resulting from the use of this service.",
    confidentiality: "CONFIDENTIALITY: Each party acknowledges that it may have access to certain confidential information of the other party. Each party agrees to maintain confidentiality and not disclose such information to third parties without prior written consent.",
    termination: "TERMINATION: This agreement may be terminated by either party with thirty (30) days written notice. Upon termination, all rights and obligations shall cease except those that by their nature should survive termination.",
    dispute: "DISPUTE RESOLUTION: Any disputes arising under this agreement shall be resolved through binding arbitration in accordance with the rules of the American Arbitration Association in the jurisdiction where this agreement was executed.",
    indemnification: "INDEMNIFICATION: Each party shall indemnify, defend and hold harmless the other party from and against any and all claims, damages, losses, costs and expenses arising out of or resulting from their breach of this agreement.",
    force_majeure: "FORCE MAJEURE: Neither party shall be liable for any failure or delay in performance under this agreement which is due to fire, flood, earthquake, elements of nature or acts of God, acts of war, terrorism, riots, civil disorders, rebellions or other similar causes beyond the reasonable control of such party."
};

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log('DocAssemble Word Add-in loaded');
        initializeAddIn();
    }
});

function initializeAddIn() {
    // Check if we already have server connection
    daServer = Cookies.get("daServer");

    if (daServer) {
        console.log('Found existing server connection:', daServer);
        daFullServer = daServer;
        showProjectsInterface();
    } else {
        console.log('No server connection found, showing server form');
        showServerForm();
    }

    // Set up event listeners for server connection
    setupEventListeners();
}

function showServerForm() {
    document.getElementById('daServerDiv').style.display = 'block';
    document.getElementById('iframeDiv').style.display = 'none';
    document.getElementById('projectsDiv').style.display = 'none';
    document.getElementById('variablesDiv').style.display = 'none';
    document.getElementById('changeServerDiv').style.display = 'none';
    hideTemplateUI();
    hideClauseUI();
}

function hideTemplateUI() {
    const templateElements = document.querySelectorAll('.template-ui');
    templateElements.forEach(el => el.style.display = 'none');
}

function hideClauseUI() {
    const clauseElements = document.querySelectorAll('.clause-ui');
    clauseElements.forEach(el => el.style.display = 'none');
}

function showProjectsInterface() {
    document.getElementById('daServerDiv').style.display = 'none';
    document.getElementById('iframeDiv').style.display = 'none';
    document.getElementById('projectsDiv').style.display = 'block';
    document.getElementById('variablesDiv').style.display = 'block';
    document.getElementById('changeServerDiv').style.display = 'block';

    // Load projects and variables
    loadProjects();

    // Initialize enhanced features after a short delay
    setTimeout(initializeEnhancedFeatures, 500);
}

function showAuthenticationDialog() {
    console.log('Starting Dialog API authentication');

    const serverUrl = document.getElementById('daServer').value || 'https://localhost:8444';
    const dialogUrl = `${serverUrl}/office-auth-dialog`;

    console.log('Opening authentication dialog:', dialogUrl);

    Office.context.ui.displayDialogAsync(dialogUrl, {
        height: 70,
        width: 70,
        displayInIFrame: false
    }, function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error('Failed to open authentication dialog:', asyncResult.error);
            alert('Failed to open authentication dialog: ' + asyncResult.error.message);
        } else {
            console.log('Authentication dialog opened successfully');
            const dialog = asyncResult.value;

            // Listen for messages from the dialog
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                console.log('Received message from dialog:', arg.message);

                try {
                    const result = JSON.parse(arg.message);

                    if (result.success) {
                        // Authentication successful
                        console.log('Authentication successful!');
                        daServer = result.server || document.getElementById('daServer').value;
                        daFullServer = daServer;

                        // Save server connection
                        Cookies.set("daServer", daServer, {expires: 30});

                        // Close dialog and show projects
                        dialog.close();
                        showProjectsInterface();

                    } else {
                        // Authentication failed
                        console.error('Authentication failed:', result.error);
                        alert('Authentication failed: ' + (result.error || 'Unknown error'));
                        dialog.close();
                    }
                } catch (e) {
                    console.error('Error parsing dialog message:', e);
                    alert('Error processing authentication response');
                    dialog.close();
                }
            });

            // Handle dialog closed event
            dialog.addEventHandler(Office.EventType.DialogEventReceived, function(arg) {
                console.log('Dialog event received:', arg);
                if (arg.error === 12006) {
                    // User closed the dialog
                    console.log('User closed authentication dialog');
                } else {
                    console.error('Dialog error:', arg.error);
                }
            });
        }
    });
}

function setupEventListeners() {
    // Server connect button
    const serverConnectBtn = document.getElementById('serverConnect');
    if (serverConnectBtn) {
        serverConnectBtn.addEventListener('click', function() {
            console.log('Server connect button clicked');
            showAuthenticationDialog();
        });
    }

    // Change server button
    const serverChangeBtn = document.getElementById('serverChange');
    if (serverChangeBtn) {
        serverChangeBtn.addEventListener('click', function() {
            console.log('Change server button clicked');
            // Clear stored server
            Cookies.remove("daServer");
            daServer = null;
            daFullServer = null;
            showServerForm();
        });
    }

    // Set default server value
    const serverInput = document.getElementById('daServer');
    if (serverInput && !serverInput.value) {
        serverInput.value = 'https://localhost:8444';
    }
}

function loadProjects() {
    // Mock projects loading - in real implementation this would fetch from server
    console.log('Loading projects...');

    const projectSelect = document.getElementById('daProjects');
    if (projectSelect) {
        projectSelect.innerHTML = '<option value="default">Default Project</option>';

        // Load variables for default project
        loadVariables();
    }
}

function loadVariables() {
    // Mock variables loading - in real implementation this would fetch from server
    console.log('Loading variables...');

    const variableSelect = document.getElementById('daVariables');
    if (variableSelect) {
        variableSelect.innerHTML = '<option value="interview">Select Interview</option>';
    }

    const playgroundTable = document.getElementById('daplaygroundtable');
    if (playgroundTable) {
        playgroundTable.innerHTML = '<tr><td>Connected to DocAssemble server</td></tr>';
    }
}

// WORD DOCUMENT SCANNING FUNCTIONALITY

function scanDocumentForVariables() {
    return Word.run(function (context) {
        console.log('Scanning document for {{ variable }} patterns...');

        // Get the document body
        var body = context.document.body;
        context.load(body, 'text');

        return context.sync().then(function() {
            const documentText = body.text;
            console.log('Document text length:', documentText.length);

            // Find all {{ variable }} patterns
            const variableRegex = /\{\{\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*\}\}/g;
            const matches = [];
            let match;

            while ((match = variableRegex.exec(documentText)) !== null) {
                const variableName = match[1];
                if (!matches.some(m => m.variable === variableName)) {
                    matches.push({
                        variable: variableName,
                        fullMatch: match[0],
                        question: generateQuestionFromVariable(variableName)
                    });
                }
            }

            foundVariables = matches;
            console.log('Found variables:', foundVariables);

            return foundVariables;
        });
    }).catch(function (error) {
        console.error('Error scanning document:', error);
        throw error;
    });
}

function generateQuestionFromVariable(variableName) {
    // Convert snake_case to readable question
    const readable = variableName
        .replace(/_/g, ' ')
        .replace(/([a-z])([A-Z])/g, '$1 $2')
        .toLowerCase();

    return `What is the ${readable}?`;
}

function createInterviewFromTemplate() {
    console.log('Creating interview from template...');

    scanDocumentForVariables().then(function(variables) {
        if (variables.length === 0) {
            alert('No {{ variables }} found in the document.\n\nPlease add some variables like {{ client_name }} or {{ address }} to your document first.');
            return;
        }

        // Show results in the existing interface
        displayFoundVariables(variables);

        // Generate YAML
        const yaml = generateInterviewYAML(variables);
        console.log('Generated YAML:', yaml);

        // For now, show success message
        setTimeout(() => {
            alert(`Interview template created successfully!\n\nFound ${variables.length} variables:\n` +
                  variables.map(v => `• {{ ${v.variable} }}`).join('\n') +
                  '\n\nYAML has been generated and can be saved to DocAssemble.');
        }, 1000);

    }).catch(function(error) {
        console.error('Error creating interview:', error);
        alert('Error scanning document: ' + error.message);
    });
}

function displayFoundVariables(variables) {
    // Update the playground table with found variables
    const table = document.getElementById('daplaygroundtable');
    if (table) {
        let html = '<tr><th>Variable</th><th>Question</th></tr>';
        variables.forEach(v => {
            html += `<tr><td><code>{{ ${v.variable} }}</code></td><td>${v.question}</td></tr>`;
        });
        table.innerHTML = html;
    }
}

function generateInterviewYAML(variables) {
    let yaml = `---\nquestion: Please provide the following information\nfields:\n`;

    variables.forEach(v => {
        yaml += `  - ${v.question.replace('What is the ', '').replace('?', '')}: ${v.variable}\n`;
    });

    yaml += `---\nquestion: |\n  Your document has been created with the following information:\n\n`;
    variables.forEach(v => {
        yaml += `  - **${v.question.replace('What is the ', '').replace('?', '')}**: {{ ${v.variable} }}\n`;
    });

    return yaml;
}

// CLAUSE INSERTION FUNCTIONALITY

function insertClauseIntoDocument(clauseKey) {
    if (!legalClauses[clauseKey]) {
        console.error('Clause not found:', clauseKey);
        return Promise.reject(new Error('Clause not found'));
    }

    return Word.run(function (context) {
        console.log('Inserting clause:', clauseKey);

        // Get current selection or end of document
        var selection = context.document.getSelection();

        // Insert the clause text
        const clauseText = '\n\n' + legalClauses[clauseKey] + '\n\n';
        selection.insertText(clauseText, Word.InsertLocation.end);

        return context.sync().then(function() {
            console.log('Clause inserted successfully');
        });
    }).catch(function (error) {
        console.error('Error inserting clause:', error);
        throw error;
    });
}

// ENHANCED INITIALIZATION

function initializeEnhancedFeatures() {
    // Create tabbed interface that wraps the existing DocAssemble content
    createTabbedInterface();
}

function createTabbedInterface() {
    // Find the main content area
    const projectsDiv = document.getElementById('projectsDiv');
    const variablesDiv = document.getElementById('variablesDiv');
    const changeServerDiv = document.getElementById('changeServerDiv');

    if (!projectsDiv || !variablesDiv) return;

    // Create tab container
    const tabContainer = document.createElement('div');
    tabContainer.className = 'docassemble-tabs';
    tabContainer.innerHTML = `
        <style>
            .docassemble-tabs {
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            }
            .tab-navigation {
                display: flex;
                border-bottom: 2px solid #e0e0e0;
                margin-bottom: 20px;
                background: #f8f9fa;
                border-radius: 6px 6px 0 0;
            }
            .tab-button {
                flex: 1;
                padding: 12px 8px;
                border: none;
                background: #f8f9fa;
                color: #666;
                cursor: pointer;
                font-size: 13px;
                font-weight: 500;
                border-bottom: 3px solid transparent;
                transition: all 0.2s;
                text-align: center;
            }
            .tab-button:hover {
                background: #e9ecef;
                color: #333;
            }
            .tab-button.active {
                background: white;
                color: #0078d4;
                border-bottom-color: #0078d4;
            }
            .tab-content {
                display: none;
                min-height: 400px;
            }
            .tab-content.active {
                display: block;
            }
            .template-section, .clause-section {
                padding: 20px;
                background: white;
                border-radius: 6px;
                margin-bottom: 20px;
            }
            .template-button, .clause-insert-btn {
                background-color: #0078d4;
                color: white;
                border: none;
                padding: 12px 20px;
                border-radius: 4px;
                font-size: 14px;
                cursor: pointer;
                width: 100%;
                margin-bottom: 10px;
            }
            .template-button:hover, .clause-insert-btn:hover {
                background-color: #106ebe;
            }
            .clause-selector {
                width: 100%;
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                margin-bottom: 15px;
                font-size: 13px;
            }
            .help-text {
                font-size: 12px;
                color: #666;
                text-align: center;
                margin-top: 10px;
            }
        </style>

        <div class="tab-navigation">
            <button class="tab-button active" onclick="switchTab('variables', this)">Build Template</button>
            <button class="tab-button" onclick="switchTab('clauses', this)">Insert Clause</button>
        </div>

        <div class="tab-content active" id="variables-tab">
            <!-- Original DocAssemble content will be moved here -->
            <div class="template-section" style="margin-top: 20px;">
                <button class="template-button" onclick="createInterviewFromTemplate()">
                    Scan Document
                </button>
                <p class="help-text">Scan document for {{ variables }} and show results</p>
            </div>
        </div>

        <div class="tab-content" id="clauses-tab">
            <div class="clause-section">
                <h4 style="margin: 0 0 15px 0; color: #333;">Legal Clause Library</h4>
                <select class="clause-selector" id="clauseSelector">
                    <option value="">Select a clause to insert...</option>
                    <option value="liability">Liability Limitation</option>
                    <option value="confidentiality">Confidentiality</option>
                    <option value="termination">Termination</option>
                    <option value="dispute">Dispute Resolution</option>
                    <option value="indemnification">Indemnification</option>
                    <option value="force_majeure">Force Majeure</option>
                </select>
                <button class="clause-insert-btn" onclick="insertSelectedClause()">
                    Insert Selected Clause
                </button>
            </div>
        </div>
    `;

    // Insert the tab container before the existing content
    projectsDiv.parentNode.insertBefore(tabContainer, projectsDiv);

    // Move existing DocAssemble content into the Variables tab
    const variablesTab = document.getElementById('variables-tab');

    // Move projects div, variables div, and change server div into the Variables tab
    variablesTab.appendChild(projectsDiv);
    variablesTab.appendChild(variablesDiv);
    if (changeServerDiv) {
        variablesTab.appendChild(changeServerDiv);
    }

    // Hide the interview selector in the Variables tab
    const interviewSelector = document.getElementById('daVariables');
    if (interviewSelector && interviewSelector.parentElement) {
        interviewSelector.parentElement.parentElement.style.display = 'none'; // Hide the entire "Interview:" line
    }

    // Add global tab switching function
    window.switchTab = function(tabName, buttonElement) {
        // Hide all tab contents
        const tabContents = document.querySelectorAll('.tab-content');
        tabContents.forEach(tab => tab.classList.remove('active'));

        // Remove active class from all tab buttons
        const tabButtons = document.querySelectorAll('.tab-button');
        tabButtons.forEach(button => button.classList.remove('active'));

        // Show selected tab content
        document.getElementById(tabName + '-tab').classList.add('active');

        // Add active class to clicked tab button
        buttonElement.classList.add('active');
    };

    // Add global clause insertion function
    window.insertSelectedClause = function() {
        const selector = document.getElementById('clauseSelector');
        const selectedClause = selector.value;

        if (!selectedClause) {
            alert('Please select a clause to insert.');
            return;
        }

        insertClauseIntoDocument(selectedClause).then(function() {
            alert('Clause inserted successfully!');
            selector.value = ''; // Reset selection
        }).catch(function(error) {
            alert('Error inserting clause: ' + error.message);
        });
    };
}

// Export functions for use by other scripts
window.showAuthenticationDialog = showAuthenticationDialog;
window.attrs_showing = attrs_showing;
window.createInterviewFromTemplate = createInterviewFromTemplate;
window.insertClauseIntoDocument = insertClauseIntoDocument;
window.scanDocumentForVariables = scanDocumentForVariables;