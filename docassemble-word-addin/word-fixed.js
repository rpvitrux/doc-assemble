// Fixed DocAssemble Word Add-in JavaScript
// Replaces iframe authentication with Office Dialog API

// Original DocAssemble variables
var daServer = null;
var daFullServer = null;
var attrs_showing = {};

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
}

function showProjectsInterface() {
    document.getElementById('daServerDiv').style.display = 'none';
    document.getElementById('iframeDiv').style.display = 'none';
    document.getElementById('projectsDiv').style.display = 'block';
    document.getElementById('variablesDiv').style.display = 'block';
    document.getElementById('changeServerDiv').style.display = 'block';

    // Load projects and variables
    loadProjects();
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

// Export functions for use by other scripts
window.showAuthenticationDialog = showAuthenticationDialog;
window.attrs_showing = attrs_showing;