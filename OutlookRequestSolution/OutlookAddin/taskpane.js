/**
 * Controller script for managing taskpane UI events, Office.js context, and request submissions.
 */

// In-memory array of currently selected files
let selectedFiles = [];

// Wait for Office to initialize
Office.onReady((info) => {
    console.log("Office.js ready. Host: " + info.host);
    
    // 1. Initialise form event listeners
    initializeForm();
    
    // 2. Pre-populate user profile if running inside Outlook
    if (info.host === Office.HostType.Outlook && Office.context.mailbox) {
        const userProfile = Office.context.mailbox.userProfile;
        if (userProfile) {
            document.getElementById("inputName").value = userProfile.displayName || "";
            document.getElementById("inputEmail").value = userProfile.emailAddress || "";
        }
    }
});

/**
 * Attaches event handlers for form fields and actions.
 */
function initializeForm() {
    const form = document.getElementById("submissionForm");
    const fileInput = document.getElementById("fileInput");
    const fileUploadZone = document.getElementById("fileUploadZone");
    const successCloseBtn = document.getElementById("successCloseBtn");
    const errorCloseBtn = document.getElementById("errorCloseBtn");

    // Form Submission
    form.addEventListener("submit", handleFormSubmit);

    // File Input selection change
    fileInput.addEventListener("change", (e) => {
        handleFileSelection(e.target.files);
        // Clear input value so same file can be selected again
        fileInput.value = "";
    });

    // Drag and Drop Zone events
    fileUploadZone.addEventListener("click", () => fileInput.click());
    
    fileUploadZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        fileUploadZone.classList.add("dragover");
    });

    fileUploadZone.addEventListener("dragleave", () => {
        fileUploadZone.classList.remove("dragover");
    });

    fileUploadZone.addEventListener("drop", (e) => {
        e.preventDefault();
        fileUploadZone.classList.remove("dragover");
        if (e.dataTransfer.files) {
            handleFileSelection(e.dataTransfer.files);
        }
    });

    // Close Dialog overlays
    successCloseBtn.addEventListener("click", () => {
        document.getElementById("successOverlay").style.display = "none";
        resetForm();
    });

    errorCloseBtn.addEventListener("click", () => {
        document.getElementById("errorOverlay").style.display = "none";
    });
}

/**
 * Adds selected files to the list, filtering duplicates and checking total limits.
 * @param {FileList} fileList 
 */
function handleFileSelection(fileList) {
    if (!fileList || fileList.length === 0) return;

    const maxAttachments = 5;
    const errors = [];

    for (let i = 0; i < fileList.length; i++) {
        const file = fileList[i];
        
        // Prevent adding exact same file name twice
        if (selectedFiles.some(f => f.name === file.name)) {
            continue;
        }

        // Enforce maximum file limit during selection check
        if (selectedFiles.length >= maxAttachments) {
            errors.push(`You can only upload a maximum of ${maxAttachments} attachments.`);
            break;
        }

        selectedFiles.push(file);
    }

    if (errors.length > 0) {
        showErrors(errors);
    }

    renderFileList();
}

/**
 * Renders list items for each selected file in the UI.
 */
function renderFileList() {
    const listContainer = document.getElementById("fileList");
    listContainer.innerHTML = "";

    selectedFiles.forEach((file, index) => {
        const li = document.createElement("li");
        li.className = "file-item";

        // File extension representation
        const ext = file.name.substring(file.name.lastIndexOf('.') + 1).toUpperCase() || 'FILE';
        const displaySize = (file.size / (1024 * 1024)).toFixed(2) + " MB";

        li.innerHTML = `
            <div class="file-info">
                <div class="file-icon-badge">${ext}</div>
                <div class="file-details">
                    <span class="file-name" title="${file.name}">${file.name}</span>
                    <span class="file-size">${displaySize}</span>
                </div>
            </div>
            <button class="file-remove-btn" type="button" data-index="${index}">&times;</button>
        `;

        // Attach event listener for the remove button
        li.querySelector(".file-remove-btn").addEventListener("click", (e) => {
            const idx = parseInt(e.target.getAttribute("data-index"), 10);
            selectedFiles.splice(idx, 1);
            renderFileList();
        });

        listContainer.appendChild(li);
    });
}

/**
 * Gathers inputs, validates, and POSTs the request using the ApiService.
 * @param {Event} event 
 */
async function handleFormSubmit(event) {
    event.preventDefault();

    const formData = {
        name: document.getElementById("inputName").value,
        email: document.getElementById("inputEmail").value,
        department: document.getElementById("selectDepartment").value,
        description: document.getElementById("textareaDescription").value
    };

    // 1. Client-Side Validation
    const validationResult = window.FormValidator.validate(formData, selectedFiles);
    if (!validationResult.isValid) {
        showErrors(validationResult.errors);
        return;
    }

    // 2. Display Processing Overlay
    showLoading(true);

    try {
        // 3. Make HTTP request via ApiService
        const response = await window.ApiService.submitRequest(formData, selectedFiles);
        
        // Hide loading
        showLoading(false);

        // 4. Show Success Dialog
        const isSuccess = response.success !== undefined ? response.success : response.Success;
        if (response && isSuccess) {
            const submissionId = response.submissionId || response.SubmissionId || "-";
            const folderUrl = response.folderUrl || response.FolderUrl;

            document.getElementById("valTrackingId").innerText = submissionId;
            
            const linkFolder = document.getElementById("linkFolder");
            if (folderUrl) {
                linkFolder.href = folderUrl;
                linkFolder.style.display = "inline";
            } else {
                linkFolder.style.display = "none";
            }
            
            document.getElementById("successOverlay").style.display = "flex";
        } else {
            // Re-map response errors if API returned custom errors
            const errors = response.errors || response.Errors || ["Submission failed without specific error messages."];
            showErrors(errors);
        }
    }
    catch (error) {
        showLoading(false);
        showErrors([error.message || "An unexpected network error occurred."]);
    }
}

/**
 * Toggles the visibility of the loading spinner.
 * @param {boolean} isLoading 
 */
function showLoading(isLoading) {
    document.getElementById("loadingOverlay").style.display = isLoading ? "flex" : "none";
    document.getElementById("submitButton").disabled = isLoading;
}

/**
 * Populates and displays the error modal.
 * @param {Array<string>} errors 
 */
function showErrors(errors) {
    const errorList = document.getElementById("errorList");
    errorList.innerHTML = "";
    
    errors.forEach(err => {
        const li = document.createElement("li");
        li.innerText = err;
        errorList.appendChild(li);
    });

    document.getElementById("errorOverlay").style.display = "flex";
}

/**
 * Resets form values and clears attachments.
 */
function resetForm() {
    document.getElementById("selectDepartment").value = "";
    document.getElementById("textareaDescription").value = "";
    selectedFiles = [];
    renderFileList();
}
