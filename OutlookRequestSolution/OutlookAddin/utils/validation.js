/**
 * Client-side validation utility for the Outlook Request Submission form.
 */
const validation = {
    MAX_FILE_SIZE_BYTES: 10 * 1024 * 1024, // 10MB
    MAX_ATTACHMENTS: 5,
    ALLOWED_EXTENSIONS: ['.pdf', '.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt', '.png', '.jpg', '.jpeg', '.txt', '.zip'],

    /**
     * Validates the request form and selected attachments.
     * @param {Object} formData - Object containing form fields.
     * @param {Array<File>} files - Array of File objects selected by the user.
     * @returns {Object} { isValid: boolean, errors: Array<string> }
     */
    validate(formData, files) {
        const errors = [];

        // 1. Field validation
        if (!formData.name || formData.name.trim() === '') {
            errors.push('Name is required.');
        }

        if (!formData.email || formData.email.trim() === '') {
            errors.push('Email is required.');
        } else if (!this.isValidEmail(formData.email)) {
            errors.push('Please enter a valid email address.');
        }

        if (!formData.department || formData.department.trim() === '') {
            errors.push('Department is required.');
        }

        if (!formData.description || formData.description.trim() === '') {
            errors.push('Description is required.');
        }

        // 2. Attachment count validation
        if (files && files.length > this.MAX_ATTACHMENTS) {
            errors.push(`You can upload a maximum of ${this.MAX_ATTACHMENTS} attachments. Current count: ${files.length}.`);
        }

        // 3. File validations (size and extension)
        if (files && files.length > 0) {
            for (let i = 0; i < files.length; i++) {
                const file = files[i];
                const fileName = file.name;
                const fileSize = file.size;
                const fileExt = this.getFileExtension(fileName);

                // Check size
                if (fileSize <= 0) {
                    errors.push(`File "${fileName}" is empty.`);
                } else if (fileSize > this.MAX_FILE_SIZE_BYTES) {
                    errors.push(`File "${fileName}" exceeds the 10MB size limit (${(fileSize / (1024 * 1024)).toFixed(2)}MB).`);
                }

                // Check extension
                if (!fileExt || !this.ALLOWED_EXTENSIONS.includes(fileExt.toLowerCase())) {
                    errors.push(`File "${fileName}" has an unsupported format. Allowed formats: ${this.ALLOWED_EXTENSIONS.join(', ')}.`);
                }
            }
        }

        return {
            isValid: errors.length === 0,
            errors: errors
        };
    },

    /**
     * Validates email format.
     * @param {string} email 
     * @returns {boolean}
     */
    isValidEmail(email) {
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        return emailRegex.test(email);
    },

    /**
     * Extracts extension including dot from filename.
     * @param {string} filename 
     * @returns {string}
     */
    getFileExtension(filename) {
        const index = filename.lastIndexOf('.');
        return index !== -1 ? filename.substring(index) : '';
    }
};

// Export validator to be used in ES6 or global scope
if (typeof module !== 'undefined' && module.exports) {
    module.exports = validation;
} else {
    window.FormValidator = validation;
}
