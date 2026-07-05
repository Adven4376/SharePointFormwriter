/**
 * Service for communicating with the Request Submission Backend API.
 */
const apiService = {
    // Configurable endpoint. Defaults to Azure Functions local runtime port (7071)
    apiEndpoint: 'http://localhost:7071/api/submit',

    /**
     * Sends the form fields and attachments as a multipart/form-data POST request.
     * @param {Object} fields - Form fields (name, email, department, description).
     * @param {Array<File>} files - Selected attachment files.
     * @returns {Promise<Object>} The API JSON response matching SubmissionResponseDto.
     */
    async submitRequest(fields, files) {
        const formData = new FormData();

        // 1. Append text fields
        formData.append('Name', fields.name.trim());
        formData.append('Email', fields.email.trim());
        formData.append('Department', fields.department.trim());
        formData.append('Description', fields.description.trim());

        // 2. Append attachment files
        if (files && files.length > 0) {
            for (let i = 0; i < files.length; i++) {
                // File name is passed as the third parameter to ensure it is transmitted properly
                formData.append('Attachments', files[i], files[i].name);
            }
        }

        // 3. Make HTTP POST Request
        // IMPORTANT: Do NOT manually set Content-Type header to "multipart/form-data".
        // The fetch API will automatically detect the FormData object, set the Content-Type,
        // and compute the correct boundary parameters.
        try {
            const response = await fetch(this.apiEndpoint, {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (!response.ok) {
                // Extract error message from API response (tolerant to camelCase or PascalCase)
                const errors = result.errors || result.Errors;
                const message = result.message || result.Message;
                const error = result.error || result.Error;

                const errorMessage = errors && errors.length > 0 
                    ? errors.join(', ')
                    : (message || error || `Request failed with status ${response.status}`);
                
                throw new Error(errorMessage);
            }

            return result;
        }
        catch (error) {
            console.error('API Service Error:', error);
            // If the error message is already descriptive, rethrow, otherwise provide a generic network message
            if (error instanceof TypeError && error.message === 'Failed to fetch') {
                throw new Error('Could not connect to the submission server. Please verify if the backend service is running.');
            }
            throw error;
        }
    }
};

// Export API service to be used in ES6 or global scope
if (typeof module !== 'undefined' && module.exports) {
    module.exports = apiService;
} else {
    window.ApiService = apiService;
}
