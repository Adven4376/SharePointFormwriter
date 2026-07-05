namespace RequestSubmissionFunctionApp.Configuration
{
    /// <summary>
    /// Configuration options for Microsoft Entra ID and Azure Key Vault integration.
    /// </summary>
    public class AzureAdSettings
    {
        /// <summary>
        /// The Microsoft Entra Tenant ID.
        /// </summary>
        public string TenantId { get; set; } = string.Empty;

        /// <summary>
        /// The Client (Application) ID of the registered Entra Application.
        /// </summary>
        public string ClientId { get; set; } = string.Empty;

        /// <summary>
        /// The Client Secret. Only populated locally or retrieved at runtime from Key Vault.
        /// </summary>
        public string ClientSecret { get; set; } = string.Empty;

        /// <summary>
        /// The Azure Key Vault URL where secrets are stored.
        /// </summary>
        public string KeyVaultUrl { get; set; } = string.Empty;
    }
}
