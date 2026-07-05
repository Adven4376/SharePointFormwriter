using System;
using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using RequestSubmissionFunctionApp.Configuration;
using RequestSubmissionFunctionApp.Interfaces;
using RequestSubmissionFunctionApp.Services;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureAppConfiguration((hostingContext, configBuilder) =>
    {
        configBuilder.AddEnvironmentVariables();
    })
    .ConfigureServices((hostContext, services) =>
    {
        var configuration = hostContext.Configuration;

        // 1. Register Application Insights telemetry for Isolated Worker
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();

        // 2. Bind and Register Configuration Settings
        var sharePointSettings = new SharePointSettings();
        configuration.GetSection("SharePoint").Bind(sharePointSettings);
        services.AddSingleton(sharePointSettings);

        var appSettings = new ApplicationSettings();
        configuration.GetSection("Application").Bind(appSettings);
        services.AddSingleton(appSettings);

        // 3. Register GraphServiceClient with secure Authentication
        services.AddSingleton<GraphServiceClient>(sp =>
        {
            var logger = sp.GetRequiredService<ILogger<Program>>();
            
            // Read authentication parameters from SharePointSettings as defined in the diagram
            string tenantId = sharePointSettings.TenantId;
            string clientId = sharePointSettings.ClientId;
            string clientSecret = sharePointSettings.ClientSecret;
            string keyVaultUrl = configuration["AzureAd__KeyVaultUrl"] ?? string.Empty;

            // If ClientSecret is empty but KeyVaultUrl is specified, fetch client secret from Key Vault
            if (string.IsNullOrEmpty(clientSecret) && !string.IsNullOrEmpty(keyVaultUrl))
            {
                try
                {
                    logger.LogInformation("Retrieving SharePoint Client Secret from Azure Key Vault: {VaultUrl}", keyVaultUrl);
                    var secretClient = new SecretClient(new Uri(keyVaultUrl), new DefaultAzureCredential());
                    KeyVaultSecret secret = secretClient.GetSecret("SharePoint--ClientSecret");
                    clientSecret = secret.Value;
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "Failed to retrieve Client Secret from Azure Key Vault. Falling back to default credentials.");
                }
            }

            // Authentication initialization
            if (!string.IsNullOrEmpty(tenantId) && 
                !string.IsNullOrEmpty(clientId) && 
                !string.IsNullOrEmpty(clientSecret))
            {
                logger.LogInformation("Initializing GraphServiceClient using ClientSecretCredential for Client {ClientId}.", clientId);
                var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
                return new GraphServiceClient(credential);
            }
            else
            {
                logger.LogInformation("Initializing GraphServiceClient using DefaultAzureCredential (e.g., Managed Identity).");
                var credential = new DefaultAzureCredential();
                return new GraphServiceClient(credential);
            }
        });

        // 4. Register Core Services and Interfaces
        services.AddTransient<IValidationService, ValidationService>();
        services.AddTransient<ISharePointDocumentService, SharePointDocumentService>();
        services.AddTransient<ISharePointListService, SharePointListService>();
        services.AddTransient<ISubmissionService, SubmissionService>();
    })
    .Build();

host.Run();
