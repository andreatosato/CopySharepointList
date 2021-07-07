# CopySharepointList

## Configuration

```json
{
    "IsEncrypted": false,
    "Values": {
        "AzureWebJobsStorage": "UseDevelopmentStorage=true",
        "FUNCTIONS_WORKER_RUNTIME": "dotnet-isolated",
        "Cron": "0 0/5 * * * *", // Impostare la cron expression desiderata.
        "AuthConfig:ClientId": "ClientId da AAD",
        "AuthConfig:ClientSecret": "Client Secrets da AAD",
        "AuthConfig:TenantId": "Tenant Id di AAD",
        "ListConfig:SiteMasterId": "Site Master ID",
        "ListConfig:ListsToCopy": "Lista Clienti;Lista Fornitori;Lista Dipendenti;Lista Brand",
        "ListConfig:FieldToCopy": "Codice,Ragione_x0020_Sociale;Codice,RagioneSociale;NomeeCognome;Brand",
        "ListConfig:DisplayNameToCopy": "Codice,Ragione Sociale;Codice,Ragione Sociale;Nome e Cognome;Brand",
        "ListConfig:SitesToCopy": "SiteIdOne;SiteIdTwo"
    }
}
```
