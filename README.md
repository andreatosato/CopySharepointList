# Copy Sharepoint Lists
Questo software serve per copiare il contenuto di un sito SharePoint (SitoMaster) in più siti SharePoint.

Per configurare il sito Master è necessario impostare la proprietà **ListConfig:SiteMasterId** su Azure la proprità diventa "ListConfig__SiteMasterId" dove i : diventano __ .

Il sito master avrà una serie di liste che dovranno essere copiate. Gli id delle liste vanno inserite nella proprietà **ListConfig:ListsToCopy**.

## Configuration
Questa è la configurazione del file **local.settings.json** che può essere usata solo per l'ambiente di sviluppo.
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
![AzureConfig](https://github.com/andreatosato/CopySharepointList/blob/main/docs/images/AzureConfig.png)

## Registrazione Autorizzazione
Creare una app registration
![CreareApp](https://github.com/andreatosato/CopySharepointList/blob/main/docs/images/2021-07-03_11h18_23.png)
Seleziona le impostazioni del tenant (single tenant). Non registrare nessuna applicazione web
![SigleTenant](https://github.com/andreatosato/CopySharepointList/blob/main/docs/images/2021-07-03_11h20_14.png)
Autorizza Graph
![AutorizzaGraph](https://github.com/andreatosato/CopySharepointList/blob/main/docs/images/2021-07-03_11h21_09.png)
Site full control
![SiteFullControll](https://github.com/andreatosato/CopySharepointList/blob/main/docs/images/2021-07-03_11h22_07.png)
Dai i consensi amministrativi
![AdminConsent](https://github.com/andreatosato/CopySharepointList/blob/main/docs/images/2021-07-03_11h23_33.png)
Crea un client secret
![ClientSecret](https://github.com/andreatosato/CopySharepointList/blob/main/docs/images/2021-07-03_18h03_47.png)
Recupera i secret da aggiungere in **AuthConfig:ClientSecret**. Nella schermata principale dell'app registration, recuperare il client id e il tenant id.
![GetClientSecret](https://github.com/andreatosato/CopySharepointList/blob/main/docs/images/2021-07-03_18h04_44.png)

## Lista di SharePoint
Le liste da copiare di SharePoint dovranno essere inserite nel campo **ListConfig:ListsToCopy** che al suo interno avrà i nomi delle liste.

Le liste dovranno essere separate con il carattere **;**

## Campi di SharePoint
Nelle liste di SharePoint, ogni campo, ha due tipi di proprietà
- **internal field** che dovrà essere inserito nella lista **ListConfig:FieldToCopy** 
- **display field** che dovrà essere inserito nella lista **ListConfig:DisplayNameToCopy**

Poichè devono essere configurate più liste da copiare, il formato dei campo dovrà essere come segue:

"campo A lista 1,campo B lista 1;campo A lista 2,campo B lista 2"
il separatore di campo è la **,** mentre il separatore di lista è **;**.

Il campo in alto a destra è l'**internal field** mentre "Nome colonna" è il campo **display field**

![FieldSpace](https://github.com/andreatosato/CopySharepointList/blob/main/docs/images/FieldSpace.png)

## Lista e Campo
E' fondamentale impostare i campi e le liste nello stesso ordine, quindi se dichiaro "Lista Clienti" come prima lista, poi i campi dovranno iniziare con la stessa lista.
