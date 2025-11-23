@{
    # Algemene tenant- of domeininstellingen voor logging/verbindingen.
    Organization = @{
        Domain          = '<contoso.com>'
        TenantId        = '<00000000-0000-0000-0000-000000000000>' # optioneel voor EXO app-only
    }

    # Serviceaccount gebruikt voor EWS/Exchange-acties (geen wachtwoord opslaan in dit bestand!).
    ServiceAccount = @{
        UserPrincipalName = '<service-account@contoso.com>'
        # Naam van de SecretManagement-entry die de PSCredential bevat.
        CredentialSecretName = '<ExoServiceCredential>'
        # Optioneel: SecretManagement-entry met een certificaat/thumbprint voor app-only.
        CertificateSecretName = '<ExoAppCert>'
    }

    # Exchange- of EWS-verbindingen; combineer met JSON-config voor mailbox-specifieke opties.
    Exchange = @{
        ConnectionType = '<EXO|OnPrem|Auto>'
        EwsUrl         = 'https://mail.contoso.com/EWS/Exchange.asmx' # optioneel bij autodiscover
        Autodiscover   = $true
        ImpersonationSmtp = '<service-account@contoso.com>'
    }

    # SecretManagement-vault die de daadwerkelijke credentials/certificaten bevat.
    SecretManagement = @{
        VaultName = '<CompanyVault>'
        # Alternatieve secretnamen indien je meerdere omgevingen of certificaten gebruikt.
        SecretNames = @{
            ExchangeCredential = '<ExoServiceCredential>'
            AppOnlyCertificate = '<ExoAppCert>'
        }
    }
}
