codeunit 50100 "PTE Sharepoint Test"
{
    Access = Internal;
    InherentPermissions = X;

    procedure VerifyCertificate(var PTESharepointSetup: Record "PTE Sharepoint Setup")
    var

        Base64Convert: Codeunit "Base64 Convert";
        CertInStream: InStream;
        CertBase64: Text;
        X509Certificate2: Codeunit X509Certificate2;
        OAuth2: Codeunit OAuth2;
        // CertPropertyJson: Text;
        TempBlob: Codeunit "Temp Blob";
    begin
        TempBlob.FromRecord(PTESharepointSetup, PTESharepointSetup.FieldNo(Certificate));
        TempBlob.CreateInStream(CertInStream);
        CertBase64 := Base64Convert.ToBase64(CertInStream);
        X509Certificate2.VerifyCertificate(CertBase64, PTESharepointSetup.GetPassword(), Enum::"X509 Content Type"::Pkcs12);
        CertInStream.ResetPosition();
        // CertBase64 := Base64Convert.ToBase64(CertInStream);
        // X509Certificate2.GetCertificatePropertiesAsJson(CertBase64, PTESharepointSetup.GetPassword(), CertPropertyJson);
        // Message(CertPropertyJson);
    end;

    procedure TestSharepoint()
    var
        SharepointSetup: Record "PTE Sharepoint Setup";
        Base64Convert: Codeunit "Base64 Convert";
        CertInStream: InStream;
        CertBase64: Text;
        X509Certificate2: Codeunit X509Certificate2;
        TempBlob: Codeunit "Temp Blob";
        SharePointClient: Codeunit "SharePoint Client";
        SharePointAuthorization: Interface "SharePoint Authorization";
        SharePointAuth: Codeunit "SharePoint Auth.";
        SharePointList: Record "SharePoint List" temporary;
        SharepointLists: Page "PTE Sharepoint Lists";
        HTTPDiagnostics: Interface "HTTP Diagnostics";
    begin
        SharepointSetup.Get();
        VerifyCertificate(SharepointSetup);

        TempBlob.FromRecord(SharepointSetup, SharepointSetup.FieldNo(Certificate));
        TempBlob.CreateInStream(CertInStream, TextEncoding::Windows);
        CertBase64 := Base64Convert.ToBase64(CertInStream);
        SharePointAuthorization := SharePointAuth.CreateClientCredentials(FormatGuid(SharepointSetup.TenantId), FormatGuid(SharepointSetup.ClientId), CertBase64, SharepointSetup.GetPassword(), GetScopes());
        SharePointClient.Initialize(SharepointSetup."Base Url", SharePointAuthorization);
        if not SharePointClient.GetLists(SharePointList) then begin
            HTTPDiagnostics := SharePointClient.GetDiagnostics();
            Error(HTTPDiagnostics.GetResponseReasonPhrase());
        end;
        SharepointLists.SetLists(SharePointList);
        SharepointLists.Run();
    end;

    procedure TestSharepointWithoutPassword()
    var
        SharepointSetup: Record "PTE Sharepoint Setup";
        Base64Convert: Codeunit "Base64 Convert";
        CertInStream: InStream;
        CertBase64: Text;
        TempBlob: Codeunit "Temp Blob";
        SharePointClient: Codeunit "SharePoint Client";
        SharePointAuthorization: Interface "SharePoint Authorization";
        SharePointAuth: Codeunit "SharePoint Auth.";
        SharePointList: Record "SharePoint List" temporary;
        SharepointLists: Page "PTE Sharepoint Lists";
        HTTPDiagnostics: Interface "HTTP Diagnostics";
    begin
        SharepointSetup.Get();

        TempBlob.FromRecord(SharepointSetup, SharepointSetup.FieldNo(Certificate));
        TempBlob.CreateInStream(CertInStream, TextEncoding::Windows);
        CertBase64 := Base64Convert.ToBase64(CertInStream);
        SharePointAuthorization := CreateClientCredentials(FormatGuid(SharepointSetup.TenantId), FormatGuid(SharepointSetup.ClientId), CertBase64, GetScopes());
        SharePointClient.Initialize(SharepointSetup."Base Url", SharePointAuthorization);
        if not SharePointClient.GetLists(SharePointList) then begin
            HTTPDiagnostics := SharePointClient.GetDiagnostics();
            Error(HTTPDiagnostics.GetResponseReasonPhrase());
        end;
        SharepointLists.SetLists(SharePointList);
        SharepointLists.Run();
    end;

    procedure CreateClientCredentials(AadTenantId: Text; ClientId: Text; Certificate: Text; Scope: Text): Interface "SharePoint Authorization"
    var
        Scopes: List of [Text];
    begin
        Scopes.Add(Scope);
        exit(CreateClientCredentials(AadTenantId, ClientId, Certificate, Scopes));
    end;

    procedure CreateClientCredentials(AadTenantId: Text; ClientId: Text; Certificate: Text; Scopes: List of [Text]): Interface "SharePoint Authorization"
    var
        PTESharepointClientCred: Codeunit "PTE SharepointClientCred.";
    begin
        PTESharepointClientCred.SetParameters(AadTenantId, ClientId, Certificate, Scopes);
        exit(PTESharepointClientCred);
    end;

    local procedure FormatGuid(GuidToFormat: Guid): Text
    begin
        exit(Format(GuidToFormat, 0, 4));
    end;

    procedure TestClientIdWithClientSecretOAuth()
    var
        SharepointSetup: Record "PTE Sharepoint Setup";
        OAuth2: Codeunit OAuth2;
    begin
        SharepointSetup.Get();
        GetToken(FormatGuid(SharepointSetup.TenantId), FormatGuid(SharepointSetup.ClientId), SharepointSetup.GetClientSecret(), GetScopes());
    end;

    local procedure GetToken(AadTenantId: Text; ClientId: Text; ClientSecret: SecretText; Scopes: List of [Text]): SecretText
    var
        ErrorText: Text;
        AccessToken: SecretText;
    begin
        if not AcquireToken(AadTenantId, ClientId, ClientSecret, Scopes, AccessToken, ErrorText) then
            Error(ErrorText);
        exit(AccessToken);
    end;

    local procedure AcquireToken(AadTenantId: Text; ClientId: Text; ClientSecret: SecretText; Scopes: List of [Text]; var AccessToken: SecretText; var ErrorText: Text): Boolean
    var
        OAuth2: Codeunit System.Security.Authentication.OAuth2;
        FailedErr: Label 'Failed to retrieve an access token.';
        //TODO: Check Authority Url
        ClientCredentialsTokenAuthorityUrlTxt: Label 'https://login.microsoftonline.com/%1/oauth2/v2.0/token', Comment = '%1 = AAD tenant ID', Locked = true;
        IsSuccess: Boolean;
        AuthorityUrl: Text;
    begin
        AuthorityUrl := StrSubstNo(ClientCredentialsTokenAuthorityUrlTxt, AadTenantId);
        ClearLastError();
        if (not OAuth2.AcquireAuthorizationCodeTokenFromCache(ClientId, ClientSecret, AuthorityUrl, '', Scopes, AccessToken)) or (AccessToken.IsEmpty()) then
            OAuth2.AcquireTokenWithClientCredentials(ClientId, ClientSecret, AuthorityUrl, '', Scopes, AccessToken);

        IsSuccess := not AccessToken.IsEmpty();

        if not IsSuccess then begin
            ErrorText := GetLastErrorText();
            if ErrorText = '' then
                ErrorText := FailedErr;
        end;

        exit(IsSuccess);
    end;

    local procedure GetScopes() Scopes: List of [Text]
    begin
        Scopes.Add('00000003-0000-0ff1-ce00-000000000000/.default'); //guid is the Application Id for Office 365 SharePoint Online
        // Scopes.Add('https://microsoft.sharepoint.com/.default');
        // Scopes.Add('https://graph.microsoft.com/.default');
    end;

}