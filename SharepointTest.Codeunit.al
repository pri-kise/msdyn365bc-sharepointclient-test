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

    local procedure GetScopes() Scopes: List of [Text]
    begin
        Scopes.Add('https://microsoft.sharepoint.com/.default');
        Scopes.Add('https://graph.microsoft.com/.default');
    end;

}