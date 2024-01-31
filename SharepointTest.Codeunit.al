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
        TempBlob: Codeunit "Temp Blob";
    begin
        TempBlob.FromRecord(PTESharepointSetup, PTESharepointSetup.FieldNo(Certificate));
        TempBlob.CreateInStream(CertInStream);
        CertBase64 := Base64Convert.ToBase64(CertInStream);
        X509Certificate2.VerifyCertificate(CertBase64, PTESharepointSetup.GetPassword(), Enum::"X509 Content Type"::Pkcs12)
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
        TempBlob.CreateInStream(CertInStream);
        CertBase64 := Base64Convert.ToBase64(CertInStream);
        SharePointAuthorization := SharePointAuth.CreateClientCredentials(FormatGuid(SharepointSetup.TenantId), FormatGuid(SharepointSetup.ClientId), CertBase64, SharepointSetup.GetPassword(), 'https://graph.microsoft.com/.default');
        SharePointClient.Initialize(SharepointSetup."Base Url", SharePointAuthorization);
        if not SharePointClient.GetLists(SharePointList) then begin
            HTTPDiagnostics := SharePointClient.GetDiagnostics();
            Error(HTTPDiagnostics.GetResponseReasonPhrase());
        end;
        SharepointLists.SetLists(SharePointList);
        SharepointLists.Run();
    end;

    local procedure FormatGuid(GuidToFormat: Guid): Text
    begin
        exit(Format(GuidToFormat, 0, 4));
    end;

}