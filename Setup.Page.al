page 50100 "PTE Sharepoint Setup"
{

    ApplicationArea = All;
    PageType = Card;
    SourceTable = "PTE Sharepoint Setup";
    Caption = 'PTE Sharepoint Setup';
    InsertAllowed = false;
    DeleteAllowed = false;
    UsageCategory = Administration;


    layout
    {
        area(content)
        {
            group(General)
            {
                field(TenantId; Rec.TenantId)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the TenantId.';
                }
                field(ClientId; Rec.ClientId)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the ClientId.';
                }
                field("Base Url"; Rec."Base Url")
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the Base Url.';
                }
                field(Password; Rec.Password)
                {
                    ApplicationArea = All;
                    ToolTip = 'Specifies the Password.';
                }
            }
            group("ClientId and ClientSecret")
            {

                field("Client Secret"; Rec."Client Secret")
                {
                    ToolTip = 'Specifies the Client Secret.';
                }
            }
        }
    }
    actions
    {
        area(Processing)
        {
            action(UploadCertificate)
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    FileManagement: Codeunit "File Management";
                    TempBlob: Codeunit "Temp Blob";
                    FileName: Text;
                begin
                    FileName := FileManagement.BLOBImport(TempBlob, 'Upload');
                    if FileName = '' then
                        exit;
                    Rec.SetCertificateFromBlob(TempBlob);
                    CurrPage.Update(true);
                end;
            }
            action(VerifyCert)
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    PTESharepointTest: Codeunit "PTE Sharepoint Test";
                begin
                    Rec.TestField(Password);
                    PTESharepointTest.VerifyCertificate(Rec);
                end;
            }
            action(TestSharepoint)
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    PTESharepointTest: Codeunit "PTE Sharepoint Test";
                begin
                    PTESharepointTest.TestSharepoint();
                end;
            }
            action(TestSharepointWithoutPassword)
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    PTESharepointTest: Codeunit "PTE Sharepoint Test";
                begin
                    PTESharepointTest.TestSharepointWithoutPassword();
                end;
            }
            action(TestClientSecret)
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    PTESharepointTest: Codeunit "PTE Sharepoint Test";
                begin
                    PTESharepointTest.TestClientIdWithClientSecretOAuth();
                end;
            }
        }
        area(Promoted)
        {
            actionref(UploadCertificate_Promoted; UploadCertificate)
            { }
            actionref(VerifyCert_Promoted; VerifyCert)
            { }
            actionref(TestSharepoint_Promoted; TestSharepoint)
            { }
            actionref(TestSharepointWithoutPassword_Promoted; TestSharepointWithoutPassword)
            { }
            actionref(TestClientSecret_Promoted; TestClientSecret)
            { }
        }
    }

    trigger OnOpenPage()
    begin
        Rec.InsertIfNotExists();
    end;

}
