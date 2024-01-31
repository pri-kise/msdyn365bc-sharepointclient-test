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
        }
    }

    trigger OnOpenPage()
    begin
        Rec.InsertIfNotExists();
    end;

}
