table 50100 "PTE Sharepoint Setup"
{

    DataClassification = CustomerContent;
    fields
    {
        field(1; "Primary Key"; Code[10])
        {
            AllowInCustomizations = Never;
        }
        field(2; "Base Url"; Text[255])
        {
            Caption = 'Base Url';
        }
        field(3; "TenantId"; Guid)
        {
            Caption = 'TenantId';
        }

        field(10; "ClientId"; Guid)
        {
            Caption = 'ClientId';
        }

        field(20; "Certificate"; Blob)
        {
            Caption = 'Certificate';
        }

        field(21; "Password"; Text[50])
        {
            Caption = 'Password';
            ExtendedDatatype = Masked;
        }



        //You might want to add fields here

    }

    keys
    {
        key(PK; "Primary Key")
        {
            Clustered = true;
        }
    }

    var
        RecordHasBeenRead: Boolean;

    procedure GetRecordOnce()
    begin
        if RecordHasBeenRead then
            exit;
        Get();
        RecordHasBeenRead := true;
    end;

    procedure InsertIfNotExists()
    begin
        Reset();
        if not Get() then begin
            Init();
            Insert(true);
        end;
    end;

    procedure SetCertificateFromBlob(TempBlob: Codeunit "Temp Blob")
    var
        RecordRef: RecordRef;
    begin
        RecordRef.GetTable(Rec);
        TempBlob.ToRecordRef(RecordRef, Rec.FieldNo(Certificate));
        RecordRef.SetTable(Rec);
    end;

    internal procedure GetPassword(): SecretText
    begin
        exit(Rec.Password);
    end;

}