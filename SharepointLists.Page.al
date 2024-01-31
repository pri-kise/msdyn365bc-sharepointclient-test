page 50101 "PTE Sharepoint Lists"
{
    Caption = 'Sharepoint Lists';
    PageType = List;
    UsageCategory = None;
    ApplicationArea = All;
    SourceTable = "SharePoint List";
    SourceTableTemporary = true;
    Editable = false;

    layout
    {
        area(Content)
        {
            repeater(Control1)
            {
                field(Id; Rec.Id)
                {
                    ToolTip = 'Specifies the Id.';
                }
                field("List Item Entity Type"; Rec."List Item Entity Type")
                {
                    ToolTip = 'Specifies the List Item Entity Type Full Name.';
                }
                field(Description; Rec.Description)
                {
                    ToolTip = 'Specifies the Description.';
                }
                field("Base Template"; Rec."Base Template")
                {
                    ToolTip = 'Specifies the Base Template.';
                }
                field("Base Type"; Rec."Base Type")
                {
                    ToolTip = 'Specifies the Base Type.';
                }
                field("Is Catalog"; Rec."Is Catalog")
                {
                    ToolTip = 'Specifies the Is Catalog.';
                }
                field(Created; Rec.Created)
                {
                    ToolTip = 'Specifies the Created.';
                }
                field(OdataId; Rec.OdataId)
                {
                    ToolTip = 'Specifies the Odata.Id.';
                }
                field(OdataEditLink; Rec.OdataEditLink)
                {
                    ToolTip = 'Specifies the Odata.EditLink.';
                }
                field(OdataType; Rec.OdataType)
                {
                    ToolTip = 'Specifies the Odata.Type.';
                }

            }
        }
        area(Factboxes)
        {

        }
    }

    internal procedure SetLists(var SharePointList: Record "SharePoint List" temporary)
    begin
        Rec.Copy(SharePointList, true);
    end;
}