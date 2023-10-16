/// <summary>
/// Table  Posted RFQ Line (ID 50140).
/// </summary>
table 50014 "Saless Posted RFQ Line"
{
    DataClassification = ToBeClassified;
    Caption = 'Sale Posted RFQ Line';
    Description = 'This is A Sale Posted RFQ Line Table';
    TableType = Normal;

    fields
    {
        field(1; Customer_No; Code[20])
        {
            DataClassification = ToBeClassified;

        }
        field(2; "No."; Code[20])
        {
            DataClassification = ToBeClassified;


        }
        field(3; RFQ_Type; Enum "G/L Account Type")
        {
            DataClassification = ToBeClassified;

        }
        field(4; item_no; Code[20])
        {
            DataClassification = ToBeClassified;

        }
        field(5; Decription; Text[100])
        {
            DataClassification = ToBeClassified;

        }
        field(6; Quantity; Decimal)
        {
            DataClassification = ToBeClassified;

        }
        field(7; UOM; Text[50])
        {
            DataClassification = ToBeClassified;

        }
        field(8; Direct_unit_cost; Decimal)
        {
            DataClassification = ToBeClassified;

        }

        field(9; rate; Decimal)
        {
            DataClassification = ToBeClassified;

        }

        field(10; brand; Text[20])
        {
            DataClassification = ToBeClassified;

        }

        field(11; salesperson; text[50])
        {
            DataClassification = ToBeClassified;
        }
        field(12; Application_master; text[100])
        {
            DataClassification = ToBeClassified;
        }
        field(13; Project_master; text[100])
        {
            DataClassification = ToBeClassified;
        }





    }

    keys
    {
        key(PK; "No.")
        {
            Clustered = true;
        }
    }

    var
        myInt: Integer;

    trigger OnInsert()
    begin
        Message('Data inserted');
    end;

    trigger OnModify()
    begin
        Message('Data Modified');
    end;

    trigger OnDelete()
    begin
        Message('Data Deleted');
    end;

    trigger OnRename()
    begin
        Message('Data Renamed');
    end;

}