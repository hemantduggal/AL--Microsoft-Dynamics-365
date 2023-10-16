/// <summary>
/// Table RFQ Line (ID 50138).
/// </summary>
table 50002 " Sale RFQ Line"
{
    DataClassification = ToBeClassified;
    Caption = ' Sale RFQ Line';
    Description = 'This is A  Sale RFQ Line Table';
    TableType = Normal;

    fields
    {
        field(1; Customer_No; Code[20])
        {
            DataClassification = ToBeClassified;
            TableRelation = Customer."No.";

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
            TableRelation = Item."No.";
            trigger OnValidate()
            var
                recitem: record Item;
            begin
                if item_no = '' then
                    Error('Enter item No');
                if recitem.get(Rec.item_no) then
                    rec.Decription := recitem.Description;
                rec.Direct_unit_cost := recitem."Unit Price";
                rec.UOM := recitem."Base Unit of Measure";
                rec.rate := recitem."Unit Cost";
                rec.brand := recitem."Item Category Code";
                // rec.Customer_No :=recitem.



                Rec.Modify()



            end;


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