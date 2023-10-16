/// <summary>
/// Table RFQ Header (ID 50137).
/// </summary>
table 50001 "Sale RFQ Header"
{
    DataClassification = ToBeClassified;
    Caption = 'Sale RFQ Header';
    Description = 'This is A RFQ Header Table';
    TableType = Normal;

    fields
    {
        field(1; "No."; Code[10])
        {
            Caption = 'No.';
            DataClassification = CustomerContent;

            trigger OnValidate()
            begin
                if "No." <> xRec."No." then begin
                    inventorySetup.Get();
                    NoSeriesMgt.TestManual(inventorySetup."Book Nos.");
                    "No. Series" := '';
                end;
            end;
        }
        field(2; "No. Series"; Code[20])
        {
            Caption = 'No. Series';
            Editable = false;
            TableRelation = "No. Series";
        }
        field(3; Customer_No; Code[20])
        {
            DataClassification = ToBeClassified;
            TableRelation = Customer."No.";


            Caption = 'Customer no';
            ValidateTableRelation = false;


            trigger OnValidate()
            var
                reccust: record Customer;
            begin
                if Customer_No = '' then
                    Error('Enter Customer No');

                if reccust.get(Rec.Customer_No) then
                    rec.Customer_No := rec.Customer_No;
                rec.Customer_name := reccust.Name;
                rec.Location_Code := reccust."Location Code";
                rec.salesperson := reccust."Salesperson Code";

                Rec.Modify();

                // if xRec.Customer_No <> Rec.Customer_No then begin
                //     reccust.Reset();
                //     if reccust.get(Customer_No) then begin

                //     end


            end;

            //end;








        }
        field(4; Customer_name; Text[100])
        {
            DataClassification = ToBeClassified;


        }
        field(10; original_customer_name; text[100])
        {
            DataClassification = ToBeClassified;
        }
        field(5; Document_Date; Date)
        {
            DataClassification = ToBeClassified;

        }
        field(6; Posting_date; Date)
        {
            DataClassification = ToBeClassified;

        }

        field(7; Location_Code; Code[20])
        {
            DataClassification = ToBeClassified;

        }
        field(8; salesperson; Text[50])
        {
            DataClassification = ToBeClassified;
        }
        field(9; remarks; Text[100])
        {
            DataClassification = ToBeClassified;
        }
        field(11; rf; Boolean)
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
    trigger OnInsert()
    var
    begin
        if "No." = '' then begin
            inventorySetup.Get();
            inventorySetup.TestField("Book Nos.");
            NoSeriesMgt.InitSeries(inventorySetup."Book Nos.", xRec."No. Series", 0D, "No.", "No. Series");
        end;
    end;

    var
        myInt: Integer;
        inventorySetup: Record "Inventory Setup";
        NoSeriesMgt: Codeunit NoSeriesManagement;
        rf: boolean;


}