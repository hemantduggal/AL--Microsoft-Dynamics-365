/// <summary>
/// Table  Posted RFQ Header (ID 50139).
/// </summary>
table 50013 "Salesss Posted RFQ Header"
{
    DataClassification = ToBeClassified;
    Caption = ' Saless Posted RFQ Header Table';
    Description = 'This is A Salesss Posted RFQ Header Table';
    TableType = Normal;


    fields
    {
        field(1; "No."; Code[10])
        {
            Caption = ' RFQ Doc No.';
            DataClassification = ToBeClassified;

        }


        field(2; "No. Series"; Code[20])
        {
            Caption = 'No. Series';
            DataClassification = ToBeClassified;


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

                if reccust.get(Rec.Customer_No) then
                    rec.Customer_No := rec.Customer_No;
                rec.Customer_name := reccust.Name;
                rec.Location_Code := reccust."Location Code";
                rec.salesperson := reccust."Salesperson Code";

                Rec.Modify();
            end;
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


    }

    keys
    {
        key(PK; Customer_No)
        {
            Clustered = true;
        }
    }


    var
        myInt: Integer;



}