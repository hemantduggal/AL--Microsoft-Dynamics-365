/// <summary>
/// TableExtension demo2 (ID 50005) extends Record Cust. Ledger Entry.
/// </summary>
tableextension 50005 demo2 extends "Cust. Ledger Entry"
{
    fields
    {
        field(1119; "Contact person"; Text[100])
        {
            DataClassification = ToBeClassified;
            TableRelation = Customer."No.";
            // trigger OnValidate()
            // var
            //     sh: record Customer;
            // begin
            //     if sh.get(Rec."Sell-to Customer No.") then
            //         Rec."Contact person" := sh.Contact;




            //     Rec.Modify()
            // end;
        }
        field(1120; "Remarks"; Text[100])
        {
            DataClassification = ToBeClassified;
        }
        field(1121; Deal_date; Date)
        {
            DataClassification = ToBeClassified;
        }
        field(1122; "order category"; Option)
        {
            DataClassification = ToBeClassified;
            OptionMembers = Process,Raw,Finished;


        }
        field(50106; "Adv.Prepayment Amount"; Decimal)
        {
            DataClassification = ToBeClassified;
        }
    }
}
