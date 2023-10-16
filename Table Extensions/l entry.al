/// <summary>
/// TableExtension G/L entry  (ID 50021) extends Record G/L Entry.
/// </summary>
tableextension 50021 "G/L entry " extends "G/L Entry"
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
        field(1123; "EFT/Document No"; Code[20])
        {
            DataClassification = ToBeClassified;
        }
    }
}