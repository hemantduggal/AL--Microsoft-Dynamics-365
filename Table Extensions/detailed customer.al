/// <summary>
/// TableExtension custo (ID 50010) extends Record Detailed Cust. Ledg. Entry.
/// </summary>
tableextension 50011 "custo" extends "Detailed Cust. Ledg. Entry"
{

    fields
    {
        field(1119; "Contact person"; Text[100])
        {
            DataClassification = ToBeClassified;
            TableRelation = Customer."No.";

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
    }
}