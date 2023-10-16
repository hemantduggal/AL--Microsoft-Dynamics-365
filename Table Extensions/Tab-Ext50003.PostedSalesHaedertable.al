/// <summary>
/// TableExtension Posted Sales Haeder table (ID 50003) extends Record Sales Invoice Header.
/// </summary>
tableextension 50003 "Posted Sales Haeder table" extends "Sales Invoice Header"
{
    fields
    {
        field(1119; "Contact person"; Text[100])
        {
            DataClassification = ToBeClassified;
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
