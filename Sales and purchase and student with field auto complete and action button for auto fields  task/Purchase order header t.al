/// <summary>
/// TableExtension items (ID 50104) extends Record Item.
/// </summary>
tableextension 50109 "Purchase Order extension" extends "Purchase Header"
{
    fields
    {
        field(1111; StudentNO; code[20])
        {
            Caption = 'StudentNO';
            DataClassification = ToBeClassified;
            TableRelation = "Student Header Table";

        }
        field(1112; StudentName; Text[50])
        {
            Caption = 'StudentName';
            DataClassification = ToBeClassified;
            //TableRelation = "Student Header Table";

        }
        field(1113; StudentAddress; Text[50])
        {
            Caption = 'StudentName';
            DataClassification = ToBeClassified;
            //TableRelation = "Student Header Table";

        }
        field(11111; StudentPhone; Integer)
        {
            Caption = 'StudentPhone';
            DataClassification = ToBeClassified;
            //TableRelation = "Student Header Table";
        }
        field(11117; StudentState; Text[20])
        {
            Caption = 'StudentState';
            //OptionMembers = Delhi,Haryana,Up,Gujrat,"Andhra-Pradesh","Arunachal Pradesh",Assam,Bihar,Chhattisgarh,Goa;
            DataClassification = ToBeClassified;
            //TableRelation = "Student Header Table";
        }
        field(11118; StudentCity; text[20])
        {
            Caption = 'StudentState';
            // OptionMembers = Mumbai,chennai,banglore,kolkata,pune,jaipur,ahmedabad,faridabad,okhla,raipur;

            DataClassification = ToBeClassified;
            //TableRelation = "Student Header Table";
        }

        field(1119; "Contact person"; Text[100])
        {
            DataClassification = ToBeClassified;
            TableRelation = Customer."No.";
            trigger OnValidate()
            var
                sh: record Customer;
            begin
                if sh.get(Rec."Sell-to Customer No.") then
                    Rec."Contact person" := sh.Contact;




                Rec.Modify()
            end;
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


        field(1123; "Index no"; code[20])
        {
            DataClassification = ToBeClassified;
        }
        field(1124; Plank_no; Integer)
        {
            DataClassification = ToBeClassified;
        }
        field(1125; "Currency Capacity"; Decimal)
        {
            DataClassification = ToBeClassified;
        }
        field(1126; "unit od specification"; option)
        {
            DataClassification = ToBeClassified;
            OptionMembers = basic,medium,high;
        }




    }
}
