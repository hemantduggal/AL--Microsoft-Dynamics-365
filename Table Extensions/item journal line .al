/// <summary>
/// TableExtension  Sales head Gen. Journal Line (ID 50009) extends Record Gen. Journal Line.
/// </summary>
tableextension 50010 " item Journal Line" extends "Item Journal Line"
{

    fields
    {
        field(1123; item_category_code; Integer)
        {
            DataClassification = ToBeClassified;
            TableRelation = Item;

            trigger OnValidate()
            var
                itemR: record item;
            begin
                if itemR.get(Rec."No.") then
                    Rec.item_category_code := itemR.itemcategorycode;
                Rec.HEat_no := itemR.heatno;
                rec.Billet_no := itemR.Billet_no;
                rec.Size := itemR.Size;
                rec.Dia := itemR.Dia;



                Rec.Modify()



            end;







        }
        field(1127; HEat_no; Integer)
        {
            DataClassification = ToBeClassified;
        }
        field(1124; Billet_no; Integer)
        {
            DataClassification = ToBeClassified;
        }
        field(1125; Size; Decimal)
        {
            DataClassification = ToBeClassified;
        }
        field(1126; Dia; Decimal)
        { DataClassification = ToBeClassified; }
        field(120; "Index no"; code[20])
        {
            DataClassification = ToBeClassified;
        }
        field(117; Plank_no; Integer)
        {
            DataClassification = ToBeClassified;
        }
        field(118; "Currency Capacity"; Decimal)
        {
            DataClassification = ToBeClassified;
        }
        field(119; "unit od specification"; option)
        {
            DataClassification = ToBeClassified;
            OptionMembers = basic,medium,high;
        }


    }
}