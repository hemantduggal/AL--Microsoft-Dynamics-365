/// <summary>
/// PageExtension demo3 (ID 50001) extends Record Customer Ledger Entries.
/// </summary>
pageextension 50009 dem extends "Customer Card"
{
    layout
    {
        addafter("State Code")
        {
            field("Contact person"; Rec."Contact person")
            {
                ApplicationArea = all;


            }
            field(Remarks; Rec.Remarks)
            {
                ApplicationArea = all;

            }
            field(Deal_date; Rec.Deal_date)
            {
                ApplicationArea = all;
            }
            field("order category"; Rec."order category")
            {
                ApplicationArea = all;

            }
        }
    }

}