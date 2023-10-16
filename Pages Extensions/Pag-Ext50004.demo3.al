/// <summary>
/// PageExtension demo3 (ID 50004) extends Record Customer Ledger Entries.
/// </summary>
pageextension 50004 demo3 extends "Customer Ledger Entries"
{
    layout
    {
        addafter("Location GST Reg. No.")
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
            field("Adv.Prepayment Amount"; Rec."Adv.Prepayment Amount")
            {
                ApplicationArea = all;
            }
        }
    }

}
