
/// <summary>
/// PageExtension Vat Entry (ID 50025) extends Record VAT Entries.
/// </summary>
pageextension 50025 "Vat Entry" extends "VAT Entries"
{


    layout
    {
        addafter("Internal Ref. No.")
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