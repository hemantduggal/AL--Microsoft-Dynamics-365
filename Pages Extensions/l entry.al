/// <summary>
/// PageExtension G?l entry (ID 50021) extends Record General Ledger Entries;.
/// </summary>
pageextension 50024 "G?l entry" extends "General Ledger Entries"
{
    layout
    {
        addafter("External Document No.")
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
            field("EFT/Document No"; Rec."EFT/Document No")
            {
                ApplicationArea = all;
            }
        }
    }

}