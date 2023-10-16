/// <summary>
/// PageExtension posted purch. (ID 50014) extends Record Posted Purchase Invoice.
/// </summary>
pageextension 50014 "posted purch." extends "Posted Purchase Invoice"
{
    layout
    {
        addafter("Due Date")
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
            field("Index no"; Rec."Index no")
            {
                ApplicationArea = all;

            }
            field(Plank_no; Rec.Plank_no)
            {
                ApplicationArea = all;

            }
            field("Currency Capacity"; Rec."Currency Capacity")
            {
                ApplicationArea = all;
            }
            field("unit od specification"; Rec."unit od specification")
            {
                ApplicationArea = all;
            }
        }
    }

}