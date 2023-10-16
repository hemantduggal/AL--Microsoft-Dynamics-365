/// <summary>
/// PageExtension item ledger entry pageext (ID 50129) extends Record Item Ledger Entries.
/// </summary>
pageextension 50129 "item ledger entry pageext" extends "Item Ledger Entries"
{
    layout
    {
        addafter("Entry No.")
        {
            field(itemcategorycode; Rec.itemcategorycode)
            {
                ApplicationArea = all;
            }
            field(heatno; Rec.heatno)
            {
                ApplicationArea = all;
            }
            field(Billet_no; Rec.Billet_no)
            {
                ApplicationArea = all;
            }
            field(Size; Rec.Size)
            {
                ApplicationArea = all;

            }
            field(Dia; Rec.Dia)
            {
                ApplicationArea = all;
            }// Add changes to page layout here
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
            field("Adv.Prepayment Amount"; Rec."Adv.Prepayment Amount")
            {
                ApplicationArea = all;
            }
        }
    }

    actions
    {
        // Adds the action called "My Actions" to the Action menu 
        addfirst(Processing)
        {
            action("My Actions")
            {
                Promoted = true;
                PromotedCategory = Process;
                ApplicationArea = All;
                trigger OnAction()
                begin

                end;
            }
        }
        addfirst(Creation)
        {
            // Adds the action "My New document" to the New Document group in the Actions menu. 
            action("My New document")
            // {
            //     ApplicationArea = All;
            //     RunObject = page "Item Ledger Entries";
            //     Image = "1099Form";
            // }
            {
                ApplicationArea = All;
                RunObject = report "Item ledger entry report";
            }
        }
        addfirst(Navigation)
        {
            // Adds the action called "My Navigate" to the Navigate menu. 
            action("My Navigate")
            // {
            //     ApplicationArea = All;
            //     RunObject = page "Item Ledger Entries";
            // }
            {
                ApplicationArea = All;
                RunObject = report "Item ledger entry report";
            }
        }
        addfirst(Reporting)
        {

            action("My Report")
            {
                ApplicationArea = All;
                RunObject = report "Item ledger entry report";

            }

        }

    }

    var
        myInt: Integer;

}
