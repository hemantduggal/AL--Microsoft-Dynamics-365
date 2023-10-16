/// <summary>
/// PageExtension inventorySetupPageExt (ID 50002) extends Record inventory setup.
/// </summary>
pageextension 50002 inventorySetupPageExt extends "inventory setup"
{
    layout
    {
        addafter("Item Nos.")
        {
            field("Book Nos."; Rec."Book Nos.")
            {
                Caption = 'Book Nos.';
                ApplicationArea = All;
            }
        }
    }
}