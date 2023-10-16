/// <summary>
/// TableExtension inventorySetupExt (ID 500012) extends Record Inventory Setup.
/// </summary>
tableextension 50002 inventorySetupExt extends "Inventory Setup"
{
    fields
    {
        field(50200; "Book Nos."; Code[20])
        {
            Caption = 'Book Nos.';
            TableRelation = "No. Series";
        }
    }
}