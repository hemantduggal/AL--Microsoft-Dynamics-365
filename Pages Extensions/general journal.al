/// <summary>
/// PageExtension general journal (ID 50025) extends Record General Journal.
/// </summary>
pageextension 50027 "general journals" extends "General Journal"
{
    layout
    {
        addafter("External Document No.")
        {

            field("EFT/Document No"; Rec."EFT/Document No")
            {
                ApplicationArea = all;
            }
        }
    }

}