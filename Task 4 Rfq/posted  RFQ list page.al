/// <summary>
/// Page RFQ list Page (ID 50144).
/// </summary>
page 50010 "Saless Posted RFQ list Page"
{
    PageType = List;
    ApplicationArea = All;
    UsageCategory = Documents;
    SourceTable = 50013;
    CardPageId = "Saless Posted RFQ Card page";





    layout
    {
        area(Content)
        {
            repeater(General)
            {
                Caption = 'general';

                field("No."; rec."No.")
                {
                    ApplicationArea = All;

                    Importance = Additional;
                }
                field(Customer_No; Rec.Customer_No)
                {
                    ApplicationArea = All;

                    Importance = Additional;

                }
                field(Customer_name; Rec.Customer_name)
                {
                    ApplicationArea = all;

                    Importance = Additional;
                }




            }
        }
    }



}