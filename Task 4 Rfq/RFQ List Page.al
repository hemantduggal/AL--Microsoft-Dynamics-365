/// <summary>
/// Page RFQ list Page (ID 50141).
/// </summary>
page 50001 " Sale RFQ list Page"
{
    PageType = List;
    ApplicationArea = All;
    UsageCategory = Documents;
    SourceTable = 50001;
    CardPageId = "Sale RFQ Card page";




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
                }
                field(Customer_No; Rec.Customer_No)
                {
                    ApplicationArea = All;
                    NotBlank = true;

                }
                field(Customer_name; Rec.Customer_name)
                {
                    ApplicationArea = all;
                    NotBlank = true;
                }
                field(rf; Rec.rf)
                {
                    ApplicationArea = all;
                }




            }
        }
    }
    trigger OnOpenPage()
    var
        r: Record "Sale RFQ Header";
    begin
        if r.Get(false) then
            if r.rf then
                exit;
        rec.FilterGroup(100);
        rec.SetRange(rf, true);
        rec.FilterGroup(0);
    end;

    // actions
    // {
    //     area(Processing)
    //     {
    //         action(Post)
    //         {
    //             ApplicationArea = All;
    //             //RunObject = page "Saless Posted RFQ Card page";
    //             // RunPageLink = "No." = field("No.");
    //             Image = PostOrder;
    //             Promoted = true;
    //             Ellipsis = true;
    //             PromotedCategory = Process;
    //             PromotedIsBig = true;


    //             trigger OnAction()
    //             var
    //                 rfq: Record 50013;
    //             begin

    //                 rfq.Reset();

    //                 rfq.SetRange(Customer_No, Rec.Customer_No);
    //                 if rfq.FindSet then
    //                     rfq.Customer_No := rec.Customer_No;
    //                 rfq.Customer_name := rec.Customer_name;
    //                 rfq.Document_Date := rec.Document_Date;
    //                 rfq.Location_Code := rec.Location_Code;
    //                 rfq.original_customer_name := rec.original_customer_name;
    //                 rfq.Posting_date := rec.Posting_date;
    //                 rfq.remarks := rec.remarks;
    //                 rfq.salesperson := rec.salesperson;


    //                 rfq.Modify();
    //                 Message('Data inserted');


    //             end;
    //         }



    //     }
    // }


}