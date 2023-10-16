/// <summary>
/// Page RFQ card page (ID 50142).
/// </summary>
page 50002 "Sale RFQ card page"
{
    PageType = Card;
    ApplicationArea = All;
    UsageCategory = Documents;
    SourceTable = 50001;
    Caption = ' Sale RFQ header card page ';
    AboutText = 'Card PAge ';
    AboutTitle = 'This Is a card Page ';


    layout
    {
        area(Content)
        {
            group(GeneralDetail)
            {
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
                    ApplicationArea = All;

                }
                field(original_customer_name; Rec.original_customer_name)
                {
                    ApplicationArea = All;
                }
                field(Document_Date; Rec.Document_Date)
                {
                    ApplicationArea = All;

                }
                field(Posting_date; Rec.Posting_date)
                {
                    ApplicationArea = All;

                }


                field(Location_Code; Rec.Location_Code)
                {
                    ApplicationArea = All;

                }
                field(salesperson; Rec.salesperson)
                {
                    ApplicationArea = All;

                }
                field(remarks; Rec.remarks)
                {
                    ApplicationArea = All;

                }
                field(rf; Rec.rf)
                {
                    ApplicationArea = all;
                }


            }
            group(otherdetails)
            {
                part(general; "Sale RFQ subform page")
                {
                    SubPageLink = "No." = field("No.");
                    ApplicationArea = All;

                }

            }

        }
    }
    actions
    {
        area(Navigation)
        {
            group("POSTING")
            {
                Caption = 'POSTING';
                Image = PostingEntries;
                action(Post)
                {
                    ApplicationArea = All;
                    //RunObject = page "Saless Posted RFQ Card page";
                    // RunPageLink = "No." = field("No.");
                    Image = PostOrder;
                    Promoted = true;
                    Ellipsis = true;
                    PromotedCategory = Process;
                    PromotedIsBig = true;


                    trigger OnAction()
                    var
                        rfq: Record "Salesss Posted RFQ Header";
                    begin

                        rfq.Reset();

                        rfq.SetRange("No.", rec."No.");
                        if rfq.FindSet then
                            rfq."No." := rec."No.";
                        rfq.Customer_No := rec.Customer_No;
                        rfq.Customer_name := rec.Customer_name;
                        rfq.Document_Date := rec.Document_Date;
                        rfq.Location_Code := rec.Location_Code;
                        rfq.original_customer_name := rec.original_customer_name;
                        rfq.Posting_date := rec.Posting_date;
                        rfq.remarks := rec.remarks;
                        rfq.salesperson := rec.salesperson;


                        xRec.rf := true;
                        //rec.Modify();


                        //rfq.Modify();
                        // rec.Delete();
                        Message('Data inserted');


                    end;
                }






            }
        }
    }





    var
        myInt: Integer;


}