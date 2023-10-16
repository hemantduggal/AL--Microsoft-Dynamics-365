// /// <summary>
// /// Page Posted RFQ card page (ID 50145).
// /// </summary>
/// <summary>
/// Page Saless Posted RFQ Card page (ID 50011).
/// </summary>
page 50011 "Saless Posted RFQ Card page"
{
    PageType = Card;
    ApplicationArea = All;
    UsageCategory = Administration;
    SourceTable = 50013;
    Caption = 'Saless Posted RFQ header card page ';
    AboutText = ' Saless Posted RFQ Card PAge ';
    AboutTitle = 'This Is a  Saless Posted rfq card Page ';


    layout
    {
        area(Content)
        {
            group(GeneralDetail)
            {
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
                    ApplicationArea = All;

                    Importance = Additional;

                }
                field(original_customer_name; Rec.original_customer_name)
                {
                    ApplicationArea = All;

                    Importance = Additional;
                }
                field(Document_Date; Rec.Document_Date)
                {
                    ApplicationArea = All;

                    Importance = Additional;
                }
                field(Posting_date; Rec.Posting_date)
                {
                    ApplicationArea = All;

                    Importance = Additional;

                }


                field(Location_Code; Rec.Location_Code)
                {
                    ApplicationArea = All;

                    Importance = Additional;

                }
                field(salesperson; Rec.salesperson)
                {
                    ApplicationArea = All;

                    Importance = Additional;


                }
                field(remarks; Rec.remarks)
                {
                    ApplicationArea = All;

                    Importance = Additional;

                }


            }
            group(otherdetails)
            {
                part(general; "Sale Posted RFQ subform page")
                {
                    SubPageLink = "No." = field("No.");
                    ApplicationArea = All;



                }

            }

        }
    }





    // actions
    // {

    //     area(Processing)
    //     {
    //         action(POST)
    //         {
    //             ApplicationArea = All;
    //             RunObject = page 50146;

    //             RunPageLink = P_RFQ_Document_no = field(p_Customer_No);
    //             Image = Relationship;
    //             Promoted = true;
    //             PromotedCategory = Process;
    //             PromotedIsBig = true;

    //             trigger OnAction()
    //             begin
    //                 Message('Data Inserted ');
    //             end;
    //         }
    //         action("Posting Report")
    //         {
    //             ApplicationArea = All;
    //             //RunObject = report "RFQ REPORTS";
    //             trigger OnAction()
    //             var
    //                 Rec: Record " Posted RFQ header";



    //             begin
    //                 // Rec.FindLast;
    //                 Rec.FindSet();

    //                 Rec.setrange(P_RFQ_Document_no, Rec.P_RFQ_Document_no);
    //                 Rec.setrecfilter;
    //                 report.run(50146, true, false, Rec)

    //             end;







    //         }

    //     }

    // }




    var
        myInt: Integer;


}