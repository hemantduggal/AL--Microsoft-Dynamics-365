/// <summary>
/// Page RFQ subform page (ID 50143).
/// </summary>
page 50003 "Sale RFQ subform page"
{
    PageType = ListPart;
    ApplicationArea = All;
    UsageCategory = Documents;
    SourceTable = 50002;
    Caption = 'Sale RFQsubform page';
    Description = 'This is a subfrom page ';
    AboutText = 'Sale RFQSubform PAge ';
    AboutTitle = 'This Is a  Sale RFQSubform Page ';


    layout
    {
        area(Content)
        {
            repeater(General)
            {
                Caption = 'general';
                field(Customer_No; Rec.Customer_No)
                {
                    ApplicationArea = All;

                }

                field("No."; Rec."No.")
                {
                    ApplicationArea = All;

                }
                field(item_no; Rec.item_no)
                {
                    ApplicationArea = All;

                }

                field(Decription; Rec.Decription)
                {
                    ApplicationArea = All;

                }
                field(Quantity; Rec.Quantity)
                {
                    ApplicationArea = All;

                }
                field(UOM; Rec.UOM)
                {
                    ApplicationArea = All;

                }
                field(Direct_unit_cost; Rec.Direct_unit_cost)
                {
                    ApplicationArea = All;

                }
                field(rate; Rec.rate)
                {
                    ApplicationArea = All;

                }
                field(brand; Rec.brand)
                {
                    ApplicationArea = All;

                }

                field(salesperson; Rec.salesperson)
                {
                    ApplicationArea = All;

                }
                field(Application_master; Rec.Application_master)
                {
                    ApplicationArea = All;

                }
                field(Project_master; Rec.Project_master)
                {
                    ApplicationArea = All;

                }





            }
        }
    }

    actions
    {
        area(Processing)
        {
            action(Post_line)
            {
                ApplicationArea = All;
                //RunObject = page "Saless Posted RFQ Card page";
                // RunPageLink = "No." = field("No.");
                Image = PostOrder;

                Ellipsis = true;



                trigger OnAction()
                var
                    rfq: Record "Saless Posted RFQ Line";
                begin

                    rfq.Reset();

                    //rfq.SetRange("No.", rec."No.");
                    if rfq.FindSet then
                        rfq."No." := rec."No.";
                    //rfq.Customer_No := rec.Customer_No;
                    rfq.item_no := rec.item_no;
                    rfq.Decription := rec.Decription;
                    rfq.Quantity := rec.Quantity;
                    rfq.UOM := rec.UOM;
                    rfq.Direct_unit_cost := rec.Direct_unit_cost;
                    rfq.rate := rec.rate;
                    rfq.brand := rec.brand;
                    rfq.salesperson := rec.salesperson;
                    rfq.Application_master := rec.Application_master;
                    rfq.Project_master := rec.Project_master;


                    rfq.Modify();
                    Message('Data inserted');


                end;
            }



        }
    }

    var
        myInt: Integer;


}