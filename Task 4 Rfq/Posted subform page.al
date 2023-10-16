/// <summary>
/// Page RFQ subform page (ID 50148).
/// </summary>
page 50012 "Sale Posted RFQ subform page"
{
    PageType = ListPart;
    ApplicationArea = All;
    UsageCategory = Documents;
    SourceTable = 50014;
    Caption = 'Sale Posted RFQsubform page';
    Description = 'This is a  Sale Posted subfrom page ';
    AboutText = ' Sale Posted RFQSubform PAge ';
    AboutTitle = 'This Is a  Sale Posted RFQSubform Page ';

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
                    NotBlank = true;

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





}