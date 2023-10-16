/// <summary>
/// PageExtension salesorderext (ID 50107) extends Record Sales Order.
/// </summary>
pageextension 50107 salesorderextenstion extends "Sales Order Subform"
{
    layout
    {
        // Add changes to page layout here
        addafter("No.")
        {
            field(Item_no; Rec.Item_no)
            {
                ApplicationArea = all;

            }
            field(sales_type; Rec.sales_type)
            {
                ApplicationArea = all;
            }
            field(sales_code; Rec.sales_code)
            {
                ApplicationArea = all;
            }
            field(location; Rec.location)
            {
                ApplicationArea = all;
            }
            field(start_date; Rec.start_date)
            {
                ApplicationArea = all;
            }
            field(end_date; Rec.end_date)
            {
                ApplicationArea = all;
            }
            field("Unit Cost"; Rec."Unit Cost") { ApplicationArea = all; }

            field(StudentNO; Rec.StudentNO)
            {
                ApplicationArea = all;
            }
            field("External Doc no"; Rec."External Doc no")
            {
                ApplicationArea = all;
                caption = 'External doc no';
            }

            field(StudentSubject; Rec.StudentSubject)
            {
                Caption = 'StudentSubjects';
                ApplicationArea = All;


                // trigger OnAfterLookup(Selected: RecordRef)
                // var
                //     salesRec: Record "Sales Line";
                // begin
                //     salesRec.Reset();
                //     if page.RunModal(Page::"Sales Order", salesRec) = Action::LookupOK then
                //         REc.StudentSubject := salesRec.StudentSubject;

                // end;
            }
            field(Marks; Rec.Marks)
            {
                Caption = 'marks';
                ApplicationArea = All;


                trigger OnValidate();
                begin

                end;
            }
            field("Date of birth"; Rec."Date of birth")
            {
                Caption = 'dateofbirth';
                ApplicationArea = All;


                trigger OnValidate();
                begin

                end;
            }
            field(Age; Rec.Age)
            {
                Caption = 'Age';
                ApplicationArea = All;


                trigger OnValidate();
                begin

                end;
            }
            field(item_category_code; Rec.item_category_code)
            {
                ApplicationArea = all;
                trigger OnValidate();
                begin

                end;

            }
            field(HEat_no; Rec.HEat_no)
            {
                ApplicationArea = all;
                trigger OnValidate();
                begin

                end;

            }
            field(Billet_no; Rec.Billet_no)
            {
                ApplicationArea = all;
                trigger OnValidate();
                begin

                end;

            }
            field(Size; Rec.Size)
            {
                ApplicationArea = all;
                trigger OnValidate();
                begin

                end;
            }
            field(Dia; Rec.Dia)
            {
                ApplicationArea = all;
                trigger OnValidate();
                begin

                end;
            }




        }
    }
    actions
    {
        addafter("F&unctions")
        {
            action(ImportSalesLine)
            {
                Caption = 'import sales line through xml port';
                // Promoted = true;
                // PromotedCategory = Process;
                Image = Import;
                ApplicationArea = all;
                trigger OnAction()
                begin
                    Xmlport.Run(Xmlport::"Sales Price Xml port", false, true);
                end;

            }
            action("Export sales line")
            {
                caption = 'export sales line';
                Image = Export;
                ApplicationArea = all;



                trigger OnAction()
                begin

                    xmlport.Run(Xmlport::"export Sales line", true, false);
                end;
            }
        }
    }
}