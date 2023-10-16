/// <summary>
/// PageExtension MyExtension (ID 50102) extends Record MyTargetPage.
/// </summary>
pageextension 50110 Purchaseorderext extends "Purchase Order"
{
    layout
    {
        // Add changes to page layout here
        addafter(Status)
        {
            field(StudentNO; Rec.StudentNO)
            {
                Caption = 'StudentNo';
                ApplicationArea = All;


                trigger OnValidate();
                begin

                end;
            }
            field(StudentName; rec.StudentName)
            {
                Caption = 'StudentName';
                ApplicationArea = All;


                trigger OnValidate();
                begin

                end;
            }
            field(StudentAddress; Rec.StudentAddress)
            {
                Caption = 'StudentAddress';
                ApplicationArea = All;


                trigger OnValidate();
                begin

                end;
            }
            field(StudentPhone; Rec.StudentPhone)
            {
                Caption = 'StudentPhone';
                ApplicationArea = All;


                trigger OnValidate();
                begin

                end;
            }
            field(StudentState; Rec.StudentState)
            {
                Caption = 'StudentState';
                ApplicationArea = All;


                trigger OnValidate();
                begin

                end;
            }
            field(StudentCity; Rec.StudentCity)
            {
                Caption = 'StudentCity';
                ApplicationArea = All;


                trigger OnValidate();
                begin

                end;
            }
            field("Contact person"; Rec."Contact person")
            {
                ApplicationArea = all;

            }
            field(Remarks; Rec.Remarks)
            {
                ApplicationArea = all;

            }
            field(Deal_date; Rec.Deal_date)
            {
                ApplicationArea = all;
            }
            field("order category"; Rec."order category")
            {
                ApplicationArea = all;

            }
            field("Index no"; Rec."Index no")
            {
                ApplicationArea = all;

            }
            field(Plank_no; Rec.Plank_no)
            {
                ApplicationArea = all;

            }
            field("Currency Capacity"; Rec."Currency Capacity")
            {
                ApplicationArea = all;

            }
            field("unit od specification"; Rec."unit od specification")
            {
                ApplicationArea = all;

            }




        }
    }

    actions
    {
        addafter("F&unctions")
        {


            action("Export purchase order")
            {
                caption = 'export purchase order order';
                Image = Export;
                ApplicationArea = all;
                // Visible = s."Safety Guidelines" = true;
                //Enabled = s."Safety Guidelines" = true;
                Visible = flag;
                Enabled = flag;



                trigger OnAction()
                begin


                    xmlport.Run(Xmlport::Exportpurchaseorder, true, false);
                end;
            }// Add changes to page actions here
        }
    }



    var
        myInt: Integer;
        s: record SafetyGuidlines;
        flag: Boolean;
}