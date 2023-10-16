/// <summary>
/// Report Sales order Report (ID 50126).
/// </summary>
report 50115 "Sale order Rep"
{
    UsageCategory = Administration;
    Caption = 'Sale order Rep';
    ApplicationArea = All;
    DefaultLayout = RDLC;
    RDLCLayout = 'ReportLayout\SAle order.rdl';

    dataset
    {
        dataitem("Sales Invoice Header"; "Sales Invoice Header")
        {
            RequestFilterFields = "posting date";
            DataItemTableView = sorting("Posting Date");

            column(starting_date; starting_date)
            {

            }
            column(ending_date; ending_date)
            {

            }


            column(CompanyPicture; CompanyInfo.Picture)
            {

            }
            column(CompanyTelephone; CompanyInfo."Phone No.")
            {

            }
            column(CompanyName; CompanyInfo.Name)
            {

            }
            column(CompanyAddress; CompanyInfo.Address)
            {

            }
            column(CompanyAddress2; CompanyInfo."Address 2")
            {

            }
            column(city; companyinfo.City)
            {

            }
            column(state; companyinfo."State Code")
            {

            }
            column(companyEmail; CompanyInfo."E-Mail")
            {

            }


            column(Posting_Date; "Posting Date")
            {

            }
            column(No_; "No.")
            {

            }
            column(Sell_to_Contact; "Sell-to Contact")
            {

            }
            column(Amount; Amount)
            {

            }
            column(Remaining_Amount; "Remaining Amount")
            {

            }
            column(total; total)
            {

            }
            column(col1Months; col1months)
            {

            }




            dataitem("Sales Invoice Line"; "Sales Invoice Line")
            {

                //DataItemTableView = sorting("Document No.", "Line No.");
                // DataItemLink = "Document No." = field("No.");

                column(Unit_Cost; "Unit Cost")
                {

                }












            }
        }
    }

    requestpage
    {
        layout
        {
            area(Content)
            {
                group(GroupName)
                {
                    field(starting_date; starting_date)
                    {
                        ApplicationArea = Basic, Suite;
                        Caption = 'Starting Date';
                        trigger OnValidate()
                        begin

                        end;
                    }
                    field(ending_date; ending_date)
                    {
                        ApplicationArea = Basic, Suite;
                        Caption = 'Ending Date';
                        trigger OnValidate()
                        begin

                        end;
                    }
                }
            }
        }


        actions
        {
            area(processing)
            {
                action(ActionName)
                {
                    ApplicationArea = All;

                }
            }
        }
    }
    trigger OnPreReport()
    var
    begin
        CompanyInfo.Get;
        CompanyInfo.CalcFields(CompanyInfo.Picture);
    end;

    var
        myInt: Integer;
        CompanyInfo: Record "Company Information";
        ending_date: Date;
        starting_date: Date;
        total: Decimal;







        Recsales: Record "Sales Invoice Header";

        col1Months: Date;
        col2Months: Date;
        col3Months: Date;

        NoOfMonths: Integer;



        TOtalMonths1: Date;


        TOtalAMtcolwise: Decimal;
        sales: Record "Sales Invoice Header";
        Balance: Decimal;



}