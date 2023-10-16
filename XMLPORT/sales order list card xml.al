/// <summary>
/// XmlPort importXmlport (ID 50101).
/// </summary>
xmlport 50101 importXmlport
{

    schema
    {
        textelement(root)
        {
            tableelement("Header"; "Sales Header")
            {
                fieldattribute(no; Header."No.")
                {

                }
                fieldattribute(cust_no; Header."Sell-to Customer No.")
                {

                }
                fieldattribute(name; Header."Sell-to Customer Name")
                {

                }
                fieldattribute(ext; Header."External Document No.")
                {

                }
                fieldattribute(loc_code; Header."Location Code")
                {

                }

                fieldattribute(doc_date; Header."Document Date")
                {

                }
                fieldattribute(status; Header.Status)
                {

                }
                fieldattribute(ship; Header."Combine Shipments")
                {

                }
                fieldattribute(am_ship; Header."Amt. Ship. Not Inv. (LCY)")
                {

                }
                fieldattribute(am2_ship; Header."Amt. Ship. Not Inv. (LCY) Base")
                {

                }
                // tableelement("line"; "Sales line")
                // {
                //     fieldattribute(amount; line.Amount)
                //     {

                //     }
                // }






            }
        }
    }

    requestpage
    {
        layout
        {
            area(content)
            {
                group(GroupName)
                {

                }
            }
        }

        actions
        {
            area(processing)
            {
                action(ActionName)
                {

                }
            }
        }
    }

    var
        myInt: Integer;
}