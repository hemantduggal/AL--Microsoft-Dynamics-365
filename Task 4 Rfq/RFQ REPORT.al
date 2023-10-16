// /// <summary>
// /// Report RFQ REPORTS (ID 50144).
// /// </summary>
// report 50146 "RFQ REPORTS"
// {
//     UsageCategory = Administration;
//     ApplicationArea = All;
//     DefaultLayout = RDLC;
//     RDLCLayout = 'ReportLayout\Posted RFQ Report.rdl';


//     dataset
//     {
//         dataitem(" Posted RFQ Header"; " Posted RFQ Header")
//         {
//             // RequestFilterFields = p_Customer_No, P_Location_Code;


//             column(P_RFQ_Document_no; P_RFQ_Document_no)
//             {

//             }
//             column(P_Document_Date; P_Document_Date)
//             {

//             }
//             column(P_Posting_date; P_Posting_date)
//             {

//             }
//             column(p_Customer_No; p_Customer_No)
//             {

//             }
//             column(P_Customer_name; P_Customer_name)
//             {

//             }
//             column(P_Location_Code; P_Location_Code)
//             {

//             }


//         }


//         dataitem(" Posted RFQ Line"; " Posted RFQ Line")
//         {
//             RequestFilterFields = L_RFQ_Type;




//             column(L_RFQ_Document_no; L_RFQ_Document_no)
//             {

//             }
//             column(L_RFQ_Type; L_RFQ_Type)
//             {

//             }
//             column(L_item_no; L_item_no)
//             {

//             }
//             column(L_Decription; L_Decription)
//             {

//             }
//             column(L_Quantity; L_Quantity)
//             {

//             }
//             column(L_UOM; L_UOM)
//             {

//             }
//             column(L_Unit_price; L_Unit_price)
//             {

//             }
//             column(L_Customer_No; L_Customer_No)
//             {

//             }
//             column(L_Location_Code; L_Location_Code)
//             {

//             }
//             column(L_Posting_date; L_Posting_date)
//             {

//             }
//         }
//     }

//     requestpage
//     {
//         layout
//         {
//             area(Content)
//             {
//                 group(Options)
//                 {
//                     field(p_Customer_No; " Posted RFQ Header".p_Customer_No)
//                     {
//                         ApplicationArea = all;
//                         TableRelation = " Posted RFQ Header";



//                     }
//                     field(P_Location_Code; " Posted RFQ Header".P_Location_Code)
//                     {
//                         ApplicationArea = all;
//                         TableRelation = " Posted RFQ Header";


//                     }
//                 }
//             }
//         }

//         actions
//         {
//             area(processing)
//             {
//                 action(Post_RFQ_Report)
//                 {
//                     ApplicationArea = All;


//                 }

//             }
//         }
//     }



//     var
//         myInt: Integer;
// }