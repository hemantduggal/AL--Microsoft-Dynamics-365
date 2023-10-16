/// <summary>
/// PageExtension ItemExt (ID 50149) extends Record Sales Order List.
/// </summary>

pageextension 50148 salesExt extends "Sales Order List"
{



    actions
    {
        addafter("F&unctions")
        {

            action(Importlines)
            {
                Caption = 'import sales order';
                Promoted = true;
                PromotedCategory = Process;
                Image = Import;
                ApplicationArea = All;
                trigger OnAction()
                begin

                    Xmlport.Run(50105, false, true);
                end;
            }
            action(exportItems)
            {
                Caption = 'export sales list';
                Promoted = true;
                PromotedCategory = Process;
                Image = Export;
                ApplicationArea = All;
                trigger OnAction()
                begin
                    Xmlport.Run(Xmlport::Exportsalesorderlist, true, false);
                end;
            }
            action("import via excel buffer")
            {
                Caption = 'Import via excel buffer';
                Image = ImportExcel;
                Promoted = true;
                PromotedCategory = Process;
                ApplicationArea = All;
                ToolTip = 'import data from excel ';
                trigger OnAction()
                var
                begin
                    ReadExcelSheet();
                    importexceldata();

                end;
            }

            action("export via excel buffer")
            {
                Caption = 'export';
                Image = Export;
                Promoted = true;
                PromotedCategory = Process;
                ApplicationArea = all;
                trigger OnAction()
                var
                begin
                    ExportExcelEntries(Rec);

                end;
            }


        }

    }
    var
        transname: Code[10];
        filename: text[100];
        sheetname: text[100];
        tempexcelbuffer: Record "Excel Buffer" temporary;
        uploadmsg: Label 'please choose the excel file ';
        nofilemsg: Label 'no excel file found';
        batchisblankmsg: label 'transaction name is blank ';
        excelimportsuccess: Label 'excel imported successfully';

    local procedure ReadExcelSheet()
    var
        myInt: Integer;
        filemanagement: Codeunit "File Management";
        istream: InStream;
        fromfile: Text[100];
    begin
        UploadIntoStream(uploadmsg, '', '', fromfile, istream);
        if fromfile <> '' then begin
            filename := filemanagement.GetFileName(fromfile);
            sheetname := tempexcelbuffer.SelectSheetsNameStream(istream);

        end else
            Error(nofilemsg);
        tempexcelbuffer.Reset();
        tempexcelbuffer.DeleteAll();
        tempexcelbuffer.OpenBookStream(istream, sheetname);
        tempexcelbuffer.ReadSheet();

    end;



    local procedure GetValueAtCell(RowNo: integer; ColNo: Integer): Text
    begin
        tempexcelbuffer.Reset();
        if tempexcelbuffer.get(RowNo, ColNo) then
            exit(tempexcelbuffer."Cell Value as Text")
        else
            exit('');


    end;

    local procedure importexceldata()
    var
        myInt: Integer;
        gsimportbuffer: Record "Sales Header";
        rowno: integer;
        colno: integer;
        lineno: Integer;
        no: Code[20];
        maxrow: Integer;
    begin
        rowno := 0;
        colno := 0;
        maxrow := 0;
        lineno := 0;

        gsimportbuffer.Reset();
        if gsimportbuffer.FindLast() then
            no := gsimportbuffer."No.";
        tempexcelbuffer.Reset();
        if tempexcelbuffer.FindLast() then begin
            maxrow := tempexcelbuffer."Row No.";

        end;
        for rowno := 2 to maxrow do begin
            //     lineno := lineno + 10000;
            no := no;


            GSImportBuffer.Init();

            GSImportBuffer."Line No." := LineNO;

            Evaluate(GSImportBuffer."No.", GetValueAtCell(RowNo, 1));
            Evaluate(GSImportBuffer."Sell-to Customer No.", GetValueAtCell(RowNo, 2));
            Evaluate(GSImportBuffer."Sell-to Customer Name", GetValueAtCell(RowNo, 3));
            Evaluate(GSImportBuffer."External Document No.", GetValueAtCell(RowNo, 4));
            Evaluate(GSImportBuffer."Location Code", GetValueAtCell(RowNo, 5));
            Evaluate(GSImportBuffer."Document Date", GetValueAtCell(RowNo, 7));
            Evaluate(GSImportBuffer.Status, GetValueAtCell(RowNo, 8));
            Evaluate(GSImportBuffer."Combine Shipments", GetValueAtCell(RowNo, 9));
            Evaluate(GSImportBuffer."Amt. Ship. Not Inv. (LCY)", GetValueAtCell(RowNo, 10));
            Evaluate(GSImportBuffer."Amt. Ship. Not Inv. (LCY) Base", GetValueAtCell(RowNo, 11));
            Evaluate(GSImportBuffer.Amount, GetValueAtCell(RowNo, 12));
            Evaluate(GSImportBuffer."Amount Including VAT", GetValueAtCell(RowNo, 13));

            GSImportBuffer.Insert();


        end;
        Message(ExcelImportSuccess);
    end;




    local procedure ExportExcelEntries(var GSExel: Record "Sales Header")
    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
        GSEntriesLbl: Label 'GS sales order Excel Entries';
        ExcelFileName: Label 'GSExcel Entries_%1_%2';
    begin
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        TempExcelBuffer.NewRow();

        TempExcelBuffer.AddColumn(GSExel.FieldCaption("No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GSExel.FieldCaption("Sell-to Customer No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GSExel.FieldCaption("Sell-to Customer Name"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GSExel.FieldCaption("External Document No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GSExel.FieldCaption("Location Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GSExel.FieldCaption("Document Date"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GSExel.FieldCaption(Status), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GSExel.FieldCaption("Combine Shipments"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GSExel.FieldCaption("Amt. Ship. Not Inv. (LCY)"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GSExel.FieldCaption("Amt. Ship. Not Inv. (LCY) Base"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GSExel.FieldCaption(Amount), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GSExel.FieldCaption("Amount Including VAT"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);

        if GSExel.FindSet() then
            repeat
                TempExcelBuffer.NewRow();
                TempExcelBuffer.AddColumn(GSExel."No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GSExel."Sell-to Customer No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GSExel."Sell-to Customer Name", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GSExel."External Document No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GSExel."Location Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GSExel."Document Date", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Date);
                TempExcelBuffer.AddColumn(GSExel.Status, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::text);
                TempExcelBuffer.AddColumn(GSExel."Combine Shipments", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::text);
                TempExcelBuffer.AddColumn(GSExel."Amt. Ship. Not Inv. (LCY)", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
                TempExcelBuffer.AddColumn(GSExel."Amt. Ship. Not Inv. (LCY) Base", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
                TempExcelBuffer.AddColumn(GSExel.Amount, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
                TempExcelBuffer.AddColumn(GSExel."Amount Including VAT", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);


            until GSExel.Next() = 0;


        TempExcelBuffer.CreateNewBook(GSEntriesLbl);
        TempExcelBuffer.WriteSheet(GSEntriesLbl, CompanyName, UserId);
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileName, CurrentDateTime, UserId));
        TempExcelBuffer.OpenExcel();

    end;




}