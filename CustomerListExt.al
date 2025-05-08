pageextension 50100 CustomerListExt extends "Customer List"
{
    actions
    {
        addfirst(processing)
        {
            action(ConvertWordToPDF)
            {
                ApplicationArea = All;
                Caption = 'Convert Word To PDF';
                Image = Change;
                Promoted = true;
                PromotedCategory = Process;
                trigger OnAction()
                var
                    WordFileName: Text[100];
                    PDFFileName: Text[100];
                    InStr: InStream;
                    OutStr: OutStream;
                    FileMgt: Codeunit "File Management";
                    TempBlob: Codeunit "Temp Blob";
                    DocumentReportMgt: Codeunit "Document Report Mgt.";
                begin
                    TempBlob.CreateOutStream(OutStr);
                    FileMgt.BLOBImport(TempBlob, WordFileName);
                    DocumentReportMgt.ConvertWordToPdf(TempBlob, 0);
                    TempBlob.CreateInStream(InStr);
                    PDFFileName := 'Converted' + '.pdf';
                    DownloadFromStream(InStr, '', '', '', PDFFileName);
                end;
            }
        }
    }
}
