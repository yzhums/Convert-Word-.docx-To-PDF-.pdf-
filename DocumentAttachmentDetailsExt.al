pageextension 50112 DocumentAttachmentDetailsExt extends "Document Attachment Details"
{
    actions
    {
        addafter(Preview_Promoted)
        {
            actionref(ConvertWordToPDF_Promoted; ConvertWordToPDF)
            {
            }
        }
        addfirst(processing)
        {
            action(ConvertWordToPDF)
            {
                Caption = 'Convert Word To PDF';
                ApplicationArea = All;
                Image = View;

                trigger OnAction()
                var
                    DocAttach: Record "Document Attachment";
                    DocAttachPDF: Record "Document Attachment";
                    InStr: InStream;
                    OutStr: OutStream;
                    TempBlob: Codeunit "Temp Blob";
                    DocumentReportMgt: Codeunit "Document Report Mgt.";
                    FileName: Text;
                begin
                    DocAttach.Reset();
                    CurrPage.SetSelectionFilter(DocAttach);
                    if DocAttach.FindFirst() then
                        if DocAttach."Document Reference ID".HasValue then begin
                            TempBlob.CreateOutStream(OutStr);
                            DocAttach."Document Reference ID".ExportStream(OutStr);
                            DocumentReportMgt.ConvertWordToPdf(TempBlob, 0);
                            TempBlob.CreateInStream(InStr);
                            FileName := DocAttach."File Name" + ' (1)';
                            DocAttachPDF.Init();
                            DocAttachPDF.Validate("Table ID", DocAttach."Table ID");
                            DocAttachPDF.Validate("No.", DocAttach."No.");
                            DocAttachPDF.Validate("File Name", FileName);
                            DocAttachPDF.Validate("File Extension", 'pdf');
                            DocAttachPDF."Document Reference ID".ImportStream(InStr, FileName);
                            DocAttachPDF.Insert(true);
                        end;
                end;
            }
        }
    }
}
