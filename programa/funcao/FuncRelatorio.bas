Attribute VB_Name = "FuncRelatorio"
Public Sub ConfiguraIcones(ByRef arpRelatorio As ActiveReport)
    With arpRelatorio.Toolbar.Tools
        .Item(2).Caption = ""
        .Item(2).Style = 0
        .Item(2).Tooltip = "Imprimir"
        .Item(3).Visible = False
        .Item(4).Tooltip = "Copiar"
        .Item(4).Caption = ""
        .Insert 5, ""
        .Item(5).Style = 0
        .Item(5).Tooltip = "Exportar"
        .Item(5).AddIcon LoadPicture(salvaric)
        .Item(6).Visible = False
        .Item(7).Tooltip = "Procurar"
        .Item(9).Tooltip = "Uma página"
        .Item(10).Tooltip = "Várias páginas"
        .Item(12).Tooltip = "Reduzir"
        .Item(13).Tooltip = "Aumentar"
        .Item(16).Tooltip = "Página anterior"
        .Item(17).Tooltip = "Página seguinte"
        .Item(20).Tooltip = "Voltar"
        .Item(20).Caption = ""
        .Item(21).Tooltip = "Avançar"
        .Item(21).Caption = ""
        .Add (" Enviar via E-mail ")
        .Item(22).Tooltip = "Enviar via E-mail"
        .Add ("  Fechar Relatório  ")
        .Item(23).Tooltip = "Fechar o Relatorio"

    End With
End Sub

Public Sub ExportaRelatorio(ByVal cdgRelatorio As Control, arpRelatorio As ActiveReport)
    cdgRelatorio.FileName = ArqEmail
    cdgRelatorio.Filter = "Formato Rich Text (.rtf)|*.rtf|Documento de texto (.txt)|*.txt|Planilha do Microsoft Excel (*.xls)|*.xls|Documento do Adobe Acrobat (.pdf)|*.pdf|Documento HTML (*.htm)|*.htm|Formato TIFF (*.tif)|*.tif"
    cdgRelatorio.ShowSave
    Select Case UCase$(Right$(cdgRelatorio.FileName, 3))
    Case "RTF"
        Call SalvaRTF(cdgRelatorio.FileName, arpRelatorio)
    Case "TXT"
        Call SalvaTXT(cdgRelatorio.FileName, arpRelatorio)
    Case "PDF"
        Call SalvaPDF(cdgRelatorio.FileName, arpRelatorio)
    Case "XLS"
        Call SalvaXLS(cdgRelatorio.FileName, arpRelatorio)
    End Select
End Sub
Public Sub EnviarPeloEmail(ByVal cdgRelatorio As Control, arpRelatorio As ActiveReport)
    cdgRelatorio.FileName = ArqEmail
    cdgRelatorio.Filter = "Documento do Adobe Acrobat (.pdf)|*.pdf"
    ' cdgRelatorio.
    Call SalvaPDF(cdgRelatorio.FileName, arpRelatorio)
    With frmEnviarPeloEmail
        .lblArquivo.Caption = ArqEmail
        .Show 1
    End With
End Sub


Public Sub SalvaPDF(ByVal strArquivo As String, ByVal arpRelatorio As ActiveReport)
    Dim expRelatorio As ActiveReportsPDFExport.ARExportPDF

    Set expRelatorio = New ActiveReportsPDFExport.ARExportPDF
    expRelatorio.FileName = strArquivo
    expRelatorio.Export arpRelatorio.Pages
    Set expRelatorio = Nothing
End Sub


Public Sub SalvaRTF(ByVal strArquivo As String, ByVal arpRelatorio As ActiveReport)
    Dim expRelatorio As ActiveReportsRTFExport.ARExportRTF

    Set expRelatorio = New ActiveReportsRTFExport.ARExportRTF
    expRelatorio.FileName = strArquivo
    expRelatorio.Export arpRelatorio.Pages
    Set expRelatorio = Nothing
End Sub

Public Sub SalvaTXT(ByVal strArquivo As String, ByVal arpRelatorio As ActiveReport)
    Dim expRelatorio As ActiveReportsTextExport.ARExportText '

    Set expRelatorio = New ActiveReportsTextExport.ARExportText
    expRelatorio.FileName = strArquivo
    expRelatorio.Export arpRelatorio.Pages
    Set expRelatorio = Nothing


End Sub

Public Sub SalvaXLS(ByVal strArquivo As String, ByVal arpRelatorio As ActiveReport)
    Dim expRelatorio As ActiveReportsExcelExport.ARExportExcel

    Set expRelatorio = New ActiveReportsExcelExport.ARExportExcel
    expRelatorio.FileName = strArquivo
    expRelatorio.Export arpRelatorio.Pages
    Set expRelatorio = Nothing
End Sub
