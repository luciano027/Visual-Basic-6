Attribute VB_Name = "BackUpMySQL"
Const MSG_01 = "Bakck UP Criado por: "
Const MSG_02 = "Base de Dados: "
Const MSG_03 = "Inicio/Hora: "
Const MSG_04 = "DD/MM/YY HH:MM:SS"
Const MSG_05 = "DBMS: MySQL v"
Const MSG_06 = "Estrutura da tabela "
Const MSG_07 = "Dados da tabela "
Const MSG_08 = "Fim do Backup: "
Public bd As ADODB.Connection, sStop As Boolean
Attribute sStop.VB_VarUserMemId = 1073741824
Dim sBaseAtual As String
Attribute sBaseAtual.VB_VarUserMemId = 1073741826

Public Sub MySQLBackup(ByVal strNomeArquivo As String, cnn As ADODB.Connection, lst As ListView)
    frmBackup_restaura.btnBackup.Enabled = False
    Dim smsg As String
    Dim nodeX2 As ListItem
    lst.View = lvwReport
    lst.ListItems.Clear
    lst.ColumnHeaders.Clear
    lst.GridLines = False
    lst.ColumnHeaders.Add , , "log"
    lst.ColumnHeaders(1).Width = 5000

    On Error Resume Next

    Dim rss As ADODB.Recordset
    Dim rssAux As ADODB.Recordset

    Dim X As Long, i As Integer

    Dim strNomeTabela As String
    Dim strLinha As String
    Dim strBuffer As String
    Dim strNomeBase As String

    X = FreeFile
    Open strNomeArquivo For Output As X

    Print #X, ""
    Print #X, "#"
    With lst

        Print #X, "# " & MSG_01 & App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
        smsg = "# " & MSG_01 & App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
        Set nodeX2 = lst.ListItems.Add(, , smsg)
        nodeX2.EnsureVisible: DoEvents

        strNomeBase = Mid(cnn.ConnectionString, InStr(cnn.ConnectionString, "DATABASE=") + 9)
        strNomeBase = Left(strNomeBase, InStr(strNomeBase, ";") - 1)
        Print #X, "# " & MSG_02 & strNomeBase & ";"
        smsg = "# " & MSG_02 & strNomeBase
        Set nodeX2 = lst.ListItems.Add(, , smsg)
        nodeX2.EnsureVisible: DoEvents

        Set rss = New ADODB.Recordset
        Set rssAux = New ADODB.Recordset

        Print #X, "# " & MSG_03 & Format(Now, MSG_04)
        smsg = "# " & MSG_03 & Format(Now, MSG_04)
        Set nodeX2 = lst.ListItems.Add(, , smsg)
        nodeX2.EnsureVisible: DoEvents


        rss.Open "show variables like 'version';", cnn
        If Not rss.EOF Then
            Print #X, "# " & MSG_05 & rss.Fields(1)
            smsg = "# " & MSG_05 & rss.Fields(1)
            Set nodeX2 = lst.ListItems.Add(, , smsg)
            nodeX2.EnsureVisible: DoEvents

        End If
        rss.Close

        Print #X, "#"
        Print #X, ""
        Print #X, "SET FOREIGN_KEY_CHECKS=0;"
        smsg = "Desativando a checagem de Constraint;"
        Set nodeX2 = lst.ListItems.Add(, , smsg)
        nodeX2.EnsureVisible: DoEvents

        Print #X, ""
        Print #X, "DROP DATABASE IF EXISTS `" & strNomeBase & "`;"
        smsg = "Excluido o banco " & strNomeBase
        Set nodeX2 = lst.ListItems.Add(, , smsg)
        nodeX2.EnsureVisible: DoEvents

        Print #X, "CREATE DATABASE `" & strNomeBase & "`;"
        smsg = "Criando o Banco " & strNomeBase
        Set nodeX2 = lst.ListItems.Add(, , smsg)
        nodeX2.EnsureVisible: DoEvents

        Print #X, "USE `" & strNomeBase & "`;"
        smsg = "Executando o comando USE `" & strNomeBase & "`;"
        Set nodeX2 = lst.ListItems.Add(, , smsg)
        nodeX2.EnsureVisible: DoEvents

        strNomeTabela = ""

        With rss
            .Open "SHOW TABLE STATUS", cnn
            frmBackup_restaura.lblTotalTabelas = rss.RecordCount: DoEvents
            smsg = rss.RecordCount & " Tabelas no Banco " & strNomeBase
            Set nodeX2 = lst.ListItems.Add(, , smsg)
            nodeX2.EnsureVisible: DoEvents

            '  frmBackup_restaura.pb.Value = 0
            '  frmBackup_restaura.pb.Max = frmBackup_restaura.lblTotalTabelas
            Dim xpb2 As Integer
            Do While Not .EOF
                xpb2 = frmBackup_restaura.lblTabelaAtual
                ' frmBackup_restaura.pb.Value = xpb2

                If sStop = True Then
                    smsg = "Processo Interrompido pelo usuário"
                    Set nodeX2 = lst.ListItems.Add(, , smsg)
                    nodeX2.EnsureVisible: DoEvents
                    frmBackup_restaura.Label1 = "Statuso do Processo : " & smsg
                    frmBackup_restaura.btnBackup.Enabled = True
                    sStop = False
                    Exit Sub
                End If
                strNomeTabela = .Fields.Item("Name").Value
                If UCase(Mid(strNomeTabela, 1, 3)) <> "CEP" Then
                    '-----------------------------------------------------------------------------------------
                    If strNomeTabela = "vendas" Then
                        Beep
                    End If

                    frmBackup_restaura.lblNomeTabela = strNomeTabela: DoEvents
                    frmBackup_restaura.lblTabelaAtual = rss.AbsolutePosition: DoEvents
                    frmBackup_restaura.Label1 = "Status do Processo :"
                    With rssAux

                        .Open "SHOW CREATE TABLE " & strNomeTabela, cnn

                        Print #X, ""
                        Print #X, "#"
                        Print #X, "# " & MSG_06 & strNomeTabela & ""
                        smsg = "# " & MSG_06 & strNomeTabela & ""
                        Set nodeX2 = lst.ListItems.Add(, , smsg)
                        nodeX2.EnsureVisible: DoEvents
                        frmBackup_restaura.Label1 = Left(frmBackup_restaura.Label1 & " " & Mid(.Fields.Item(1).Value, 1, InStr(.Fields.Item(1).Value, "(") - 1), 255): DoEvents
                        Print #X, "#"

                        Print #X, .Fields.Item(1).Value & ";"


                        .Close

                    End With

                    With rssAux
                        .Open "SELECT * FROM " & strNomeTabela & "", cnn
                        frmBackup_restaura.lblQtdeRegistros = rssAux.RecordCount & " Registro(s)": DoEvents
                        smsg = "Selecionando a tabela " & strNomeTabela
                        Set nodeX2 = lst.ListItems.Add(, , smsg)
                        nodeX2.EnsureVisible: DoEvents

                        Print #X, ""
                        Print #X, "#"
                        Print #X, "# " & MSG_07 & strNomeTabela & ""
                        smsg = "# " & MSG_07 & strNomeTabela & ""
                        Set nodeX2 = lst.ListItems.Add(, , smsg)
                        nodeX2.EnsureVisible: DoEvents

                        Print #X, "#"
                        Print #X, "lock tables `" & strNomeTabela & "` write;"
                        smsg = "Bloqueando a tabela " & strNomeTabela & " contra gravação;"
                        Set nodeX2 = lst.ListItems.Add(, , smsg)
                        nodeX2.EnsureVisible: DoEvents


                        If Not .EOF Then
                            frmBackup_restaura.pb2.Max = .RecordCount
                            frmBackup_restaura.pb2.Value = 0
                            Dim xpb As Integer
                            xpb = frmBackup_restaura.pb2.Max / .RecordCount

                            Print #X, "INSERT INTO `" & strNomeTabela & "` VALUES "
                            smsg = "Inserindo os dados na tabela " & strNomeTabela
                            Set nodeX2 = lst.ListItems.Add(, , smsg)
                            nodeX2.EnsureVisible: DoEvents

                            Do While Not .EOF

                                On Error Resume Next
                                frmBackup_restaura.pb2.Value = frmBackup_restaura.pb2.Value + xpb: DoEvents
                                frmBackup_restaura.Label1 = "Inserindo Registro nº " & .AbsolutePosition
                                Err.Clear

                                strLinha = ""
                                For i = 0 To .Fields.Count - 1
                                    strBuffer = .Fields.Item(i).Value
                                    'Debug.Print .Fields.Item(i).Name

                                    If .Fields.Item(i).Type = 5 Then
                                        strBuffer = Replace(Format(strBuffer, "0.00"), ",", ".")
                                    End If

                                    If .Fields.Item(i).Type = 131 Then
                                        strBuffer = Replace(Format(strBuffer, "0.00"), ",", ".")
                                    End If

                                    If .Fields.Item(i).Type = 135 Then
                                        strBuffer = Format(strBuffer, "yyyy-MM-dd hh:mm:ss")
                                    End If

                                    strBuffer = Replace(strBuffer, "\", "\\")
                                    strBuffer = Replace(strBuffer, "'", "\'")
                                    strBuffer = Replace(strBuffer, Chr(10), "")
                                    strBuffer = Replace(strBuffer, Chr(13), "\r\n")

                                    If strLinha <> "" Then
                                        strLinha = strLinha & ", "
                                    End If
                                    strLinha = strLinha & "'" & strBuffer & "'"
                                Next i

                                .MoveNext

                                strLinha = "(" & strLinha & ")"
                                If .EOF Then
                                    Print #X, strLinha & ";"
                                Else
                                    Print #X, strLinha & ","
                                End If

                            Loop

                        End If

                        .Close
                    End With

                    Print #X, "unlock tables;"
                    smsg = "Desbloqueando a Tabela;"
                    Set nodeX2 = lst.ListItems.Add(, , smsg)
                    nodeX2.EnsureVisible: DoEvents

                    Print #X, "#--------------------------------------------"
                    smsg = "#--------------------------------------------"
                    Set nodeX2 = lst.ListItems.Add(, , smsg)
                    nodeX2.EnsureVisible: DoEvents
                End If
                .MoveNext

                '----------------------------------------------------------------------------------------------
            Loop

            Print #X, ""
            Print #X, "SET FOREIGN_KEY_CHECKS=1;"
            smsg = "Ativando a checagem de Constraint;"
            Set nodeX2 = lst.ListItems.Add(, , smsg)
            nodeX2.EnsureVisible: DoEvents

            Print #X, ""
            Print #X, "# " & MSG_08 & Format(Now, MSG_04)
            smsg = "# " & MSG_08 & Format(Now, MSG_04)
            Set nodeX2 = lst.ListItems.Add(, , smsg)
            nodeX2.EnsureVisible: DoEvents


            .Close
        End With

        Close #X
        smsg = "Concluído !!!!!"
        Set nodeX2 = lst.ListItems.Add(, , smsg)
        nodeX2.EnsureVisible: DoEvents
        frmBackup_restaura.Label1.Caption = "Status do Processo :" & smsg
        frmBackup_restaura.btnBackup.Enabled = True
        ' frmBackup_restaura.pb.Value = 22
        frmBackup_restaura.pb2.Value = 0
        MsgBox "Processo concluído !!!"
    End With
End Sub

Public Sub MySQLRestore(ByVal strNomeArquivo As String, cnn As ADODB.Connection, lst As ListView, pb As ProgressBar)
'  frmRestaura_backup.btnIniciarRestaura.Enabled = False
'  frmRestaura_backup.btnBackup.Enabled = False
    Dim smsg As String
    Dim nodeX2 As ListItem
    lst.View = lvwReport
    lst.ListItems.Clear
    lst.ColumnHeaders.Clear
    lst.GridLines = False
    lst.ColumnHeaders.Add , , "log"
    lst.ColumnHeaders(1).Width = 5000

    Dim lngTotalBytes As Long, lngCurrentBytes As Long
    Dim X As Integer, strLinha As String, strAux As String
    Dim blnPassLines As Boolean
    Dim blnAnalizeIt As Boolean

    X = FreeFile

    On Error GoTo ErrDrv

    Open strNomeArquivo For Input As #X
    lngTotalBytes = LOF(X)
    smsg = "Abrindo Arquivo de Backup"
    Set nodeX2 = lst.ListItems.Add(, , smsg)
    nodeX2.EnsureVisible: DoEvents

    blnPassLines = False
    frmRestaura_backup.pb.Max = lngTotalBytes
    frmRestaura_backup.pb.Value = 0
    Dim xpb As Long
    Dim LINHA As String
    Dim IZ As Integer
    Do While Not EOF(X)
        IZ = IZ + 1

        Line Input #X, strLinha

        If UCase(Left(strLinha, 16)) = UCase("# Base de Dados:") Then
            sNomeBase = Mid(strLinha, InStr(strLinha, "# Base de Dados: ") + 17)
            sNomeBase = "`" & Left(sNomeBase, InStr(sNomeBase, ";") - 1) & "`"

            If frmRestaura_backup.optNovoBanco.Value = True Then
                sBaseReplace = " `" & frmRestaura_backup.txtNovoBanco & "`"
            ElseIf frmRestaura_backup.optOriginal.Value = True Then
                sBaseReplace = sNomeBase
            ElseIf frmRestaura_backup.optOutroBanco.Value = True Then
                sBaseReplace = " `" & frmRestaura_backup.cmbBanco & "`"    'Mid(strLinha, InStr(strLinha, "DROP DATABASE IF EXISTS `") + 24)
            End If

        End If





        Select Case IZ

        Case 11
            If Left(UCase(strLinha), 4) = "DROP" Then
                strLinha = Replace(strLinha, sNomeBase, sBaseReplace)
            End If



        Case 12
            If Left(UCase(strLinha), 6) = "CREATE" Then
                strLinha = Replace(strLinha, sNomeBase, sBaseReplace)
            End If



        Case 13
            If Left(UCase(strLinha), 3) = "USE" Then
                strLinha = Replace(strLinha, sNomeBase, sBaseReplace)
            End If

        End Select


        lngCurrentBytes = lngCurrentBytes + Len(strLinha)

        frmRestaura_backup.lbltamanhoArquivo = lngTotalBytes & " Byte(s)": DoEvents
        frmRestaura_backup.lblProcessoatual = lngCurrentBytes
        xpb = frmRestaura_backup.pb.Max / lngTotalBytes
        frmRestaura_backup.pb.Value = lngCurrentBytes: DoEvents

        blnAnalizeIt = True
        strLinha = Trim(strLinha)
        If Not blnPassLines Then
            If Left(strLinha, 1) = "#" Then
                blnAnalizeIt = False
            ElseIf Left(strLinha, 2) = "/*" Then
                blnAnalizeIt = False
                blnPassLines = True
            End If
        ElseIf Right(Trim(strLinha), 2) = "*/" Then
            blnPassLines = False
            blnAnalizeIt = False
        End If

        If blnAnalizeIt And strLinha <> "" Then

            While Mid(strLinha, Len(strLinha), 1) <> ";"
                strAux = strLinha
                Line Input #X, strLinha
                lngCurrentBytes = lngCurrentBytes + Len(strLinha)
                strLinha = Trim(strLinha)
                strLinha = strAux & strLinha
            Wend

            smsg = "Executando comando " & Left(strLinha, 255)
            Set nodeX2 = lst.ListItems.Add(, , smsg)
            nodeX2.EnsureVisible: DoEvents

            cnn.Execute strLinha

        End If

    Loop

    Close #X

    frmRestaura_backup.lblProcessoatual = lngTotalBytes
    smsg = "Processo concluído com sucesso !!!"
    MsgBox smsg, vbInformation
    Set nodeX2 = lst.ListItems.Add(, , smsg)
    nodeX2.EnsureVisible: DoEvents
    frmRestaura_backup.btnIniciarRestaura.Enabled = True
    frmRestaura_backup.btnBackup.Enabled = True
    Exit Sub
ErrDrv:

    smsg = "ERROR:" & Err.Number & vbNewLine & Err.Description & vbNewLine
    Set nodeX2 = lst.ListItems.Add(, , smsg)
    nodeX2.EnsureVisible: DoEvents

    Err.Clear

End Sub



