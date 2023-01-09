Attribute VB_Name = "mySQL"
Public Server As String
Public username As String
Public password As String
Public Database As String
Public Last_Query As String
Public Array_Query As Variant
Public banco As ADODB.Connection

Function Conectar() As Boolean
    Dim Server As String
    Dim username As String
    Dim password As String
    Dim Database As String

    ' Server = "estagiaria"  '"localhost"
    ' username = "micro1" '"root"

    Server = ReadINI("Servidor", "Server", App.Path & "\vendas.ini")
    username = ReadINI("Servidor", "username", App.Path & "\vendas.ini")
    password = "2562"
    Database = "Papelaria"

    Conectar = True
    On Error GoTo Err

    ' string de conexao
    strConnect = "Driver={MySQL ODBC 5.1 Driver};" _
               & " Server= " & Server & ";" _
               & " Database=" & Database & ";" _
               & " User=" & username & ";" _
               & " Password=" & password & ";" _
               & " OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384    'Option=3"

    Set banco = New ADODB.Connection
    'preparando o objeto connection

    banco.CursorLocation = adUseClient
    'usamos um cursor do lado do cliente pois os dados
    'serao acessados na maquina do cliente e nao de um servidor

    If banco.State = adStateOpen Then banco.Close
    banco.Open strConnect

    Exit Function

Err:
    Conectar = False
    MsgBox "Impossivel conectar ao servidor !!!!"
    End
End Function

Public Function sqlIncluir(ByVal sqlTabela As String, ByVal sqlCampo As String, ByVal sqlDados As String, Janela As Form, ByVal Mensagem As String) As Boolean
    On Error GoTo ErrEdit
    Dim Tabela As ADODB.Recordset
    Set Tabela = CreateObject("ADODB.Recordset")
    If Mensagem = "S" Then Aguarde_Process Janela, True
    Sql = "INSERT INTO " & sqlTabela & "(" & sqlCampo & ") values (" & sqlDados & ")"
    With Tabela
        .CursorType = adOpenStatic    'Este é o unico tipo de cursor a ser usado com um cursor localizado no lado do cliente
        .CursorLocation = adUseClient    'estamos usando o cursor no cliente
        .LockType = adLockPessimistic    'Isto garente que o registros que esta sendo editado pode ser salvo
        .Source = Sql    'altere para tabela que desejar a fonte de dados usamos uma instrucal SQL
        .ActiveConnection = banco  'O recordset precisa saber qual a conexao em uso
        .Open    'abre o recordset com isto o evento MoveComplete sera disparado
        If Mensagem = "S" Then Aguarde_Process Janela, False
        If Mensagem = "S" Then MsgBox ("Dados incluidos com sucesso.."), vbInformation
    End With
    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Function
ErrEdit:
    If Mensagem = "S" Then Aguarde_Process Janela, False

    MsgBox "Error ao incluir o registro", vbInformation
End Function

Public Function sqlEditar(ByVal sqlTabela As String, ByVal Sqlconsulta As String, Janela As Form, ByVal Mensagem As String) As Boolean
    Dim Tabela As ADODB.Recordset
    Set Tabela = CreateObject("ADODB.Recordset")
    On Error GoTo ErrEdit

    If Mensagem = "S" Then Aguarde_Process Janela, True

    Sql = "Select * from " & sqlTabela & " where " & Sqlconsulta
    With Tabela
        .CursorType = adOpenStatic    'Este é o unico tipo de cursor a ser usado com um cursor localizado no lado do cliente
        .CursorLocation = adUseClient    'estamos usando o cursor no cliente
        .LockType = adLockPessimistic    'Isto garente que o registros que esta sendo editado pode ser salvo
        .Source = Sql    'altere para tabela que desejar a fonte de dados usamos uma instrucal SQL
        .ActiveConnection = banco  'O recordset precisa saber qual a conexao em uso
        .Open    'abre o recordset com isto o evento MoveComplete sera disparado

        If Mensagem = "S" Then Aguarde_Process Janela, False

    End With

    banco.BeginTrans
    Tabela.Resync adAffectCurrent
    Tabela.Move 0

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Function
ErrEdit:
    If Mensagem = "S" Then Aguarde_Process Janela, False
    Select Case Err.Number
    Case -2147217885
        MsgBox "Registro não pode ser editado, pois foi excluído por outro usuário!", vbExclamation
        Tabela.CancelUpdate
        banco.RollbackTrans
    Case -2147467259
        MsgBox "Registro não pode ser editado, pois está sendo usado por outro usuário!", vbExclamation
        Tabela.CancelUpdate
        banco.RollbackTrans
    Case -2147217864
        MsgBox "Registro foi excluído ou alterado recentemente por outro usuário! Atualizando, aguarde...", vbExclamation
    End Select
End Function

Public Function sqlDeletar(ByVal sqlTabela As String, ByVal Sqlconsulta As String, Janela As Form, ByVal Mensagem As String) As Boolean
    On Error GoTo ErroSalvar
    Dim Tabela As ADODB.Recordset
    Set Tabela = CreateObject("ADODB.Recordset")

    If Mensagem = "S" Then Aguarde_Process Janela, True

    Sql = "Delete from " & sqlTabela & " where " & Sqlconsulta

    With Tabela
        .CursorType = adOpenStatic    'Este é o unico tipo de cursor a ser usado com um cursor localizado no lado do cliente
        .CursorLocation = adUseClient    'estamos usando o cursor no cliente
        .LockType = adLockPessimistic    'Isto garente que o registros que esta sendo editado pode ser salvo
        .Source = Sql    'altere para tabela que desejar a fonte de dados usamos uma instrucal SQL
        .ActiveConnection = banco  'O recordset precisa saber qual a conexao em uso
        .Open    'abre o recordset com isto o evento MoveComplete sera disparado
        If Mensagem = "S" Then Aguarde_Process Janela, False
        If Mensagem = "S" Then MsgBox ("Dados excluidos com sucesso..."), vbInformation
    End With

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing

    Exit Function
ErroSalvar:
    If Mensagem = "S" Then Aguarde_Process Janela, False
    Select Case Err.Number
    Case 0
        banco.CommitTrans
    Case -2147217864
        MsgBox "Este registro já foi excluída por outro usuário!", vbInformation, "Erro de Gravação"
        banco.RollbackTrans
    Case -2147467259
        MsgBox "Exclusão impossibilitada no momento. O registro encontra-se bloqueado por outro usuário." & vbCr & "Você pode cancelar a exclusão ou tentar excluir mais tarde...", vbExclamation, "Erro de gravacao"
        Tabela.CancelUpdate
        banco.RollbackTrans
    Case Else
        MsgBox ("Error ao Excluir o registro..."), vbExclamation

    End Select
End Function

Public Function sqlAlterar(ByVal sqlTabela As String, ByVal sqlDados As String, ByVal Sqlconsulta As String, Janela As Form, ByVal Mensagem As String) As Boolean
    On Error GoTo ErroSalvar
    Dim Tabela As ADODB.Recordset
    Set Tabela = CreateObject("ADODB.Recordset")

    If Mensagem = "S" Then Aguarde_Process Janela, True

    Sql = "UPDATE " & sqlTabela & " SET " & sqlDados & " WHERE " & Sqlconsulta
    With Tabela
        .CursorType = adOpenStatic    'Este é o unico tipo de cursor a ser usado com um cursor localizado no lado do cliente
        .CursorLocation = adUseClient    'estamos usando o cursor no cliente
        .LockType = adLockPessimistic    'Isto garente que o registros que esta sendo editado pode ser salvo
        .Source = Sql    'altere para tabela que desejar a fonte de dados usamos uma instrucal SQL
        .ActiveConnection = banco  'O recordset precisa saber qual a conexao em uso
        .Open    'abre o recordset com isto o evento MoveComplete sera disparado
        If Mensagem = "S" Then Aguarde_Process Janela, False
        If Mensagem = "S" Then MsgBox ("Dados alterados com sucesso.."), vbInformation
    End With

    If Tabela.State = 1 Then Tabela.Close
    Set Tabela = Nothing


    Exit Function
ErroSalvar:
    If Mensagem = "S" Then Aguarde_Process Janela, False
    Select Case Err.Number
    Case -2147217864
        confirma = MsgBox("Este registro foi alterado ou excluído por outro usuário. Deseja atualizar o registro?", vbQuestion + vbYesNo)
        If confirma = vbYes Then
            Tabela.CancelUpdate
            banco.RollbackTrans
        Else
            Tabela.Requery adAffectCurrent
        End If
    Case -2147467259
        MsgBox ("Este código já foi implementado por outro usuário!")
    Case Else
        MsgBox ("Error ao alterar o registro..."), vbExclamation

    End Select
End Function

Function Query(Sql As String) As Boolean
    On Error Resume Next
    With rs
        rs.Open Sql, Mysql_Connection
        mySQL.Last_Query = rs.GetString
        rs.Close
    End With
End Function

Function Fetch_Array(STR As String) As Boolean
    mySQL.Array_Query = Split(STR, vbTab)
End Function


Function CloseConnection()
'On Error Resume Next
    Mysql_Connection.Close
End Function

