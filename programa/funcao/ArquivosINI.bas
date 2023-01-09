Attribute VB_Name = "ArquivosINI"
' API usada para ler os arquivos INI . Geralmente voc� faz esta declara��o em um m�dulo:
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                                         (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal _
                                                                                                                                lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

' API usada para escrever em uma arquivo INI. Geralmente voc� faz esta declara��o em um m�dulo
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                                           (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal _
                                                                                                                              lpFileName As String) As Long


' Fun��o - ReadINI - l� um arquivo INI. Precisa de tr�s par�metros :
' O nome da Se��o
' o nome da Entrada
' o nome do Arquivo INI.

Public Function ReadINI(Secao As String, Entrada As String, Arquivo As String)
'Arquivo=nome do arquivo ini
'Secao=O que esta entre []
'Entrada=nome do que se encontra antes do sinal de igual
    Dim retlen As String
    Dim Ret As String
    Ret = String$(255, 0)
    retlen = GetPrivateProfileString(Secao, Entrada, "", Ret, Len(Ret), Arquivo)
    Ret = Left$(Ret, retlen)
    ReadINI = Ret
End Function

' A fun��o - WriteINI - escreve em um arquivo INI.
' Precisa de quatro par�metros :
' o nome da Se��o
' o nome da Entrada
' o nome do Texto ( Valor )
' o nome do arquivo INI.

Public Sub WriteINI(Secao As String, Entrada As String, texto As String, Arquivo As String)
'Arquivo=nome do arquivo ini
'Secao=O que esta entre []
'Entrada=nome do que se encontra antes do sinal de igual
'texto= valor que vem depois do igual
    WritePrivateProfileString Secao, Entrada, texto, Arquivo
End Sub


' No nosso caso para Ler os valores do arquivo CONSTRUTORA.INI usamos o seguinte c�digo:

' valortempo = ReadINI("Geral", "Tempo", App.Path & "\CONSTRUTORA.ini")
' valorajuda = ReadINI("Geral", "Ajuda", App.Path & "\CONSTRUTORA.ini")
' atualizaperguntas = ReadINI("Geral", "Atualiza", App.Path & "\CONSTRUTORA.ini")
'
' As vari�veis valortempo, valorajuda e atualizaperguntas ir�o armazenar os valores
' lidos do arquivo Show.ini atrav�s da fun��o ReadINI.
'
' Para Escrever em um arquivo INI alterando os valores das entradas:
' Tempo, Ajuda e Atualiza , usamos o seguinte c�digo:'
'
' Call WriteINI("Geral", "Tempo", txttempo.text, App.Path & "\show.ini")
' Call WriteINI("Geral", "Ajuda", txtajuda.text, App.Path & "\show.ini")
' Call WriteINI("Geral", "Atualiza", txtatualiza.text, App.Path & "\show.ini")

