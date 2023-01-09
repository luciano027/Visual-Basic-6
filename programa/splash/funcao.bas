Attribute VB_Name = "funcao"
Option Explicit

'Adiciona Pausa no programa
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Declare Function Arc Lib "gdi32" (ByVal HDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long

Declare Function RoundRect Lib "gdi32" (ByVal HDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2
Public sResultadoDaConsulta As String

Public contapagina As Integer


'Encontra o Handle da janela
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
                                   (ByVal lpClassName As String, _
                                    ByVal lpWindowName As String) As Long

'Envia mensagem para a janela
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
                                    (ByVal hWnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     ByVal lParam As Long) As Long

'Executa uma aplicação
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                                     (ByVal hWnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long


'-------------------- variaveis do grafico ----------------------
Public Titulo1 As String
Public Titulo2 As String
Public NBarras As Integer
Public TituloB(10) As String
Public ValorB(10) As Double
Public LegendaB(10) As String
'----------------------------------------------------------------
Public Const WM_CLOSE = &H10
Public Const SW_SHOW = 1

' API usada para ler os arquivos INI . Geralmente você faz esta declaração em um módulo:
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                                         (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal _
                                                                                                                                lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

' API usada para escrever em uma arquivo INI. Geralmente você faz esta declaração em um módulo
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                                           (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal _
                                                                                                                              lpFileName As String) As Long



Public Sub Exibe_Erros(strMensagem_Erro As String)
    On Error Resume Next
    MsgBox strMensagem_Erro, vbInformation, "Ocorreu um erro !"
End Sub

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
