*****************************************************************************
* Programa...: REST.PRG                                                     *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......:                   Data da Ultima Atualizacao:                *
* Objetivo...: RESTAURA DADOS BACKUP FEITO EM DISCO                         *
* Sistema....: SISPAN - Sistema de Panifica‡Æo                              *
*****************************************************************************

do sinal with "RESTAURA   ARQUIVOS "

resp = "N"
save screen to tela_1
set color to w/b
@ 23,05 say "<ESC> Para sair."
Janela2("W",.F.,10,10,20,63,.F.)
set color to w/r
@ 10,20 say "Recupera‡„o dos Arquivos do Sistema"
SetColor("N/W,W+/R")
@ 11,11 say "Ser  iniciado o processo de c¢pia para o winchester."
@ 12,11 say "Ap¢s o primeiro disquete ser copiado ser  solicitado"
@ 13,11 say "o  segundo  e  assim  por  diante, at‚ que todos os "
@ 14,11 say "arquivos do Sistema tenham sido recuperados.        "
@ 16,11 say "Nao interrompa de forma alguma este Processo !"
@ 17,11 say "Certifique-se de colocar os Disquetes corretos !"
@ 19,11 say "Deseja iniciar o Processo ?"
@ 19,40 get resp pict "!"
read
if resp = "N"
    restore scree from tela_1
    close all
    return
endif
restore scree from tela_1


*--------------------------------------------------------------------------
* Standard Clipper HEADER files
*--------------------------------------------------------------------------
#include "SetCurs.CH"

*--------------------------------------------------------------------------
* RBCKLIB header file
*--------------------------------------------------------------------------
#include "RBCKLib.CH"

*--------------------------------------------------------------------------
* Static variables
*--------------------------------------------------------------------------
*-- Initialize the array with error messages (from RBckLib.CH)
aErrTxt := BU_ERRMSG

*!*****************************************************************************
*!
*!       Function: RBCKDEMO()
*!
*!*****************************************************************************

do sinal with "BACKUP DOS ARQUIVOS "


nChoice := 1                    && Menu choice
lEscape := .T.                  && Backup/Restore interruptable
lSilent := .F.                  && Use the beeper ...

*-- Enable blinking colors
SetBlink ( .T. )

*-- Set color to BRIGHT WHITE on BLUE
SetColor ( 'W+/B' )

*-- Disable scoreboard
Set ( _SET_SCOREBOARD, .f. )

*-- Set cursor OFF
SetCursor ( SC_NONE )



      *--------------------------------------------------------------------
      * RESTORE FROM DRIVE A:
      *--------------------------------------------------------------------

      *-- Overwrite confirmation initial to TRUE
M->lConfirm := .T.

      *-- Initialize the progress bar variables
nBarLen     := 50
nBarRow     := 7
nBarCol     := 15

      *-- Draw a fancy box around the progress bar ...
DrawBar ( nBarRow, nBarCol, 'Restaurando Arquivos no drive A:', lEscape )

      *-- Draw a box for displaying file names
cOldCol := SetColor ( 'W+/RB' )
Scroll   ( nBarRow+6, nBarCol-2, nBarRow+13, nBarCol+nBarLen+1 )
DispBox  ( nBarRow+6, nBarCol-2, nBarRow+13, nBarCol+nBarLen+1 )
DevPos   ( nBarRow+6, nBarCol+(nBarLen-14)/2 )
DevOut   ( 'Restaurando...' )
SetColor ( cOldCol )
nLastCol := NIL

      *---------------------------------------------------------------------
      * Syntax :
      * FUNCTION R_Restore ( cSrcDrv, aMask, cDestSpec, cFileName,
      *    aBar, lEscape, lSilent, bRBckBlk, aMsg, bConfirm )
      *---------------------------------------------------------------------
nRetCode := R_Restore ( 'A:', { '*.*' }, zdiretorio, nil, ;
{ nBarLen, nBarRow, nBarCol, Chr (177), 'GR+/R' }, ;
lEscape, lSilent, ;
{ |fname| ShowName ( fname, nBarRow+4, nBarCol, nBarLen ) }, ;
{ MaxRow () - 1 }, ;
{ |fname| Confirm ( fname ) } )

      *--------------------------------------------------------------------
      * You may also use the (more readable) COMMAND FORM syntax
      * (as included in RBckLib.Ch) :
      *
      * RRESTORE FROM A: TO C:\RBCKTEST ;
      *    BAR { nBarLen, nBarRow, nBarCol, Chr(177), 'GR+/R' } ;
      *    ESCAPE ;
      *    BLOCK { |fname| ShowName ( fname, nBarRow+4, nBarCol, nBarLen ) } ;
      *    MESSAGE { MaxRow () - 1, nil, nil, nil } ;
      *    CONFIRM { |fname| Confirm ( fname ) } ;
      *    RETCODE nRetCode
      *--------------------------------------------------------------------

      *-- Clear screen
      * Scroll ( 2, 0 )

IF nRetCode != BU_OKAY
                                                                                          *-- Oops, error detected : display error message !
    Alert ( 'Restauracao com Erro : ' + aErrTxt [ nRetCode ] )
ELSE
                                                                                          *-- No errors detected !
    Alert ( 'Arquivos Restaurados !' )
ENDIF

close all
return
*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
