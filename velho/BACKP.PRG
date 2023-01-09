*****************************************************************************
* Programa...: BACK1.PRG                                                    *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao:                *
* Objetivo...: BACKUP DOS ARQUIVOS EM DISCO                                 *
* Sistema....: SISPAN - Sistema de Panifica‡Æo                              *
*****************************************************************************

resp = space(01)

save screen to tela_1
set color to w/b
@ 23,05 say "<ESC> Para sair."
Janela2("W",.F.,10,10,20,63,.F.)
set color to w/r
@ 10,20 say "Back-Up dos Arquivos do Sistema"
SetColor("N/W,W+/R")
@ 11,11 say "Ser  iniciado o processo de c¢pia para os disquetes."
@ 12,11 say "Ap¢s o primeiro disquete ser completado ser  solici-"
@ 13,11 say "tado o segundo e assim por diante, at‚ que todos  os"
@ 14,11 say "arquivos do Sistema tenham sido copiados.           "
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
restore screen from tela_1


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
      * MAKE A BACKUP TO DRIVE A:
      *--------------------------------------------------------------------

      *-- Initialize progress bar variables
nBarLen  := 50
nBarRow  := 7
nBarCol  := 15
nLastCol := NIL
      *-- Draw a fancy box around the progress bar ...
DrawBar ( nBarRow, nBarCol, 'Criando backup no drive A:', lEscape )

      *-- Draw a box for displaying file names
cOldCol := SetColor ( 'W+/RB' )
Scroll   ( nBarRow+6, nBarCol-2, nBarRow+13, nBarCol+nBarLen+1 )
DispBox  ( nBarRow+6, nBarCol-2, nBarRow+13, nBarCol+nBarLen+1 )
DevPos   ( nBarRow+6, nBarCol+(nBarLen-10)/2 )
DevOut   ( 'Copiando... ' )
SetColor ( cOldCol )
nLastCol := NIL

      *--------------------------------------------------------------------
      * Syntax :
      * FUNCTION R_Backup ( aSkeleton, aExclude, cDestDrv, cFileName,
      *    aBar, lEscape, lSilent, bRBckBlk, aMsg )
      *--------------------------------------------------------------------
nRetCode := R_BackUp ( ;
{ '*.TXT', nil, '*.dbf', '*.dbt'},{ 'DUTCH.REG' }, ;
'A:', nil, ;
{ nBarLen, nBarRow, nBarCol, Chr (177), 'GR+/R' }, ;
lEscape, lSilent, ;
{ |fname| ShowName ( fname, nBarRow+4, nBarCol, nBarLen ) }, ;
{ MaxRow () - 1, nil, nil, 'W+*/B' } )

      *--------------------------------------------------------------------
      * You may also use the (more readable) COMMAND FORM syntax
      * (as included in RBckLib.Ch) :
      *
      * RBACKUP FILES { '*.TXT', '*.NG', '*.PRG', '*.DOC', '*.REG' } ;
      *        EXCLUDE { 'DUTCH.REG' } ;
      *        TO A: ;
      *        BAR { nBarLen, nBarRow, nBarCol, Chr(177), 'GR+/R' } ;
      *        ESCAPE ;
      *        BLOCK { |fname| ShowName ( fname, nBarRow+4, nBarCol, nBarLen ) } ;
      *        MESSAGE { MaxRow () - 1, nil, nil, 'W+*/B' } ;
      *        RETCODE nRetCode
      *--------------------------------------------------------------------

      *-- Clear screen
* Scroll ( 2, 0 )

IF nRetCode != BU_OKAY
                                                                                                  *-- Oops, error detected : display error message !

    Alert ( 'BackUp Error : ' + aErrTxt [ nRetCode ] )

ELSE
                                                                                                  *-- No errors detected !

    Alert ( 'Backup Concluido !' )

ENDIF


restore scree from tela_1

inkey(0)
close all
return


*!*****************************************************************************
*!
*!       Function: DRAWBAR()
*!
*!*****************************************************************************
FUNCTION DrawBar ( nRow, nCol, cHeader, lEscape )
*==========================================================================
* FUNCTION TO DRAW A FANCY BOX AROUND THE PROGRESS BAR
*==========================================================================

LOCAL	cOldCol := SetColor ( 'W+/BG' )	&& Save & set screen color
LOCAL	i				&& Counter

*-- Decrease row/column
nRow -= 2
nCol -= 2

*-- Clear box area
Scroll  ( nRow, nCol, nRow+4, nCol+53 )

*-- Draw box
DispBox ( nRow, nCol, nRow+4, nCol+53 )

*-- Draw header (centered)
DevPos ( nRow, ( MaxCol() + 1 - Len ( cHeader ) ) / 2 )
DevOut ( ' ' + cHeader + ' ' )

*-- Draw footer
IF lEscape
    DevPos ( nRow+4, 31 )
    DevOut ( 'ESC para Sair ' )
ENDIF

*-- Draw ruler lines
DevPos ( nRow+1, nCol+2 )
*-- Draw percentages
DevOut ( '0  10   20   30   40   50   60   70   80   90  100' )
DevPos ( nRow+3, nCol+2 )
*-- Draw marks
DevOut ( Chr ( 16 ) )
FOR i := 1 TO 9
    DevPos ( nRow+3, nCol+1+5*i )
    DevOut ( Chr ( 30 ) )
NEXT
DevPos ( nRow+3, nCol+51 )
DevOut ( Chr ( 17 ) )

*-- Restore old color
SetColor ( cOldCol )

RETURN nil


*!*****************************************************************************
*!
*!       Function: SHOWNAME()
*!
*!*****************************************************************************
FUNCTION ShowName ( cFName, nRow, nCol, nBarLen )
*==========================================================================
* FUNCTION TO DISPLAY THE CURRENT FILE NAME
*==========================================================================

LOCAL	cOldCol := SetColor ( 'W+/RB' )	&& Save & set screen color

*-- Clear line
IF nLastCol != NIL

                                                                                        *-- Not the first file
    DevPos ( nRow+8, nLastCol+1 )
    DevOut ( '[ok]' )

ENDIF

nLastCol := nCol + Len ( cFName )

Scroll ( nRow+3, nCol, nRow+8, nCol+nBarLen, 1 )

DevPos ( nRow+8, nCol )
DevOut ( cFName + ' ...' )

SetColor ( cOldCol )

RETURN BU_CONT                          && Continue backup/restore function


*!*****************************************************************************
*!
*!       Function: CONFIRM()
*!
*!*****************************************************************************
FUNCTION Confirm ( cFileSpec )
*==========================================================================
* CONFIRMATION FOR OVERWRITING AN EXISTING FILE (WHILE RESTORING)
*==========================================================================

LOCAL	nChoice			&& Choice made from ALERT

IF M->lConfirm
                                                                                        *-- Confirmation is ON

    nChoice := Alert ( 'File ' + cFileSpec + ' exists !', ;
    { 'Skip file', 'Replace file', 'Overwrite ALL', 'Abort' } )

    DO CASE
        CASE nChoice = 1

            RETURN BU_SKIPFILE

        CASE nChoice = 2

            RETURN BU_OVERWRITE

        CASE nChoice = 3

            M->lConfirm := .F.
            RETURN BU_OVERWRITE

        CASE nChoice = 4

            RETURN BU_ABORT
    ENDCASE

ENDIF

RETURN BU_OVERWRITE			&& Confirmation = OFF
*
* EOF RBckDemo.PRG ========================================================





*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
