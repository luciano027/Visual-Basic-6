*****************************************************************************
* Programa...: ARQPROC.PRG                                                  *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 05/06/95          Data da Ultima Atualizacao: 05/06/95       *
* Objetivo...: Arquivos de Procedures e Funcoes                             *
* Sistema....: SISVEN (Sistema Controle Estoque/Vendas)                     *
*****************************************************************************


*--------------------------------------------------------------------------
* Objetivo..: Apresentar sinais nos quadrinhos de orientacao `a direita no
*             topo da tela, indicando para o usuario o modulo  do  sistema
*             que esta sendo utilizado.
* Parametros: s1: Primeiro Sinal e s2: Segundo Sinal
*--------------------------------------------------------------------------
procedure sinal
parameters s1
s1 = space((14 - len(substr(s1,1,14)))/2) + s1
 * Centralizar o sinal dentro dos 12 espacos disponiveis no quadrinho
s1 = s1 + space(14 - len(s1))
set color to r/w
@ 24,66 say s1 pict "@s14"
set color to
return


*--------------------------------------------------------------------------
* Objetivo..: Apresentar, por um determinado tempo, uma mensagem de alerta
*             (piscante e em video reverso) na linha 23 da tela.
* Parametros: O texto da mensagem, tempo que a mesma devera' ser apresen-
*             tada na tela e um sinal sonoro de advertencia.
*--------------------------------------------------------------------------
procedure mensagem
parameters mensagem, tempo , som
if pcount()<3
    som=2
endif
private    vtab
vtab = (80-len(mensagem))/2
   * Centraliza a mensagem em uma cadeia de 80 caracteres
set color to w/b
@ 23,02 say space(78)
beep(som)  && A Funcao Beep() faz soar um sinal sonoro de advertencia.
set color to *gr+/b
@ 23,vtab say mensagem
set color to w/b
inkey(2)
@ 23,02 say space(78)
return

*--------------------------------------------------------------------------
* Sintaxe...: Beep()
* Objetivo..: Soar um sinal para chamar a atencao do usuario
* Parametros: som, indicando o tipo de som a soar
* Retorna...: Um sinal sonoro para o usuario.
*--------------------------------------------------------------------------
function beep
parameters som
set console on
if som=1         && Finalizacao
    tone(87.3,2)
    tone(40,7)
elseif som=2     && Erro
    tone(261.7,2)
    tone(392,7)
elseif som=3     && Than
    tone(300,4)
    tone(300,4)
    tone(300,4)
    tone(300,7)
endif
return(.t.)
set console off

*--------------------------------------------------------------------------
* Objetivo..: Apresentar um aviso centralizado em uma linha da tela.
* Parametros: O texto do aviso e a linha onde devera' ser apresentado
*--------------------------------------------------------------------------
procedure aviso
parameters linha,aviso
private    vtab
vtab = (80-len(aviso))/2
   * Centraliza o texto do aviso na tela
@ linha,vtab say aviso
return

function quadro(ls,cs,li,ci)
cor := setcolor()
col := savescreen(ls+1,ci+1,li+1,ci+2)
lit := savescreen(li+1,cs+1,li+1,ci+1)
for i = 2 to len(col) step 2
    col = stuff(col,i,1,chr(8))
next
for i = 2 to len(lit) step 2
    lit = stuff(lit,i,1,chr(8))
next
restscreen(ls+1,ci+1,li+1,ci+2,col)
restscreen(li+1,cs+1,li+1,ci+1,lit)
setcolor(cor)
@ ls,cs clear to li,ci
@ ls,cs to li,ci
latgura := ci-cs-1
return(.t.)

*--------------------------------------------------------------------------
procedure limpa
set color to  w/b
@ 23,02 say space(75)
return

*--------------------------------------------------------------------------
procedure inicia
a:=fcount()
for i = 1 to a
    if type(field(i))<>"M"
        nomecampo:=field(i)
        nomevar:="w"+nomecampo
        public &nomevar
        do case
            case type(field(i)) ="C"
                &nomevar:=space(len(&nomecampo))
            case type(field(i)) = "N"
                &nomevar:=0
            case type(field(i)) = "D"
                &nomevar:=ctod("")
            case type(field(i)) = "L"
                &nomevar:=.f.
        endcase
    endif
next i
return

*--------------------------------------------------------------------------
procedure revar
a:=fcount()
for i = 1 to a
    if type(field(i)) <> "M"
        nomecampo:=field(i)
        nomevar:="w"+nomecampo
        if !empty(nomevar)
            if type("&nomevar") <>"U"
                replace &nomecampo with &nomevar
            endif
        endif
    endif
next i
return

*--------------------------------------------------------------------------
procedure libera
a:=fcount()
for i = 1 to a
    if type(field(i)) <> "M"
        nomevar := "w"+field(i)
        release &nomevar
    endif
next i
return

*--------------------------------------------------------------------------
procedure iguala
a:=fcount()
for i = 1 to a
    if type(field(i))<> "M"
        nomecampo=fieldname(i)
        nomevar="w"+nomecampo
        public &nomevar
        &nomevar:=&nomecampo
    endif
next i
return

*--------------------------------------------------------------------------
procedure impri
set color to gr+/b
@ 23,03 say "Ajuste a Impressora e Tecle <Enter> quando PRONTO ou <Esc> para CANCELAR !"
do while .t.
    tk=inkey(0)
    if tk <> 13
        do mensagem with "Impress„o Cancelada !",3
        return .f.
    endif
    if .not. isprinter()
        limpa()
        set color to gr+/b
        @ 23,03 say "A Impressora N„o Est  Ativa, Verifique por favor !"
        loop
    endif
    exit
enddo
limpa()
return .t.

*--------------------------------------------------------------------------
function rolatela(nli,nci,nlf,ncf)
local ncontador
for ncontador := 1 to abs(nlf - nli + 1)
    scroll(nli,nci,nlf,ncf,1)
    inkey(0.03)
next
return nil

*--------------------------------------------------------------------------
procedure inicio
set color to n/w
@ 01,00 say "            º Sistema de Vendas e Controle de Estoque  º           º            "
set color to gr/w+
@ 01,00 say "PC-Sistemas"
set color to n/w
@ 24,00 say "F2-CalculadoraºF3-VendaºF4-ConsultaºF5-Orcam.ºF6-FichaºF7-Agenaº               "
@ 01,70 say date() pict "@d"
@ 01,58 say showtime() pict "99:99:99"
set color to  w/b
@ 02,00 say "                                                                                "
@ 04,01 say "                                                                              "
@ 05,01 say "                                                                              "
@ 06,01 say "      ÜÜÜÜÜÜÜÜ  ÜÜÜÜÜ  ÜÜÜÜÜÜÜÜ  ÜÜÜÜ    ÜÜÜÜ  ÜÜÜÜÜÜÜÜÜÜ  ÜÜÜÜ     ÜÜÜ       "
@ 07,01 say "     ²²²²²²²²Û ²²²²²Û ²²²²²²²²Û ²²²²Û   ²²²²Û ²²²²²²²²²²Û ²²²²Û    ²²²Û       "
@ 08,01 say "     ²²²²²²²²  ²²²²²Û ²²²²²²²²  ²²²²Û   ²²²²Û ²²²²²²²²²²  ²²²²²Û   ²²²Û       "
@ 09,01 say "     ²²²Û       ²²²Û  ²²²Û      ²²²²Û   ²²²²Û ²²²Û        ²²²²²²Û  ²²²Û       "
@ 10,01 say "     ²²²ÛÜÜÜÜÜ  ²²²Û  ²²²ÛÜÜÜÜÜ ²²²²Û   ²²²²Û ²²²ÛÜÜÜÜÜÜÜ ²²²²²²²Û ²²²Û       "
@ 11,01 say "     ²²²²²²²²Û  ²²²Û  ²²²²²²²²Û ²²²²Û   ²²²²Û ²²²²²²²²²²Û ²²²²²²²²Û²²²Û       "
@ 12,01 say "     ²²²²²²²²Û  ²²²Û  ²²²²²²²²Û ²²²²Û   ²²²²Û ²²²²²²²²²²  ²²²Û ²²²²²²²Û       "
@ 13,01 say "         ²²²²Û  ²²²Û      ²²²²Û ²²²²Û   ²²²²Û ²²²Û        ²²²Û  ²²²²²²Û       "
@ 14,01 say "      ÜÜÜ²²²²Û  ²²²ÛÜ  ÜÜÜ²²²²Û ²²²²²Û ²²²²²Û ²²²ÛÜÜÜÜÜÜÜ ²²²Û   ²²²²²Û       "
@ 15,01 say "     ²²²²²²²²Û ²²²²²Û ²²²²²²²²Û  ²²²²²²²²²²Û  ²²²²²²²²²²Û ²²²Û    ²²²²Û       "
@ 16,01 say "     ²²²²²²²²  ²²²²²  ²²²²²²²²    ²²²²²²²²    ²²²²²²²²²²  ²²²      ²²²        "
@ 17,01 say "                                                                              "
@ 18,01 say "                 Sistema de Vendas e Controle de Estoque - V1.0               "
@ 19,01 say "                                                                              "
@ 21,00 say "                                                                                "
set color to   w/b
@ 22,00 say "ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ Mensagem ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿"
@ 23,00 say "                                                                                "
select 1
use empresa
locate for !empty(empresa)
if found()
    wempresa  = empresa
    ep = space((80 - len(alltrim(substr(wempresa,1,80))))/2) + wempresa
    set color to gr+/b
    @ 19,01 say ep pict "@!"
endif
close all
inkey(.5)
limpa()
return

*--------------------------------------------------------------------------
function memoed(varia,titulo,lt,ct,lb,cb)
save screen to tela
set color to gr+/b
@ 23,03 say "[Ctrl]+[W] Finaliza Edicao e Salva. [ESC] Aborta Edicao."
set color to /w
varia = memoedit(varia,lt,ct,lb,cb)
restore screen from tela
return(varia)

*--------------------------------------------------------------------------
procedure atualiza
set color to w/b*
@ 23,03 say "Aguarde.. Atualizando Arquivos."
close all
if !file("????????.dbf")
    do mensagem with "Diretorio invalido...",8
    return
endif
set excl on
! del ind*.*
! del *.ntx
dir_dbf  = directory("*.dbf")
dir      = len(dir_dbf)
for y = 1 to dir
    arquivo = dir_dbf[y,1]
    use &arquivo
    pack
next
close all
do mensagem with "Fim da atualizacao...",8
set excl off
return

*--------------------------------------------------------------------------
*********************************************
* Funcao que imprime numero por extenso     *
* Luis Aguiar                               *
*********************************************
do while .t.
    kvalor=0
    cls
    @ 1,1 to 24,79 double
    @ 8,10 say'ENTRE COM O VALOR' GET kvalor pict '999,999,999.99'
    READ
    if lastkey()=27
        cls
        quit
    endif
    convext(kvalor,10,2)
    @ 23,10 say 'Enter para continuar...'
    inkey(0)
enddo
*************************************
FUNCTION CONVEXT
PARAMETERS kvalor,lin,col
public kvalor,lin,col
EXT = "("
A1="Cem"
A2="Duzentos"
A3="Trezentos"
A4="Quatrocentos"
A5="Quinhentos"
A6="Seiscentos"
A7="Setecentos"
A8="Oitocentos"
A9="Novecentos"
B1="Um"
B2="Dois"
B3="Tres"
B4="Quatro"
B5="Cinco"
B6="Seis"
B7="Sete"
B8="Oito"
B9="Nove"
B10="Dez"
B11="Onze"
B12="Doze"
B13="Treze"
B14="Quatorze"
B15="Quinze"
B16="Dezesseis"
B17="Dezessete"
B18="Dezoito"
B19="Dezenove"
C2="Vinte"
C3="Trinta"
C4="Quarenta"
C5="Cinquenta"
C6="Sessenta"
C7="Setenta"
C8="Oitenta"
C9="Noventa"
P=1
STG=0
V1=SUBSTR(STR(kvalor,12,2),1,3)
V2=SUBSTR(STR(kvalor,12,2),4,3)
V3=SUBSTR(STR(kvalor,12,2),7,3)
V4=" "+SUBSTR(STR(kvalor,12,2),11,2)
DO WHILE P<5
    ST1=0
    ST2=0
    ST3=0
    PA="V"+STR(P,1,0)
    VA=&PA
    CT=1
    DO WHILE VAL(SUBSTR(VA,1,1))#0
        IF VAL(SUBSTR(VA,1,1))=CT
            IF STG=1
                IF P#4
                    CONEC=" e "
                ENDIF
            ENDIF
            EXT1="A"+STR(CT,1,0)
            EXT1=&EXT1
            IF VAL(SUBSTR(VA,1,1))=1 .AND. VAL(SUBSTR(VA,2,2))#0
                EXT1="Cento"
            ENDIF
            IF STG=0
                EXT=EXT+EXT1
            ELSE
                EXT=EXT+CONEC+EXT1
            ENDIF
            ST1=1
            EXIT
        ENDIF
        CT=CT+1
        IF CT=10
            EXIT
        ENDIF
    ENDDO
    CT=2
    DO WHILE VAL(SUBSTR(VA,2,1))#0 .AND. VAL(SUBSTR(VA,2,1))#1
        IF VAL(SUBSTR(VA,2,1))=CT
            EXT1="C"+STR(CT,1,0)
            IF ST1=0 .AND. STG=0
                EXT=EXT+&EXT1
            ELSE
                EXT=EXT+" e "+&EXT1
            ENDIF
            ST2=1
            EXIT
        ENDIF
        CT=CT+1
        IF CT=10
            EXIT
        ENDIF
    ENDDO
    CT=10
    DO WHILE VAL(SUBSTR(VA,2,1))=1
        IF VAL(SUBSTR(VA,2,2))=CT
            EXT1="B"+LTRIM(STR(CT,2,0))
            IF ST1=0 .AND. ST2=0 .AND. STG=0
                EXT=EXT + &ext1
            ELSE
                EXT=EXT+" e " + &EXT1
            ENDIF
            ST3=1
            EXIT
        ENDIF
        CT=CT+1
        IF CT=20
            EXIT
        ENDIF
    ENDDO
    CT=1
    DO WHILE VAL(SUBSTR(VA,2,1))=0 .OR. VAL(SUBSTR(VA,2,1))#1
        IF VAL(SUBSTR(VA,3,1))=CT
            EXT1="B" + LTRIM(STR(CT,2,0))
            IF ST1=0 .AND. ST2=0 .AND. STG=0
                EXT=EXT+&EXT1
            ELSE
                EXT=EXT+" e " +&EXT1
            ENDIF
            ST3=1
            EXIT
        ENDIF
        CT=CT+1
        IF CT=10
            EXIT
        ENDIF
    ENDDO
    CT=1
    IF ST1=1 .OR. ST2=1 .OR. ST3=1 .OR. (P=3 .AND. INT(kvalor)#0)
        IF P=1
            IF VAL(VA)=1
                EXT=EXT+" Milhao"
            ELSE
                EXT=EXT+" Milhoes"
            ENDIF
        ENDIF
        IF P=2
            EXT=EXT+" Mil"
        ENDIF
        IF P=3
            IF INT(kvalor)=1
                EXT=EXT+" Real"
            ELSE
                EXT=EXT+" Reais"
            ENDIF
        ENDIF
        IF P=4
            IF VAL(VA)=1
                EXT=EXT+" Centavo"
            else
                EXT=EXT+" Centavos"
            ENDIF
        ENDIF
        STG=1
    ENDIF
    P=P+1
ENDDO
EXT=EXT+")"
@ lin,col say ext
RETURN //(EXT)


*--------------------------------------------------------------------------
procedure verfil
para tfile
/***
*  GenericViewer( <cFile> ) --> cFile
*  View the contents of a text file (?)
*
*/
#define gv_blocksize    70000
#include "Memoedit.ch"
verfile(tfile)

static function verfile( cfile )
local cbuffer := "", nhandle := 0, nbytes := 0

cbuffer := space( gv_blocksize )
nhandle := fopen( cfile )

if ferror() != 0
    cbuffer := "Error Leitura do Arquivo!"
else
    nbytes = fread( nhandle, @cbuffer, gv_blocksize )
endif
fclose( nhandle )

cbuffer := rtrim( cbuffer )

set color to n/w
dt = quadro(03,00,20,78)
set color to r/w
@ 20,15 say "[ Use "+chr(27)+chr(18)+chr(26)+" para movimentar o cursor.  (Esc para Exit) ]"
set color to n/w
memoedit( cbuffer, 4, 01, 20 - 1, 78 - 1, .f., "MemoUDF" , 300 )


do while .t.
    save screen to tela_an
    set color to w/b
    @ 23,05 say "<ESC> Para sair."
    SetColor("W/W")
    Janela2("W",.F.,02,65,05,75,.F.)
    SetColor("N/W,W+/R")
    @ 03,66   prompt "RETORNO  "
    @ 04,66   prompt "IMPRIMIR "
    menu to xop
    restore screen from tela_an
    do case
        case xop = 1
            exit
        case xop = 2
            if !impri()
                return
            endif
            set color to w+/b
            dt = quadro(15,18,17,35)
            set color to *+w/b
            @ 16,20 say "Imprimindo..."
            copy file arq.txt to prn
            exit
        other
            loop
    endcase
enddo

return( cfile )

#undef gv_blocksize

/***
*  MemoUDF( <nMode>, <nLine>, <nColumn> ) -->
*
*
*/
function MemoUDF( nmode, nline, ncolumn )
return( ME_DEFAULT )


*--------------------------------------------------------------------------
procedure senha1
x = 0
do while .t.
    save screen to tela_se
    set color to n/w
    dt = quadro(10,22,12,40)
    set color to /w
    @ 11,23 say "Senha..:"
    set console off
    @ 11,33 say ""
    set color to /w
    wait to d1
    tone(100,1)
    @ row(),col() say "*"
    wait to d2
    tone(200,1)
    @ row(),col() say "*"
    wait to d3
    tone(300,1)
    @ row(),col() say "*"
    wait to d4
    tone(400,1)
    @ row(),col() say "*"
    wait to d5
    tone(500,1)
    @ row(),col() say "*"
    wait to d6
    tone(600,1)
    @ row(),col() say "*"
    set console on
    senha = upper(alltrim(d1+d2+"/"+d3+d4+"/"+d5+d6))
    if ctod(senha) # date()+1
        do mensagem with "Senha Nao Confere..."
        x++
        restore screen from tela_se
        if x = 3
            do mensagem with "Nao Autorizado para esta Consulta...",8
            inkey(.5)
            return .f.
        endif
        loop
    endif
    restore screen from tela_se
    exit
enddo
return .t.

*--------------------------------------------------------------------------
*---Cria uma janela com borda pela metade de cor diferente
Procedure Janela2()
Parameters Cor,But,X,Y,XF,YF,CB
Corant := SetColor()
Cora := "N/" +Cor
Corb := "W+/" +Cor
SetColor(Cora)
quadro(x,y,xf,yf)
Scroll(X,Y,XF,YF)
SetPos(X,Y)
Dispbox(X,Y,XF,YF,1)
SetColor(Corb)
IF (But)
    DevPos(X + 1, Y - 1)
    Devout(CHR(07))
    DevPos(X + 1, YF - 1)
    Devout(CHR(07))
    DevPos(XF - 1, Y + 1)
    Devout(CHR(07))
    DevPos(XF - 1, YF - 1)
    Devout(CHR(07))
EndIF
IF (CB)
    DevPos(X,Y)
    Devout("Ú")
    For C:= Y + 1 to YF - 1
        DevPos(X,C)
        Devout("Ä")
    Next
    For C:= X + 1 to XF - 1
        DevPos(C,Y)
        Devout("³")
    Next
    DevPos(XF,Y)
    Devout("À")
Else
    DevPos(X,YF)
    Devout("¿")
    For C:= X + 1 to XF - 1
        DevPos(C,YF)
        Devout("³")
    Next
    For C:= Y + 1 to YF - 1
        DevPos(XF,C)
        Devout("Ä")
    Next
    DevPos(XF,YF)
    Devout("Ù")
EndIF
SetColor("Corant")
Return(.T.)

*--------------------------------------------------------------------------
*ÄÄÄÄÄÄÄÄ Gera uma tela para senha acesso ao sistema.
Function Abertura()
Parameter NULO
Tela_Senha = SAVESCREEN(04,00,22,79)
SetColor("+W/N")
quadro(05,03,18,75)
Janela2("W",.F.,05,03,18,75,.F.)
SetColor("N/W")
@ 07,19 Say "  Sistema de Controle de Estoque e Vendas "
@ 08,19 Say "     Programa    distribuido     por      "
@ 09,19 Say "         Luciano Moreira Araujo           "
@ 10,19 Say "                                          "
SetColor("N/W")
@ 11,19 Say "          P r o g r a m a d o r           "
SetColor("N/W")
@ 13,19 Say "     Telefone: 3234-3189   /  9949-8502   "
SetColor("R/W")
@ 14,19 Say "ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß"
SetColor("R/W")
@ 15,04 SAY "þþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþþ"
set color to gr/w+
@ 17,63 say "PC-Sistemas"
SetColor("R/W")
set curso off
inkey(0)
set curso on
RESTSCREEN(04,00,22,79,Tela_Senha)


*--------------------------------------------------------------------------
Function Confirma
Parameter TEXTO
Private LES,CES,LARGJAN,TELAVELHA
LARGJAN=LEN(TEXTO)+2
IF LARGJAN < 16
    LARGJAN = 16
EndIF
CES = (80-LARGJAN)/2
LES = 19
TELAVELHA = SAVESCREEN(LES,0,LES+4,79)
Corantig  = SetColor()
quadro(LES,CES,LES+3,CES+LARGJAN)
Janela2("W",.F.,LES,CES,LES+3,CES+LARGJAN,.F.)
SetColor("N/W")
@ LES+1,CES+2 Say TEXTO
COLUNA = (LARGJAN-9)/2
SetColor("N/W,W+/W")
@ LES+2,CES+COLUNA   Prompt " Sim "
@ LES+2,CES+COLUNA+6 Prompt " N„o "
MENU To CONFIRM
Set Color To &CORANTIG
RESTSCREEN(LES,0,LES+4,79,TELAVELHA)
Return (CONFIRM)

*--------------------------------------------------------------------------
procedure testa_arq
if !file("c:\dos\arq.dos") .and. !file("c:\windows\arq.wim")
    waiton("Copia nao autorizada. Favor Contactar com o Programador...")
    inkey(1)
    close all
    quit
endif
save screen to tela01
waiton("Organizando Arquivos...")

wch = 1

select &wch
use arquivos
do while !eof()
    iguala()
    wch = wch + 1
    select &wch
    use &warq_dbf
    if !empty(wchave01)
        index on &wchave01 to &warq_i01
    endif
    if !empty(wchave02)
        index on &wchave02 to &warq_i02
    endif
    if !empty(wchave03)
        index on &wchave03 to &warq_i03
    endif
    use
    select 1
    skip
enddo
close all
restore screen from tela01
return

                    * Função Testa_Drive()
                    * Desenvolvida por: Edson Melo
                    * Objetivo : Testar se drive esta pronto
                    * Parametros : Drive
                    * ------------------------------------------------------------------
Function Testa_Drive(strDrive)
Local nArq
strDrive = strDrive + ':\TESTE.TST'
nArq := Fcreate(strDrive) // cria arquivo para testar o drive
If Ferror()#0  // drive nao existe ou nao esta formatado
    Fclose(strDrive) // ou nao exite disco no drive
    do mensagem with "Disco nao formatado ou nao existe disco no drive",8
    //
    // tratar o erro aqui
    //
    Return(.F.)
Else
    Ferase(strDrive) && apaga arquivo criado
    Return(.T.)
EndIf

Return(Nil)

procedure lim1
set color to /w
@ 17,02 say "³                                                  ³"
@ 18,02 say "³                                                  ³"
@ 19,02 say "³                                                  ³"
@ 20,02 say "³                                                  ³"
return
procedure lim2
set color to /w
@ 18,01 say "³                                                         ³"
@ 19,01 say "³                                                         ³"
@ 20,01 say "³                                                         ³"
@ 21,01 say "³                                                         ³"
return




procedure vl1
set color to /w
@ 17,02 say "³                                                  ³"
@ 18,02 say "³                                                  ³"
@ 19,02 say "³                                                  ³"
@ 20,02 say "³                                                  ³"
return


procedure vl2
set color to /w
@ 18,01 say "³                                                         ³"
@ 19,01 say "³                                                         ³"
@ 20,01 say "³                                                         ³"
@ 21,01 say "³                                                         ³"
return


procedure excluir
select 2
use estoque
set index to ind01.est,ind02.est
select 3
use ticket
go bott
zticket = strzero(val(ticket)+1,6,0)
select 4
use cticket
select 5
use saida
select 6
use csaida

select 5
go top
do while !eof()
    iguala()
    iguala()
    if wpago = ""
        select 2
        set index to ind01.est,ind02.est
        seek wcodigo_est
        if found()
            zsaldo = saldo
            zc_saldo = c_saldo
            if zc_saldo = "S"
                if rec_lock()
                    replace saldo with wquantidade + zsaldo
                    unlock
                endif
            endif
        endif
    endif
    select 6
    if add_rec(10)
        revar()
    endif
    select 5
    skip
enddo
select 3
go top
do while !eof()
    iguala()
    iguala()
    select 4
    if add_rec(10)
        revar()
    endif
    select 3
    skip
enddo

select 3
go top
do while !eof()
    iguala()
    if rec_lock()
        delete
    endif
    select 3
    skip
enddo
select 3
go top
if add_rec(10)
    replace ticket with zticket
endif

select 5
go top
do while !eof()
    iguala()
    if rec_lock()
        delete
    endif
    select 5
    skip
enddo

close all
return


*--------------------------------------------------------------------------



***************************************************************************                                                                           *
*                             F       I       M                           *
***************************************************************************
