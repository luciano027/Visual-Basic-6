***************************************************************************
* Recebe um arquivo texto, cria um dbf, grava  neste arquivo e mostra na
* tela, quando sair os arquivos temporarios criados serao apagados.
***************************************************************************
Proce lertexto
Para arquivo
#INCLUDE "INKEY.CH"
#INCLUDE "SETCURS.CH"
local1:= {}
declare acamp:=[15]
cor_ant = setcolor()    // Salva cor antiga
dbf_veio= select()      // Salva arquivo que estava selecionado
aadd(Local1, {"TELA01", "C", 13, 0})
aadd(Local1, {"TELA02", "C", 13, 0})
aadd(Local1, {"TELA03", "C", 13, 0})
aadd(Local1, {"TELA04", "C", 13, 0})
aadd(Local1, {"TELA05", "C", 13, 0})
aadd(Local1, {"TELA06", "C", 13, 0})
aadd(Local1, {"TELA07", "C", 13, 0})
aadd(Local1, {"TELA08", "C", 13, 0})
aadd(Local1, {"TELA09", "C", 13, 0})
aadd(Local1, {"TELA10", "C", 13, 0})
aadd(Local1, {"TELA11", "C", 13, 0})
aadd(Local1, {"TELA12", "C", 13, 0})
aadd(Local1, {"TELA13", "C", 13, 0})
aadd(Local1, {"TELA14", "C", 13, 0})
aadd(Local1, {"TELA15", "C", 13, 0})
zdbf = alltrim(zmicro)
dbcreate(zdbf, Local1)
select 99
use &zdbf alias &zdbf
append from &arquivo sdf   //nomecompl_arq sdf
goto top
@ 02,00 to 22,79 colo 'w+/b' //BOX CHR(219)
//               /-------------> Cor do TBrowse
//              /    /---------> Cor da barra em destaque
//             /    /      /---> Cor da barra luminosa
set color to "w+/b,n/w,,,g+/b"

oTab := TBrowseDB(04,01,21,78)
oTab:colsep := ""  // CHR(219)+" "
OTAB:HEADSEP:= 'Ä' // CHR(205)+CHR(203)+CHR(205)
OTAB:COLSEP := ''  // CHR(032)+CHR(186)+CHR(032)
OTAB:FOOTSEP:= 'Ä' // CHR(205)+CHR(202)+CHR(205)

for i = 1 to fcount()
    OTAB:ADDCOLUMN(TBCOLUMNNEW("", FIELDWBLOCK(FIELDNAME(i),SELE() )))
next

@ 03,01 say '                Tecle <ESC> para Sair ou <F10> para Imprimir                  ' colo 'w-/b'
lin_br=30
Do while .T.
    nRow := ROW()
    nCol := COL()

    do while !oTab:Stabilize() ; enddo

        oTAB:hilite()

        setpos( nRow, nCol )
        nKey := inkey()

        IF oTAB:stable
            setpos( nRow, nCol )
            nKey := Inkey(0)
        endif
        if nKey == K_F1
        elseif nKey == K_F2
        elseif nKey == K_F3
        elseif nKey == K_F4
        elseif nKey == K_F5
        elseif nKey == K_F6
        elseif nKey == K_F7
        elseif nKey == K_UP         ; oTab:UP()
        elseif nKey == K_DOWN       ; oTab:DOWN()
        elseif nKey == K_LEFT       ; oTab:LEFT()
        elseif nKey == K_CTRL_LEFT  ; oTab:PANLEFT()
        elseif nKey == K_RIGHT      ; oTab:RIGHT()
        elseif nKey == K_CTRL_RIGHT ; oTab:PANRIGHT()
        elseif nKey == K_PGUP       ; oTab:PAGEUP()
        elseif nKey == K_CTRL_PGUP  ; oTab:GOTOP()
        elseif nKey == K_PGDN       ; oTab:PAGEDOWN()
        elseif nKey == K_CTRL_PGDN  ; oTab:GOBOTTOM()
        elseif nKey == K_HOME       ; oTab:HOME()
        elseif nKey == K_CTRL_HOME  ; oTab:PANHOME()
        elseif nKey == K_END        ; oTab:END()
        elseif nKey == K_CTRL_END   ; oTab:PANEND()
        elseif nKey == K_ESC        ; EXIT
        elseif nkey == 99 .or. nkey==67
            if otab:freeze=0
                otab:freeze := otab:colpos -1
                lin_br=21
            else
                otab:freeze = 0
                lin_br= 30
                otab:panhome()
                otab:refreshall()
            endif
        elseif  nkey == K_F10
            if !impri()
                return
            endif
            set color to w+/b
            dt = quadro(15,18,17,35)
            set color to *+w/b
            @ 16,20 say "Imprimindo..."
            copy file &zdireto to prn
            exit

        endif

        oTab:colorRect({1,1,lin_br,oTab:freeze},{2,5})

    enddo
    arq=dbf()
    use
    arq_dbf=alltrim(arq)+'.DBF'
    erase &arq_dbf   // apaga o arquivo dbf que mostra na tela a impressao.
                                                             **********  erase &arquivo   // apaga o arquivo prn que mostra na tela a impressao.
    sele &dbf_veio   // seleciona novamente o arquivo que estava aberto
    set color to '&cor_ant.'


