*********************************************************************
* Programa cadastro de alunos
*
* Data : 11/07/94
*
* Programador : Luciano Moreira Araujo
*
*
*********************************************************************
*****> Abre arquivo de fornecedores

select 1
if !net_use("tipo",.f.,"ind01.tip")
    return
endif

do while .t.
    inicia()
    ycodigo    = 0
    wco = space(01)
    set color to n/w
    dt = quadro(08,21,11,76)
    set color to w/r
    @ 08,37 say " TIPOS DE DOCUMENTOS"
    set color to /w
    @ 09,23 say "Codigo.....:"
    @ 10,23 say "Descricao..:"
    set color to w/b
    @ 23,05 say "<ESC> Para sair."
    save screen to telat
    set color to w,n/bg,,,n/w
    @ 09,36 get wcodigo pict "999" when tab_tip(10,35,17,70,1)
    read
    set key 5 to
    set key 24 to
    if readkey() = 12  .or. lastkey() = 27 .or. empty(wcodigo)
        close all
        return
    endif
    restore screen from telat
    set color to /w
    @ 09,36 say wcodigo pict "999"
    select 1
    set index to ind01.tip
    seek wcodigo
    if found()
        limpa()
        wdescricao = descricao
        wdias      = dias
        wpercete   = percete
        set color to n/w
        @ 10,36 say wdescricao pict "@!"
        limpa()
        save screen to tela_an
        set color to w/b
        @ 23,05 say "<ESC> Para sair."
        SetColor("W/W")
        Janela2("W",.F.,03,61,07,70,.F.)
        SetColor("N/W,W+/R")
        @ 04,62   prompt "RETORNO "
        @ 05,62   prompt "ALTERAR "
        @ 06,62   prompt "EXCLUIR "
        menu to xop
        restore screen from tela_an
        do case
            case xop = 1
                limpa()
                loop
            case xop = 2
                limpa()
                set color to w,n/bg,,,n/w
                @ 10,36 get wdescricao pict "@!" valid !empty(wdescricao)
                read
                if lastkey() = 27 .or. readkey() = 12
                    @ 23,03 say space(75)
                    close all
                    return
                endif
                limpa()
                if Confirma ("Confirma Dados?")=1    && gravacao dos dados
                    limpa()
                    select 1
                    set index to ind01.tip
                    seek wcodigo
                    if found()
                        if rec_lock()
                            revar()
                            unlock
                        endif
                    endif
                    loop
                endif
                loop
            case xop = 3
                limpa()
                if Confirma ("Confirma Exclusao <S/N>?")=1    && gravacao dos dados     select 1
                    limpa()
                    select 1
                    set index to ind01.tip
                    seek wcodigo
                    if found()
                        if rec_lock()
                            delete
                        endif
                    endif
                endif
                loop
            other
                loop
        endcase
    endif
    limpa()
    @ 23,05 say "<ESC> Retorna."
    set color to w,n/bg,,,n/w
    @ 10,36 get wdescricao  pict "@!"
    read
    limpa()
    if Confirma ("Confirma Dados?")=1    && gravacao dos dados
        limpa()
        select 1
        set index to ind01.tip
        if add_rec(10)
            revar()
        endif
        loop
    endif
enddo
close all
return
    *********************************************
function tab_tip(x1,y1,x2,y2,sel)
sele &sel
set orde to 2
go top
set color to w/b
@ 23,01 say space(79)
@ 23,10 say "<"+chr(24)+"> Passa Tela. <"+chr(25)+"> Retorna Tela. <ESC> Retorna."
public telav:=savescreen(x1,y1,x2,y2)
@ x1,y1 clear to x2,y2
dt = quadro(x1,y1,x2,y2+6)
public xx1,xx2,yy1,yy2
xx1 = x1
xx2 = x2
yy1 = y1
yy2 = y2
linha = x1+1
@ linha,y1+1     say "Codigo"
@ linha,col()+5  say "Tipo do Documento"
linha++
@ linha,y1+1 to linha,(y2+6)-1
linha++
for x = 1 to 04
    @ linha,yy1+1    say codigo     pict "999"
    @ linha,col()+05 say descricao  pict "@s20"
    if paga = "1"
        @ linha,col()+3 say "P"
    endif
    if paga = "2"
        @ linha,col()+3 say "R"
    endif
    linha++

    skip
next
set key 5  to sobe_tip
set key 24 to desce_tip
return .t.

function sobe_tip
linha = 0
@ xx1,yy1 clear to xx2,yy2
dt = quadro(xx1,yy1,xx2,yy2+6)
linha = xx1+1
@ linha,yy1+1     say "Codigo"
@ linha,col()+5  say "Tipo do Documento"
linha++
@ linha,yy1+1 to linha,(yy2+6)-1
linha++
skip -10
for x = 1 to 04
    @ linha,yy1+1    say codigo     pict "999"
    @ linha,col()+05 say descricao  pict "@s20"
    if paga = "1"
        @ linha,col()+3 say "P"
    endif
    if paga = "2"
        @ linha,col()+3 say "R"
    endif
    linha++
    skip
next
return .t.

function desce_tip
@ xx1,yy1 clear to xx2,yy2
dt = quadro(xx1,yy1,xx2,yy2+6)
linha = xx1+1
@ linha,yy1+1    say "Codigo"
@ linha,col()+5  say "Tipo do Documento"
linha++
@ linha,yy1+1 to linha,(yy2+6)-1
linha++
if .not. eof()
    for x = 1 to 04
        @ linha,yy1+1    say codigo     pict "999"
        @ linha,col()+05 say descricao  pict "@s20"
        if paga = "1"
            @ linha,col()+3 say "P"
        endif
        if paga = "2"
            @ linha,col()+3 say "R"
        endif
        linha++
        skip
    next
endif
return .t.
    *********************************************
