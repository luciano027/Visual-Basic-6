*****************************************************************************
* Programa...: CADVEN.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 02/12/98          Data da Ultima Atualizacao:                *
* Objetivo...: CADASTRO DE VENDEDORES                                       *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************
select 1
use empresa
select 2
use impresso
if !file("ind01.imp")
    index on codigo_imp to ind01.imp
endif
set index to ind01.imp
do while .t.
    select 1
    inicia()
    select 2
    inicia()
    Janela2("W",.F.,08,21,11,76,.F.)
    set color to w/r
    @ 08,37 say "CONFIGURA IMPRESSORA"
    SetColor("N/W,W+/R")
    @ 09,23 say "Codigo....:"
    @ 10,23 say "Impressora:"
    select 1
    locate for !empty(empresa)
    if found()
        iguala()
        if !empty(wcodigo_imp)
            select 2
            set index to ind01.imp
            seek wcodigo_imp
            if found()
                iguala()
                set color to r/w
                @ 09,34 say wcodigo_imp pict "@!"
                @ 10,34 say wtipo       pict "@!"
                limpa()
                save screen to tela_an
                set color to w/b
                @ 23,05 say "<ESC> Para sair."
                Janela2("W",.F.,04,61,07,70,.F.)
                SetColor("N/W,W+/R")
                @ 05,62   prompt "RETORNO "
                @ 06,62   prompt "ALTERAR "
                menu to xop
                restore screen from tela_an
                do case
                    case xop = 1
                        limpa()
                        close all
                        return
                    case xop = 2
                        limpa()
                        do while .t.
                            @ 23,05 say "<ESC> Retorna."
                            set color to w,n/bg,,,n/w
                            @ 09,34 get wcodigo_imp pict "9" valid !empty(wcodigo_imp)
                            read
                            if readkey() = 12  .or. lastkey() = 27
                                close all
                                return
                            endif
                            if wcodigo_imp = "0"
                                save screen to telave
                                select 2
                                do pes_imp
                                restore screen from telave
                                iguala()
                            endif
                            set color to /w
                            @ 09,34 say wcodigo_imp pict "9"
                            select 2
                            set index to ind01.imp
                            seek wcodigo_imp
                            if found()
                                iguala()
                                set color to r/w
                                @ 09,34 say wcodigo_imp pict "@!"
                                @ 10,34 say wtipo       pict "@!"
                            else
                                do mensagem with "Codigo da impressora nao cadastrado...",8
                                loop
                            endif
                            exit
                        enddo
                        limpa()
                        if Confirma (" Dados Corretos? ")=1    && gravacao dos dados
                            select 1
                            locate for !empty(empresa)
                            replace codigo_imp with wcodigo_imp
                        endif
                        loop
                    other
                        loop
                endcase
            endif
        else
            set color to w/b
            @ 23,05 say "<ESC> Para sair.<0> para Consulta."
            set color to w,n/bg,,,n/w
            @ 09,34 get wcodigo_imp pict "9"
            read
            if readkey() = 12  .or. lastkey() = 27 .or. empty(wcodigo_imp)
                close all
                return
            endif
            if wcodigo_imp = "0"
                save screen to telave
                select 2
                do pes_imp
                restore screen from telave
                iguala()
            endif
            set color to /w
            @ 09,34 say wcodigo_imp pict "9"
            select 2
            set index to ind01.imp
            seek wcodigo_imp
            if found()
                iguala()
                set color to r/w
                @ 09,34 say wcodigo_imp pict "@!"
                @ 10,34 say wtipo       pict "@!"
                if Confirma (" Dados Corretos? ")=1    && gravacao dos dados
                    select 1
                    locate for !empty(empresa)
                    replace codigo_imp with wcodigo_imp
                endif
            endif
            close all
            return
        endif
    else
        do mensagem with "Empresa nao cadastrada...",8
        close all
        return
    endif
enddo
close all
return

****************************************************************************

procedure pes_imp
set index to ind02.ven
go top
sai=.f.
set color to n/bg
lp=06
tk=00
ul=09
save screen to tela93
set color to gr+/bg
dt = quadro(04,01,19,56)
@ 05,03 say   "  Codigo              Impressora                "
for q = 06 to 16
    set color to w+/bg
    @ q,03 say "                                               "
next
set color to n/bg
@ 17,04 say "컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴"
set color to n/bg
do while .not. sai
    set color to gr+/b
    @ 23,03 say "Movimento : <HOME>Inicio  <F9>Consulta <Esc>Sai "
    l=6   && primeira linha da tela l = 9
    do while l<=16 .and. .not. eof()
        set color to /bg
        do dadosimp with l
        skip
        l=l+1
    enddo
    nl=l
    do while nl<=16
        @ nl,03 say space(46)
        nl=nl+1
    enddo
    l=l-1
    ul=l
    if tk<0 .or. tk=22
        goto rec
        lp=6
    else
        skip lp-(l+1)
    endif
    pl=.f.
    if l<16
        if lp>6
            lp = 6
        endif
        pl=.t.
    endif
    set color to w+/r
    do dadosimp with lp
    set color to n/bg
    do while .t.
        tk=inkey(0)
        do case
            case tk=24
                skip
                if eof()
                    do mensagem with "Fim do Arquivo !",4
                    set color to n/bg
                    loop
                else
                    skip -1
                endif
                if lp<16 &&-----> 20
                    do dadosimp with lp
                    lp=lp+1
                    skip
                    set color to w+/r
                    do dadosimp with lp
                    set color to n/bg
                    loop
                else
                    set color to n/bg
                    do dadosimp with lp
                    skip
                    if .not. eof()
                        scroll(6,3,16,54,1)
                    endif
                    set color to w+/r
                    do dadosimp with lp
                    set color to n/bg
                    loop
                endif
            case tk=5
                if bof()
                    do mensagem with "Inicio do Arquivo !",4
                    set color to n/bg
                    loop
                endif
                if lp>6
                    if eof()
                        skip -1
                    endif
                    do dadosimp with lp
                    lp=lp-1
                    skip -1
                    set color to w+/r
                    do dadosimp with lp
                    set color to n/bg
                    loop
                else
                    set color to n/bg
                    if eof()
                        skip -1
                    endif
                    do dadosimp with lp
                    skip -1
                    if .not. bof()
                        scroll(6,3,16,54,-1)
                    endif
                    set color to w+/r
                    do dadosimp with lp
                    set color to n/bg
                    loop
                endif
            case tk=13
                save screen
                set color to n/bg
                return(.t.)
            case tk=27
                sai=.t.
                t = 1
                return(.t.)
                exit
            case tk=18
                if bof()
                    do mensagem with "Inicio do Arquivo !",4
                    loop
                endif
                if eof()
                    if lastrec()>=14
                        skip -7
                    else
                        go top
                    endif
                else
                    if lastrec()>14
                        skip -(lp-9+7)
                    else
                        go top
                    endif
                endif
                exit
            case tk=3
                if ul<16
                    do mensagem with "Fim do Arquivo !",4
                    set color to n/bg
                    loop
                endif
                skip 9-lp
                exit
            case tk=1
                go top
                l  = 6
                lp = 9
                exit
            case tk=6
                go bottom
                l  = 6
                lp = 9
                exit
            otherwise
                loop
        endcase
    enddo
    limpa()
enddo
close all
return


**********************************************************
procedure dadosimp
* Apresentacao de uma linha de dados para consulta na tela
**********************************************************
parameters lin
iguala()
@ lin, 06 say wcodigo_imp pict "9"
@ lin, 17 say wtipo       pict "@s40"
return
**************************

*   exemplo como utilizar os codigo das impressoras  - Cores  HP
*   @ prow()+1,02 say &wnormal+"nornal"
*   @ prow()+1,02 say &wcondesado+"condesado"
*   @ prow()+1,02 say &wnormal+"nornal"
*   @ prow()+1,02 say &wexpandido+"expandido"
*   @ prow()+1,02 say &wnormal+"nornal"
*   @ prow()+1,02 say &wcores+&wazul+"Azul"
*   @ prow()+1,02 say &wcores+&wvermelho+"vermelho"
*   @ prow()+1,02 say &wcores+&wverde+"verde"
*   @ prow()+1,02 say &wcores+&wpreto+"preto"
