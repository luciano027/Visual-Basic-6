select 1
if !net_use("juros",.f.,"ind01.jur")
    return
endif
do while .t.
    select 1
    go bott
    iguala()
    ycodigo_ju = strzero(val(wcodigo_ju)+1,1,0)
    inicia()
    wcodigo_ju = ycodigo_ju
    set color to n/w
    dt = quadro(05,02,21,77)
    set color to w/r
    @ 05,25 say "  CADASTRO JUROS  "
    set color to /w
    @ 06,03 say "Codigo...:"
    @ 07,03 say "Descricao:"
    @ 08,03 say "Taxa Administativa:"
    @ 09,03 say "ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄTabela de FinanciamentoÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿"
    @ 10,03 say "³  Prestacao    Fator     Limite         ³ Prestacao    Fator     Limite ³"
    @ 11,03 say "³                                        ³                               ³"
    @ 12,03 say "³                                        ³                               ³"
    @ 13,03 say "³                                        ³                               ³"
    @ 14,03 say "³                                        ³                               ³"
    @ 15,03 say "³                                        ³                               ³"
    @ 16,03 say "³                                        ³                               ³"
    @ 17,03 say "³                                        ³                               ³"
    @ 18,03 say "³                                        ³                               ³"
    @ 19,03 say "³                                        ³                               ³"
    @ 20,03 say "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ"
    set color to w/b
    @ 23,05 say "<ESC> Para sair.<0> para consulta"
    set color to w,n/bg,,,n/w
    save screen to telafor
    @ 06,13 get wcodigo_ju pict "9"
    read
    if readkey() = 12 .or. lastkey() = 27
        libera()
        close all
        return
    endif
    if wcodigo_ju = "0"
        save screen to tela_0
        do pes_ju
        restore screen from tela_0
    endif
    ycodigo_ju = wcodigo_ju
    set color to /w
    @ 06,13 say ycodigo_ju pict "9"
    select 1
    set index to ind01.jur
    seek ycodigo_ju
    if found()
        iguala()
        set color to n/w
        @ 07,13 say wdesc_jur   pict "@!"
        @ 08,23 say wtaxa       pict "@e@z 99,999.99"
        @ 11,08 say wpr_1       pict "@!"
        @ 11,18 say wfator_1    pict "@e@z 9.99999999"
        @ 11,29 say wlimi_1     pict "@e@z 9,999.99"
        @ 11,49 say wpr_2        pict "@!"
        @ 11,57 say wfator_2    pict "@e@z 9.99999999"
        @ 11,67 say wlimi_2     pict "@e@z 9,999.99"
        @ 12,08 say wpr_3       pict "@!"
        @ 12,18 say wfator_3    pict "@e@z 9.99999999"
        @ 12,29 say wlimi_3     pict "@e@z 9,999.99"
        @ 12,49 say wpr_4        pict "@!"
        @ 12,57 say wfator_4    pict "@e@z 9.99999999"
        @ 12,67 say wlimi_4     pict "@e@z 9,999.99"
        @ 13,08 say wpr_5       pict "@!"
        @ 13,18 say wfator_5    pict "@e@z 9.99999999"
        @ 13,29 say wlimi_5     pict "@e@z 9,999.99"
        @ 13,49 say wpr_6        pict "@!"
        @ 13,57 say wfator_6    pict "@e@z 9.99999999"
        @ 13,67 say wlimi_6     pict "@e@z 9,999.99"
        @ 14,08 say wpr_7       pict "@!"
        @ 14,18 say wfator_7    pict "@e@z 9.99999999"
        @ 14,29 say wlimi_7     pict "@e@z 9,999.99"
        @ 14,49 say wpr_8        pict "@!"
        @ 14,57 say wfator_8    pict "@e@z 9.99999999"
        @ 14,67 say wlimi_8     pict "@e@z 9,999.99"
        @ 15,08 say wpr_9       pict "@!"
        @ 15,18 say wfator_9    pict "@e@z 9.99999999"
        @ 15,29 say wlimi_9     pict "@e@z 9,999.99"
        @ 15,49 say wpr_10       pict "@!"
        @ 15,57 say wfator_10   pict "@e@z 9.99999999"
        @ 15,67 say wlimi_10    pict "@e@z 9,999.99"
        @ 16,08 say wpr_11      pict "@!"
        @ 16,18 say wfator_11   pict "@e@z 9.99999999"
        @ 16,29 say wlimi_11    pict "@e@z 9,999.99"
        @ 16,49 say wpr_12       pict "@!"
        @ 16,57 say wfator_12   pict "@e@z 9.99999999"
        @ 16,67 say wlimi_12    pict "@e@z 9,999.99"
        @ 17,08 say wpr_13      pict "@!"
        @ 17,18 say wfator_13   pict "@e@z 9.99999999"
        @ 17,29 say wlimi_13    pict "@e@z 9,999.99"
        @ 17,49 say wpr_14       pict "@!"
        @ 17,57 say wfator_14   pict "@e@z 9.99999999"
        @ 17,67 say wlimi_14    pict "@e@z 9,999.99"
        @ 18,08 say wpr_15      pict "@!"
        @ 18,18 say wfator_15   pict "@e@z 9.99999999"
        @ 18,29 say wlimi_15    pict "@e@z 9,999.99"
        @ 18,49 say wpr_16       pict "@!"
        @ 18,57 say wfator_16   pict "@e@z 9.99999999"
        @ 18,67 say wlimi_16    pict "@e@z 9,999.99"
        @ 19,08 say wpr_17      pict "@!"
        @ 19,18 say wfator_17   pict "@e@z 9.99999999"
        @ 19,29 say wlimi_17    pict "@e@z 9,999.99"
        @ 19,49 say wpr_18       pict "@!"
        @ 19,57 say wfator_18   pict "@e@z 9.99999999"
        @ 19,67 say wlimi_18    pict "@e@z 9,999.99"
        limpa()
        save screen to tela_an
        set color to w/b
        @ 23,05 say "<ESC> Para sair."
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
                set color to w/b
                @ 23,05 say "<ESC> Retorna."
                set color to w,n/bg,,,n/w
                @ 07,13 get wdesc_jur   pict "@!"
                @ 08,23 get wtaxa       pict "@e@z 99,999.99"
                @ 11,08 get wpr_1       pict "@!"
                @ 11,18 get wfator_1    pict "@e@z 9.99999999"
                @ 11,29 get wlimi_1     pict "@e@z 9,999.99"
                @ 11,49 get wpr_2        pict "@!"
                @ 11,57 get wfator_2    pict "@e@z 9.99999999"
                @ 11,67 get wlimi_2     pict "@e@z 9,999.99"
                @ 12,08 get wpr_3       pict "@!"
                @ 12,18 get wfator_3    pict "@e@z 9.99999999"
                @ 12,29 get wlimi_3     pict "@e@z 9,999.99"
                @ 12,49 get wpr_4        pict "@!"
                @ 12,57 get wfator_4    pict "@e@z 9.99999999"
                @ 12,67 get wlimi_4     pict "@e@z 9,999.99"
                @ 13,08 get wpr_5       pict "@!"
                @ 13,18 get wfator_5    pict "@e@z 9.99999999"
                @ 13,29 get wlimi_5     pict "@e@z 9,999.99"
                @ 13,49 get wpr_6        pict "@!"
                @ 13,57 get wfator_6    pict "@e@z 9.99999999"
                @ 13,67 get wlimi_6     pict "@e@z 9,999.99"
                @ 14,08 get wpr_7       pict "@!"
                @ 14,18 get wfator_7    pict "@e@z 9.99999999"
                @ 14,29 get wlimi_7     pict "@e@z 9,999.99"
                @ 14,49 get wpr_8        pict "@!"
                @ 14,57 get wfator_8    pict "@e@z 9.99999999"
                @ 14,67 get wlimi_8     pict "@e@z 9,999.99"
                @ 15,08 get wpr_9       pict "@!"
                @ 15,18 get wfator_9    pict "@e@z 9.99999999"
                @ 15,29 get wlimi_9     pict "@e@z 9,999.99"
                @ 15,49 get wpr_10       pict "@!"
                @ 15,57 get wfator_10   pict "@e@z 9.99999999"
                @ 15,67 get wlimi_10    pict "@e@z 9,999.99"
                @ 16,08 get wpr_11      pict "@!"
                @ 16,18 get wfator_11   pict "@e@z 9.99999999"
                @ 16,29 get wlimi_11    pict "@e@z 9,999.99"
                @ 16,49 get wpr_12       pict "@!"
                @ 16,57 get wfator_12   pict "@e@z 9.99999999"
                @ 16,67 get wlimi_12    pict "@e@z 9,999.99"
        @ 17,08 get wpr_13      pict "@!"
        @ 17,18 get wfator_13   pict "@e@z 9.99999999"
        @ 17,29 get wlimi_13    pict "@e@z 9,999.99"
        @ 17,49 get wpr_14       pict "@!"
        @ 17,57 get wfator_14   pict "@e@z 9.99999999"
        @ 17,67 get wlimi_14    pict "@e@z 9,999.99"
        @ 18,08 get wpr_15      pict "@!"
        @ 18,18 get wfator_15   pict "@e@z 9.99999999"
        @ 18,29 get wlimi_15    pict "@e@z 9,999.99"
        @ 18,49 get wpr_16       pict "@!"
        @ 18,57 get wfator_16   pict "@e@z 9.99999999"
        @ 18,67 get wlimi_16    pict "@e@z 9,999.99"
        @ 19,08 get wpr_17      pict "@!"
        @ 19,18 get wfator_17   pict "@e@z 9.99999999"
        @ 19,29 get wlimi_17    pict "@e@z 9,999.99"
        @ 19,49 get wpr_18       pict "@!"
        @ 19,57 get wfator_18   pict "@e@z 9.99999999"
        @ 19,67 get wlimi_18    pict "@e@z 9,999.99"
                read
                if lastkey() = 27 .or. readkey() = 12
                    limpa()
                    close all
                    return
                endif
                limpa()
                if Confirma ("Confirma Dados?")=1    && gravacao dos dados
                    select 1
                    set index to ind01.jur
                    seek wcodigo_ju
                    if found()
                        if rec_lock()
                            revar()
                            unlock
                        endif
                    endif
                endif
                loop
            case xop = 3
                limpa()
                if Confirma ("Confirma Exclusao?")=1    && gravacao dos dados
                    limpa()
                    select 1
                    set index to ind01.jur
                    seek wcodigo_ju
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
    else
        limpa()
        set color to w/b
        @ 23,05 say "<ESC> Retorna."
        set color to w,n/bg,,,n/w
        @ 07,13 get wdesc_jur   pict "@!"
        @ 08,23 get wtaxa       pict "@e@z 99,999.99"
        @ 11,08 get wpr_1       pict "@!"
        @ 11,18 get wfator_1    pict "@e@z 9.99999999"
        @ 11,29 get wlimi_1     pict "@e@z 9,999.99"
        @ 11,49 get wpr_2        pict "@!"
        @ 11,57 get wfator_2    pict "@e@z 9.99999999"
        @ 11,67 get wlimi_2     pict "@e@z 9,999.99"
        @ 12,08 get wpr_3       pict "@!"
        @ 12,18 get wfator_3    pict "@e@z 9.99999999"
        @ 12,29 get wlimi_3     pict "@e@z 9,999.99"
        @ 12,49 get wpr_4        pict "@!"
        @ 12,57 get wfator_4    pict "@e@z 9.99999999"
        @ 12,67 get wlimi_4     pict "@e@z 9,999.99"
        @ 13,08 get wpr_5       pict "@!"
        @ 13,18 get wfator_5    pict "@e@z 9.99999999"
        @ 13,29 get wlimi_5     pict "@e@z 9,999.99"
        @ 13,49 get wpr_6        pict "@!"
        @ 13,57 get wfator_6    pict "@e@z 9.99999999"
        @ 13,67 get wlimi_6     pict "@e@z 9,999.99"
        @ 14,08 get wpr_7       pict "@!"
        @ 14,18 get wfator_7    pict "@e@z 9.99999999"
        @ 14,29 get wlimi_7     pict "@e@z 9,999.99"
        @ 14,49 get wpr_8        pict "@!"
        @ 14,57 get wfator_8    pict "@e@z 9.99999999"
        @ 14,67 get wlimi_8     pict "@e@z 9,999.99"
        @ 15,08 get wpr_9       pict "@!"
        @ 15,18 get wfator_9    pict "@e@z 9.99999999"
        @ 15,29 get wlimi_9     pict "@e@z 9,999.99"
        @ 15,49 get wpr_10       pict "@!"
        @ 15,57 get wfator_10   pict "@e@z 9.99999999"
        @ 15,67 get wlimi_10    pict "@e@z 9,999.99"
        @ 16,08 get wpr_11      pict "@!"
        @ 16,18 get wfator_11   pict "@e@z 9.99999999"
        @ 16,29 get wlimi_11    pict "@e@z 9,999.99"
        @ 16,49 get wpr_12       pict "@!"
        @ 16,57 get wfator_12   pict "@e@z 9.99999999"
        @ 16,67 get wlimi_12    pict "@e@z 9,999.99"
        @ 17,08 get wpr_13      pict "@!"
        @ 17,18 get wfator_13   pict "@e@z 9.99999999"
        @ 17,29 get wlimi_13    pict "@e@z 9,999.99"
        @ 17,49 get wpr_14       pict "@!"
        @ 17,57 get wfator_14   pict "@e@z 9.99999999"
        @ 17,67 get wlimi_14    pict "@e@z 9,999.99"
        @ 18,08 get wpr_15      pict "@!"
        @ 18,18 get wfator_15   pict "@e@z 9.99999999"
        @ 18,29 get wlimi_15    pict "@e@z 9,999.99"
        @ 18,49 get wpr_16       pict "@!"
        @ 18,57 get wfator_16   pict "@e@z 9.99999999"
        @ 18,67 get wlimi_16    pict "@e@z 9,999.99"
        @ 19,08 get wpr_17      pict "@!"
        @ 19,18 get wfator_17   pict "@e@z 9.99999999"
        @ 19,29 get wlimi_17    pict "@e@z 9,999.99"
        @ 19,49 get wpr_18       pict "@!"
        @ 19,57 get wfator_18   pict "@e@z 9.99999999"
        @ 19,67 get wlimi_18    pict "@e@z 9,999.99"
        read
        if lastkey() = 27 .or. readkey() = 12
            limpa()
            close all
            return
        endif
        if Confirma ("Confirma Dados?")=1    && gravacao dos dados
            limpa()
            select 1
            set index to ind01.jur
            if add_rec(10)
                revar()
            endif
            loop
        endif
    endif
    exit
enddo
close all
return

    ****************************************************************************

procedure pes_ju
go top
sai=.f.
set color to n/bg
lp=06
tk=00
ul=09
save screen to tela93
set color to gr+/bg
quadro(04,01,20,76)
@ 05,03 say "  Codigo                    Descricao              "
for q = 06 to 16
    set color to w+/bg
    @ q,03 say "                                                    "
next
set color to n/bg
@ 18,04 say "ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ"
set color to r/bg
do while .not. sai
    set color to gr+/b
    l=6   && primeira linha da tela l = 9
    do while l<=16 .and. .not. eof()
        set color to /bg
        do dadosc5 with l
        skip
        l=l+1
    enddo
    nl=l
    do while nl<=16
        @ nl,03 say space(52)
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
    do dadosc5 with lp
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
                    do dadosc5 with lp
                    lp=lp+1
                    skip
                    set color to w+/r
                    do dadosc5 with lp
                    set color to n/bg
                    loop
                else
                    set color to n/bg
                    do dadosc5 with lp
                    skip
                    if .not. eof()
                        scroll(6,3,16,54,1)
                    endif
                    set color to w+/r
                    do dadosc5 with lp
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
                    do dadosc5 with lp
                    lp=lp-1
                    skip -1
                    set color to w+/r
                    do dadosc5 with lp
                    set color to n/bg
                    loop
                else
                    set color to n/bg
                    if eof()
                        skip -1
                    endif
                    do dadosc5 with lp
                    skip -1
                    if .not. bof()
                        scroll(6,3,16,54,-1)
                    endif
                    set color to w+/r
                    do dadosc5 with lp
                    set color to n/bg
                    loop
                endif
            case tk=13
                save screen
                set color to n/bg
                sai = .t.
                t = 1
                iguala()
                return(wcodigo_ju)
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
                l  = 16
                lp = 6
                exit
            case tk=6
                go bottom
                l  = 16
                lp = 6
                exit
            case tk=-8
                do limpa
                set order to 2
                vnome=space(40)
                @ 23,10 say "Nome : " get vnome picture "@!"
                read
                do limpa
                if empty(vnome)
                    loop
                endif
                rec=recno()
                vnome=trim(vnome)
                seek vnome
                if .not. found()
                    do mensagem with "Desculpe, Produto n„o Encontrado !",6,1
                    set color to n/bg
                    go rec
                    loop
                endif
                rec=recno()
                l  = 16
                lp = 6
                exit
            otherwise
                loop
        endcase
    enddo
    do limpa
enddo
close all
return

    ************************
procedure dadosc5
parameters lin
iguala()
@ lin, 06 say wcodigo_ju pict "@!"
@ lin, 14 say wdesc_jur  pict "@s40"
return
    **************************

