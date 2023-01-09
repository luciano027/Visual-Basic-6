*****************************************************************************
* Programa...: APAGAR.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao:                *
* Objetivo...: CADASTRO DE CONTAS A PAGAR                                   *
* Sistema....: SISPAN - Sistema de Panifica‡Æo                              *
*****************************************************************************
select 2
use tipo
if !file("ind01.tip")
    index on codigo to ind01.tip
endif
set index to ind01.tip
select 3
use receber
if !file("ind01.rec") .or. !file("ind02.rec") .or. !file("ind03.rec") .or. !file("ind04.rec")
    index on tipo+documento         to ind01.rec
    index on dtos(venci)            to ind02.rec
    index on dtos(pagto)            to ind03.rec
    index on tipo+documento+parcela to ind04.rec
endif
set index to ind01.rec,ind02.rec,ind03.rec,ind04.rec
do while .t.
    select 1
    inicia()
    select 2
    inicia()
    select 3
    inicia()
    Janela2("W",.F.,04,08,22,70,.F.)
    set color to r/bg
    @ 04,30 say "CADASTRO CONTAS A RECEBER"
    SetColor("N/W,W+/R")
    @ 05,09 say "Tipo........:"
    @ 06,09 say "Documento...:"
    @ 07,09 say "Cliente.....:"
    @ 08,09 say "Historico...:                        "
    @ 09,09 say "ÚÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄ¿"
    @ 10,09 say "³Parcela³ Vencimento ³   Valor   ³ Data Pagto ³   Valor   ³"
    @ 11,09 say "ÃÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄ´"
    @ 12,09 say "³       ³            ³           ³            ³           ³"
    @ 13,09 say "³       ³            ³           ³            ³           ³"
    @ 14,09 say "³       ³            ³           ³            ³           ³"
    @ 15,09 say "³       ³            ³           ³            ³           ³"
    @ 16,09 say "³       ³            ³           ³            ³           ³"
    @ 17,09 say "³       ³            ³           ³            ³           ³"
    @ 18,09 say "³       ³            ³           ³            ³           ³"
    @ 19,09 say "ÃÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄ´"
    @ 20,09 say "³  Total a Pagar :                 Total Pago.:           ³"
    @ 21,09 say "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ"
    limpa()
    save screen to telat
    do while .t.
        set color to w,n/bg,,,n/w
        @ 05,23 get wtipo pict "999" when tab_tip(07,23,14,58,2)
        read
        set key 5 to
        set key 24 to
        if readkey() = 12  .or. lastkey() = 27 .or. empty(wtipo)
            close all
            return
        endif
        restore screen from telat
        set color to /w
        @ 05,23 say wtipo pict "999"
        select 2
        set index to ind01.tip
        seek wtipo
        if !found()
            do mensagem with "Tipo de documento nao cadastrado...",8
            loop
        endif
        iguala()
        set color to /w
        @ 05,30 say wdescricao pict "@!"
        exit
    enddo
    save screen to tela_21
    set color to w,n/bg,,,n/w
    @ 06,23 get wdocumento pict "@!"  when tab_doc(07,23,20,67,3,wtipo)
    read
    set key  5 to
    set key 24 to
    restore screen from tela_21
    set color to r/w
    @ 06,23 say wdocumento pict "@!"
    ztipo      = wtipo
    zdocumento = wdocumento
    do while .t.
        select 3
        set index to ind01.rec
        seek ztipo+zdocumento
        if found()
            iguala()
            select 2
            set index to ind01.tip
            seek wtipo
            iguala()
            set color to r/w
            @ 05,23 say wtipo       pict "999"
            @ 05,30 say wdescricao  pict "@!"
            @ 06,23 say wdocumento  pict "@!"
            @ 07,23 say wcodigo_cli pict "@!"
            @ 08,23 say whistorico  pict "@!"
            ydocumento = wdocumento
            ytipo      = wtipo
            ycodigo_cli= wcodigo_cli
            yhistorico = whistorico
            totaa = 0
            totap = 0
            li = 12
            select 3
            index on parcela to indcon.rec for documento = wdocumento .and. tipo = ztipo
            set index to indcon.rec
            do while !eof()
                @ li,12 say wparcela    pict "@!"
                @ li,19 say wvenci      pict "@d"
                @ li,32 say wvalor_v    pict "@e@z 99,999.99"
                @ li,45 say wpagto      pict "@d"
                @ li,57 say wvalor_p    pict "@e@z 99,999.99"
                li = li + 1
                if li = 19
                    li = 12
                    lip()
                endif
                totaa = totaa + wvalor_v
                totap = totap + wvalor_p
                select 3
                skip
                iguala()
            enddo
            wdocumento = ydocumento
            wtipo      = ytipo
            wcodigo_cli= ycodigo_cli
            whistorico = yhistorico
            set color to r/w
            @ 20,28 say totaa-totap  pict "@e@z 99,999.99"
            @ 20,57 say totap        pict "@e@z 99,999.99"
            limpa()
            save screen to tela_a
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            SetColor("W/W")
            Janela2("W",.F.,03,66,09,75,.F.)
            SetColor("N/W,W+/R")
            @ 04,67   prompt "RETORNO "
            @ 05,67   prompt "ALTERAR "
            @ 06,67   prompt "EXCLUIR "
            @ 07,67   prompt "BAIXA   "
            @ 08,67   prompt "PARCELA "
            menu to xop
            restore screen from tela_a
            do case
                case xop = 1
                    limpa()
                    exit
                case xop = 2
                    save screen to telat
                    @ 23,05 say "<ESC> Para sair. "
                    set color to w,n/bg,,,n/w
                    @ 07,23 get wcodigo_cli pict "@!"
                    read
                    if readkey() = 12  .or. lastkey() = 27
                        close all
                        return
                    endif
                    set color to w,n/bg,,,n/w
                    @ 08,23 get whistorico pict "@!"
                    read
                    limpa()
                    if Confirma (" Dados Corretos? ")=1    && gravacao dos dados     select 1
                        select 3
                        set index to ind01.rec,ind02.rec,ind03.rec,ind04.rec
                        seek wtipo+wdocumento
                        if found()
                            replace codigo_cli with wcodigo_cli
                            replace historico with whistorico
                        endif
                    endif
                    loop
                case xop = 3
                    limpa()
                    if Confirma ("Confirma Exclusao ? Todo o Documento sera excluido? ")=1    && gravacao dos dados
                        select 3
                        index on parcela to indcon.rec for documento = ydocumento
                        set index to indcon.rec
                        do while !eof()
                            iguala()
                            if wdocumento = ydocumento
                                delete
                            endif
                            skip
                        enddo
                        exit
                    endif
                    loop
                case xop = 4
                    save screen to tela09
                    select 3
                    index on parcela to indcon.vis for documento = ydocumento
                    set index to indcon.vis
                    go top
                    do pescon_4 with 3
                    reg = recno()
                    locate for reg = recno()
                    iguala()
                    set color to w/b
                    dt = quadro(09,20,16,65)
                    @ 10,22 say "Parcela...:"
                    @ 11,22 say "Vencimento:"
                    @ 12,22 say "Valor.....:"
                    @ 13,22 say "Pagamento.:"
                    @ 14,22 say "Valor Pago:"
                    @ 10,33 say wparcela    pict "@!"
                    @ 11,33 say wvenci      pict "@d"
                    @ 12,33 say wvalor_v    pict "@e 99,999.99"
                    @ 13,33 say wpagto      pict "@d"
                    @ 14,33 say wvalor_p    pict "@e 99,999.99"
                    limpa()
                    set color to  w/b
                    @ 15,22      prompt "RETORNO "
                    @ 15,col()+2 prompt "ALTERAR "
                    @ 15,col()+2 prompt "EXCLUIR "
                    @ 15,col()+2 prompt "BAIXA "
                    menu to qop
                    do case
                        case qop = 1
                            set color to w/b
                            @ 23,01 say space(79)
                            restore screen from tela09
                            loop
                        case qop = 2
                            set color to w,n/bg,,,n/w
                            @ 10,33 get wparcela    pict "@!"
                            @ 11,33 get wvenci      pict "@d"
                            @ 12,33 get wvalor_v    pict "@e 99,999.99"
                            read
                            limpa()
                            if Confirma (" Dados Corretos? ")=1    && gravacao dos dados     select 1
                                select 3
                                set index to ind01.rec,ind02.rec,ind03.rec,ind04.rec
                                seek ztipo+zdocumento
                                if found()
                                    replace parcela    with wparcela
                                    replace venci      with wvenci
                                    replace valor_v    with wvalor_v
                                endif
                            endif
                            restore screen from tela09
                            loop
                        case qop = 3
                            limpa()
                            if Confirma ("Confirma Exclusao? ")=1    && gravacao dos dados     select 1
                                select 3
                                set index to ind04.rec,ind02.rec,ind03.rec,ind01.rec
                                seek ztipo+zdocumento+wparcela
                                if found()
                                    delete
                                endif
                            endif
                            restore screen from tela09
                            loop
                        case qop = 4
                            set color to w,n/bg,,,n/w
                            @ 13,33 get wpagto      pict "@d"
                            @ 14,33 get wvalor_p    pict "@e 99,999.99"
                            read
                            limpa()
                            if Confirma (" Dados Corretos? ")=1    && gravacao dos dados
                                zparcela = strzero(val(wparcela)+1,3,0)
                                wfalta   = wvalor_v - wvalor_p
                                select 3
                                set index to ind04.rec,ind02.rec,ind03.rec,ind01.rec
                                seek ztipo+zdocumento+wparcela
                                if found()
                                    replace pagto      with wpagto
                                    replace valor_p    with wvalor_p
                                    if wfalta > 0
                                        replace valor_v with wvalor_p
                                    endif
                                endif
                                if wfalta > 0
                                    select 3
                                    set index to ind04.rec,ind02.rec,ind03.rec,ind01.rec
                                    append blank
                                    replace documento  with wdocumento
                                    replace tipo       with wtipo
                                    replace codigo_cli with wcodigo_cli
                                    replace historico  with whistorico
                                    replace parcela    with zparcela
                                    replace valor_v    with wfalta
                                    replace venci      with wvenci
                                endif
                            endif
                            restore screen from tela09
                            loop
                        other
                            restore screen from tela09
                            loop
                    endcase
                    restore screen from tela09
                    loop
                case xop = 5
                    set color to w/b
                    @ 23,05 say "<ESC> Para sair."
                    set color to w,n/bg,,,n/w
                    @ li,12 get wparcela    pict "@!"
                    @ li,19 get wvenci      pict "@d"
                    @ li,32 get wvalor_v    pict "@e@z 99,999.99"
                    read
                    if readkey() = 12  .or. lastkey() = 27
                        exit
                    endif
                    limpa()
                    if Confirma (" Dados Corretos? ")=1    && gravacao dos dados     select 1
                        select 3
                        set index to ind01.rec,ind02.rec,ind03.rec,ind04.rec
                        append blank
                        revar()
                        li = li + 1
                        if li = 19
                            li = 12
                            lip()
                        endif
                    endif
                    loop
                other
                    loop
            endcase
            loop
        else
            save screen to telat
            set color to w/b
            @ 23,05 say "<ESC> Para sair. "
            set color to w,n/bg,,,n/w
            @ 07,23 get wcodigo_cli pict "@!"
            read
            if readkey() = 12  .or. lastkey() = 27
                close all
                return
            endif
            set color to w,n/bg,,,n/w
            @ 08,23 get whistorico pict "@!"
            @ 12,12 get wparcela    pict "@!"
            @ 12,19 get wvenci      pict "@d"
            @ 12,32 get wvalor_v    pict "@e@z 99,999.99"
            read
            limpa()
            if Confirma (" Dados Corretos? ")=1    && gravacao dos dados     select 1
                select 3
                set index to ind01.rec,ind02.rec,ind03.rec,ind04.rec
                append blank
                revar()
            endif
        endif
    enddo
enddo
close all
return


*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
