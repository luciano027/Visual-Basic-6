*****************************************************************************
* Programa...: CADSAI.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 02/12/98          Data da Ultima Atualizacao:                *
* Objetivo...: SAIDA DE PRODUTOS NO ESTOQUE - BAIXA ACERTO SALDO            *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************
select 1
if !net_use("estoque",.f.,"ind01.est","ind02.est")
    return
endif
select 2
if !net_use("fornece",.f.,"ind01.for","ind02.for")
    return
endif
select 3
if !net_use("saida",.f.,"ind01.sai","ind02.sai")
    return
endif


do while .t.
    select 1
    inicia()
    select 2
    inicia()
    select 3
    inicia()
    do telaent31
    set color to w/b
    @ 23,05 say "<ESC> Para sair."
    set color to w,n/bg,,,n/w
    save screen to telafor
    @ 04,17 get wnf pict "@!"
    read
    if readkey() = 12 .or. lastkey() = 27 .or. empty(wnf)
        libera()
        close all
        return
    endif
    qnf = wnf
    do while .t.
        select 3
        set index to ind02.sai
        seek qnf
        if found()
            limpa()
            iguala()
            set color to r/w
            @ 05,17 say wdata pict "@d"
            ynf      = wnf
            ydata    = wdata
            lir      = 10
            wtotal   = 0
            do while wnf = ynf
                set color to /w
                @ lir,07 say wcodigo_est   pict "@!"
                select 1
                set index to ind01.est
                seek wcodigo_est
                iguala()
                @ lir,18 say wdescricao    pict "@s30"
                @ lir,49 say wquantidade   pict "@e 99,999.99"
                @ lir,61 say wpreco_ven    pict "@e 99,999.99"
                lir = lir + 1
                if lir > 16
                    inkey(.5)
                    do limpa1
                    lir = 10
                endif
                wtotal = wtotal + (wquantidade*wpreco_ven)
                zitem = val(witem)+1
                select 3
                skip
                iguala()
            enddo
            wnf       = ynf
            wdata     = ydata
            wcod_ven  = space(03)
            @ 19,61 say wtotal     pict "@e 99,999.99"
            limpa()
            save screen to tela_an
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            SetColor("W/W")
            Janela2("W",.F.,03,61,08,72,.F.)
            SetColor("N/W,W+/R")
            @ 04,62   prompt "RETORNO  "
            @ 05,62   prompt "INCLUIR  "
            @ 06,62   prompt "CONSULTAR"
            @ 07,62   prompt "EXTRATO  "
            menu to xop
            restore screen from tela_an
            do case
                case xop = 1
                    limpa()
                    exit
                case xop = 2
                    limpa()
                    do while .t.
                        wcodigoest  = 0
                        wquantidade = 0
                        wnf         = ynf
                        wdata       = ydata
                        set color to w/b
                        @ 23,05 say "<ENTER> Para Consulta. <ESC> para Sair."
                        set color to n/bg
                        @ lir,07 get wcodigoest pict "999999"
                        read
                        if readkey() = 12 .or. lastkey() = 27
                            exit
                        endif
                        wcodigo_est = strzero(wcodigoest,6,0)
                        if empty(wcodigo_est) .or. wcodigo_est = "000000"
                            save screen to tela_a
                            select 1
                            index on descricao to indcon.pec
                            set index to indcon.pec
                            do pes_est
                            iguala()
                            restore screen from tela_a
                        endif
                        set color to /w
                        @ lir,07 say wcodigo_est pict "999999"
                        select 1
                        set index to ind01.est
                        seek wcodigo_est
                        if !found()
                            do mensagem with "Codigo nao Cadastrado...",8
                            loop
                        endif
                        iguala()
                        limpa()
                        set color to  /w
                        @ lir,18 say wdescricao   pict "@s30"
                        @ lir,61 say wpreco_ven    pict "@e 99,999.99"
                        set color to w,n/bg,,,n/w
                        @ lir,49 get wquantidade   pict "@e 99,999.99"
                        read
                        if wc_saldo = "S"
                            if wquantidade > wsaldo
                                do mensagem with " Quantidade maior que o disponivel...",8
                                loop
                            endif
                        endif
                        wtotal = wtotal + (wquantidade*wpreco_ven)
                        set color to /w
                        @ 19,61  say wtotal  pict "@e 99,999.99"
                        p = "S"
                        limpa()
                        @ 23,03 say "Confirma Compras <S/N>? " get p pict "@!" valid p $ "SN"
                        read
                        limpa()
                        if upper(p) = "S"
                            witem       = strzero(zitem,3,0)
                            wpreco      = wpreco_ven
                            wtipo_ven   = "F"
                            wcodigo_ven = space(03)
                            walicota    = " "
                            select 1
                            set index to ind01.est,ind02.est
                            seek wcodigo_est
                            iguala()
                            if wc_saldo = "S"
                                if rec_lock()
                                    replace saldo  with wsaldo-wquantidade
                                    unlock
                                endif
                            endif
                            select 3
                            set index to ind01.sai,ind02.sai
                            if add_rec(10)
                                revar()
                            endif
                            set color to /w
                            @ 19,61 say wtotal     pict "@e 99,999.99"
                            lir = lir + 1
                            if lir > 16
                                limpa1()
                                lir = 10
                            endif
                            zitem = zitem + 1
                            loop
                        else
                            wtotal = wtotal - (wquantidade*wpreco_ven)
                            set color to /w
                            @ 19,61 say wtotal     pict "@e 99,999.99"
                            loop
                        endif
                    enddo
                    limpa()
                    loop
                case xop = 3
                    save screen to tela09
                    select 3
                    index on nf to indcon.vis for nf = wnf
                    set index to indcon.vis
                    go top
                    limpa1()
                    do pescon_33 with 3
                    reg = recno()
                    locate for reg = recno()
                    set color to w/b
                    dt = quadro(09,20,15,75)
                    @ 10,22 say "Codigo....:"
                    @ 11,22 say "Descricao.:"
                    @ 12,22 say "Quantidade:"
                    @ 13,22 say "Pr.Compra.:"
                    wcodigo_est   = codigo_est
                    yquantidade = quantidade
                    wquantidade = quantidade
                    select 1
                    set index to ind01.est
                    seek wcodigo_est
                    iguala()
                    wvalor     = wpreco_ven * wquantidade
                    @ 10,33  say wcodigo_est   pict "@!"
                    @ 11,33  say wdescricao  pict "!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
                    @ 12,33  say wquantidade pict "@e 99,999.99"
                    @ 13,33  say wvalor      pict "@e 99,999.99"
                    limpa()
                    set color to  w/b
                    @ 14,22      prompt " RETORNO "
                    @ 14,col()+2 prompt " ALTERAR "
                    @ 14,col()+2 prompt " EXCLUIR "
                    menu to qop
                    do case
                        case qop = 1
                            set color to w/b
                            @ 23,01 say space(79)
                            restore screen from tela09
                            loop
                        case qop = 2
                            do while .t.
                                ycodigo_est = wcodigo_est
                                yquant      = wquantidade
                                ydata       = wdata
                                set color to w/b
                                @ 10,33 say wcodigo_est   pict "999999"
                                @ 11,33 say wdescricao   pict "!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
                                set color to w,n/bg,,,n/w
                                @ 12,33 get wquantidade pict "@e 99,999.99"
                                read
                                if wc_saldo = "S"
                                    if wquantidade > wsaldo+yquant
                                        do mensagem with " Quantidade maior que o disponivel...",8
                                        loop
                                    endif
                                endif
                                exit
                            enddo
                            p = "S"
                            wtotal = wtotal + (wpreco_ven  * wquantidade)
                            limpa()
                            @ 23,03 say "Confirma Dados <S/N> ?" get p pict "@!" valid p $ "SN"
                            read
                            limpa()
                            if upper(p) = "S"
                                zquant = wquantidade
                                select 1
                                set index to ind01.est,ind02.est
                                seek wcodigo_est
                                iguala()
                                if wc_saldo = "S"
                                    if rec_lock()
                                        replace saldo  with (wsaldo+yquant)-wquantidade
                                        unlock
                                    endif
                                endif
                                select 3
                                set index to ind01.sai,ind02.sai
                                go reg
                                iguala()
                                wt_qua =  (wquantidade - yquant)+zquant
                                if wt_qua < 0
                                    wt_qua = 0
                                endif
                                if rec_lock()
                                    replace quantidade with wt_qua
                                    unlock
                                endif
                                lir = lir + 1
                                if lir > 16
                                    limpa1()
                                    lir = 10
                                endif
                                restore screen from tela09
                                loop
                            endif
                            restore screen from tela09
                            loop
                        case qop = 3
                            yquant      = wquantidade
                            p = "S"
                            limpa()
                            @ 23,03 say "Confirma Delecao <S/N> ?" get p pict "@!" valid p $ "SN"
                            read
                            limpa()
                            if upper(p) = "S"
                                qual = witem
                                select 1
                                set index to ind01.est,ind02.est
                                seek wcodigo_est
                                iguala()
                                if wc_saldo = "S"
                                    if rec_lock()
                                        replace saldo  with wsaldo+yquant
                                        unlock
                                    endif
                                endif
                                select 3
                                set index to ind01.sai,ind02.sai
                                go reg
                                if rec_lock()
                                    delete
                                endif
                                restore screen from tela09
                                limpa1()
                                loop
                            endif
                        other
                            restore screen from tela09
                            loop
                    endcase
                    restore screen from tela09
                    limpa1()
                    loop
                case xop = 4
                    save screen to tela12
                    do while .t.
                        select 3
                        index on nf to indext.vis for nf = qnf
                        set index to indext.vis
                        limpa()
                        set color to w+/r
                        dt = quadro(15,18,17,35)
                        set color to *+w/r
                        @ 16,20 say "Aguarde..."
                        set color to w/b
                        set print to &zdireto
                        set devi to print
                        set cons off
                        wpag = 1
                        geral = 0
                        do while inkey()#27 .and. !eof()
                            @ prow(),02    say zempresa  pict "@!"
                            @ prow()+1,02  say zendereco pict "@!"
                            @ prow()+2,02  say "Extrato Saida de Mercadorias"
                            @ prow(),72    say "Pag."
                            @ prow(),76    say wpag pict "999"
                            @ prow()+1,02     say "Data........:"
                            @ prow(),pcol()+1 say wdata pict "@d"
                            @ prow()+2,02  say"Codigo    Especificacao                  Quantidade Preco      Valor     "
                            @ prow()+1,02  say"~~~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~ ~~~~~~~~~~ ~~~~~~~~~~"
                            pagina = 0
                            do while prow()<50 .and. !eof()
                                iguala()
                                wquant = wquantidade
                                select 1
                                seek wcodigo_est
                                iguala()
                                @ prow()+1,02 say wcodigo_est pict "@!"
                                @ prow(),12   say wdescricao  pict "!!!!!!!!!!!!!!!!!!!!!!!!"
                                @ prow(),43   say wquant      pict "@e@z 99,999.99"
                                @ prow(),54   say wpreco_ven  pict "@e@z 99,999.99"
                                compras = (wquant * wpreco_ven)
                                @ prow(),65   say compras    pict "999,999.99"
                                wlin = 1
                                pagina = pagina + compras
                                select 3
                                skip
                                esc = inkey()
                                if esc = 12
                                    @ prow()+2,40 say "***********  Impressao Cancelada..."
                                    close all
                                    return
                                endif
                            enddo
                            geral = geral + pagina
                            @ prow()+1,01  say "-------------------------------------------------------------------------"
                            @ prow()+1,03  say "Total na Pagina........."
                            @ prow(),63    say pagina pict "@e@z 999,999.99"
                            @ prow()+1,03  say "Total Geral............."
                            @ prow(),63    say geral  pict "@e@z 999,999.99"
                            @ prow()+1,01  say "-------------------------------------------------------------------------"
                            @ prow()+2,01  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
                            @ prow()+1,01  say "              Sistema de Administracao Estoque - Versao 1.0"
                            @ prow(),02    say date() pict "@d"
                            @ prow(),69    say time() pict "99:99"
                            eject
                            wpag = wpag + 1
                        enddo
                        set cons on
                        set devi to screen
                        set print off
                        set print to
                        lertexto(zdireto)
                        exit
                    enddo
                    restore screen from tela12
                    loop
                other
                    restore screen from tela_an
                    loop
            endcase
            loop
        endif
        limpa()
        lir = 10
        t = "S"
        total = 0
        wdata = date()
        zitem = 1
        set color to w/b
        @ 23,05 say "<ESC> Para sair."
        set color to w,n/bg,,,n/w
        @ 05,17 get wdata pict "@d"
        read
        if readkey() = 12 .or. lastkey() = 27
            saida = .t.
            return
        endif
        do while .t.
            wcodigoest  = 0
            wquant      = 0
            set color to w/b
            @ 23,05 say "<ENTER> Para Consulta. <ESC> para Sair."
            set color to n/bg
            @ lir,07 get wcodigoest pict "999999"
            read
            if readkey() = 12 .or. lastkey() = 27
                exit
            endif
            wcodigo_est = strzero(wcodigoest,6,0)
            if empty(wcodigo_est) .or. wcodigo_est = "000000"
                save screen to tela_a
                select 1
                index on descricao  to indcon.pec
                set index to indcon.pec
                do pes_est
                iguala()
                restore screen from tela_a
            endif
            set color to /w
            @ lir,07 say wcodigo_est pict "@!"
            select 1
            set index to ind01.est
            seek wcodigo_est
            if !found()
                do mensagem with "Codigo nao Cadastrado...",8
                loop
            endif
            iguala()
            limpa()
            set color to  /w
            @ lir,18 say wdescricao     pict "@s30"
            @ lir,61 say wpreco_ven     pict "@e@z 99,999.99"
            set color to w,n/bg,,,n/w
            @ lir,49 get wquant         pict "@e 99,999.99"
            read
            if wc_saldo = "S"
                if wquant > wsaldo
                    do mensagem with " Quantidade maior que o disponivel...",8
                    loop
                endif
            endif
            total = total + (wquant*wpreco_ven)
            set color to /w
            @ 19,61  say total         pict "@e 99,999.99"
            p = "S"
            limpa()
            @ 23,03 say "Confirma Dados <S/N>? " get p pict "@!" valid p $ "SN"
            read
            limpa()
            if upper(p) = "S"
                witem       = strzero(zitem,3,0)
                wpreco      = wpreco_ven
                wtipo_ven   = "F"
                wquantidade = wquant
                wcodigo_ven = space(03)
                walicota    = " "
                select 3
                set index to ind01.sai,ind02.sai
                if add_rec(10)
                    revar()
                endif
                select 1
                set index to ind01.est,ind02.est
                seek wcodigo_est
                iguala()
                if wc_saldo = "S"
                    if rec_lock()
                        replace saldo      with wsaldo - wquantidade
                        replace preco_ven  with wpreco_ven
                        unlock
                    endif
                endif
                set color to /w
                @ 19,61 say total      pict "@e 99,999.99"
                zitem = zitem + 1
                lir = lir + 1
                if lir > 16
                    limpa1()
                    lir = 10
                endif
            else
                total = total - (wquantidade*wpreco_ven)
                set color to /w
                @ 19,61 say total      pict "@e 99,999.99"
            endif
        enddo
    enddo
    loop
enddo
close all
return

procedure telaent31
set color to w/b
@ 23,05 say "<ESC> Para sair."
SetColor("W/W")
quadro(03,05,21,72)
Janela2("W",.F.,03,05,21,72,.F.)
set color to w/r
@ 03,27 say "  ACERTO NO ESTOQUE  "
SetColor("N/W,W+/R")
@ 04,06 say "Documento.:"
@ 05,06 say "Data......:"
@ 07,06 say "ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄ¿"
@ 08,06 say "³  Codigo  ³         Descricao            ³ Quantid. ³   Valor   ³"
@ 09,06 say "ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄ´"
@ 10,06 say "³          ³                              ³          ³           ³"
@ 11,06 say "³          ³                              ³          ³           ³"
@ 12,06 say "³          ³                              ³          ³           ³"
@ 13,06 say "³          ³                              ³          ³           ³"
@ 14,06 say "³          ³                              ³          ³           ³"
@ 15,06 say "³          ³                              ³          ³           ³"
@ 16,06 say "³          ³                              ³          ³           ³"
@ 17,06 say "³          ³                              ³          ³           ³"
@ 18,06 say "ÃÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄ´"
@ 19,06 say "³                                  T O T A L                     ³"
@ 20,06 say "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ"
return

procedure limpa1
set color to /w
@ 10,06 say "³          ³                              ³          ³           ³"
@ 11,06 say "³          ³                              ³          ³           ³"
@ 12,06 say "³          ³                              ³          ³           ³"
@ 13,06 say "³          ³                              ³          ³           ³"
@ 14,06 say "³          ³                              ³          ³           ³"
@ 15,06 say "³          ³                              ³          ³           ³"
@ 16,06 say "³          ³                              ³          ³           ³"
@ 17,06 say "³          ³                              ³          ³           ³"
return


procedure pescon_33
para sele
select &sele
go top
sai=.f.
lp=10
tk=0
ul=22
set color to w/b
@ 23,03 say "Sobe :"+chr(24)+" Desce :"+chr(25)+" <ENTER> Confirma  <ESC> Sai "
set color to /w
do while .not. sai
    l=10   && primeira linha da tela l = 8
    do while l<=17 .and. .not. eof()
        do dados66c with l
        skip
        l=l+1
    enddo
    nl=l
    do while nl<=17
        nl=nl+1
    enddo
    l=l-1
    ul=l
    if tk<0 .or. tk=22
        goto rec
        lp=10
    else
        skip lp-(l+1)
    endif
    pl=.f.
    if l<22
        if lp>10
            lp=l
            go bottom
        endif
        pl=.t.
    endif
    set color to w+/r
    do dados66c with lp
    set color to /w
    do while .t.
        tk=inkey(0)
        do case
            case tk=24
                skip
                if eof()
                    do mensagem with "Fim do Arquivo !",4
                    set color to /w
                    skip -1
                    loop
                else
                    skip -1
                endif
                if lp<15
                    do dados66c with lp
                    lp=lp+1
                    skip
                    set color to w+/r
                    do dados66c with lp
                    set color to /w
                    loop
                else
                    set color to /w
                    do dados66c with lp
                    skip
                    if .not. eof()
                        scroll(10,06,17,71,1)
                        @ lp,06 say "³          ³                              ³          ³           ³"
                        @ 17,06 say "³          ³                              ³          ³           ³"
                    endif
                    set color to w+/r
                    do dados66c with lp
                    set color to /w
                    loop
                endif
            case tk=5
                if bof()
                    do mensagem with "Inicio do Arquivo !",4
                    set color to /w
                    loop
                endif
                if lp>10
                    if eof()
                        skip -1
                    endif
                    do dados66c with lp
                    lp=lp-1
                    skip -1
                    set color to w+/r
                    do dados66c with lp
                    set color to /w
                    loop
                else
                    set color to /w
                    if eof()
                        skip -1
                    endif
                    do dados66c with lp
                    skip -1
                    if .not. bof()
                        scroll(10,06,17,71,-1)
                        @ lp,06 say "³          ³                              ³          ³           ³"
                        @ 17,06 say "³          ³                              ³          ³           ³"
                    endif
                    set color to w+/r
                    do dados66c with lp
                    set color to /w
                    loop
                endif
            case tk=13
                save screen
                set color to w+/r
                t = recno()
                return(t)
            case tk=27
                sai=.t.
                t = 0
                return(t)
            other
                loop
        endcase
    enddo
    set color to /w
enddo

    **********************************************************
procedure dados66c
    * Apresentacao de uma linha de dados para consulta na tela
    **********************************************************
parameters linha,valor
iguala()
@ linha,07  say wcodigo_est   pict "@!"
select 1
set index to ind01.est
seek wcodigo_est
iguala()
@ linha,18 say wdescricao    pict "@s30"
@ linha,49 say wquantidade   pict "@e 99,999.99"
@ linha,61 say wpreco_ven    pict "@e 99,999.99"
select 3
return

    *****************************************************************************
    *                          F   I   M                                        *
    *****************************************************************************

