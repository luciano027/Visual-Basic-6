*****************************************************************************
* Programa...: V_PRAZO.PRG                                                  *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......:                   Data da Ultima Atualizacao:                *
* Objetivo...: VENDAS A PRAZO                                               *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************
select 1
if !net_use("estoque",.f.,"ind01.est","ind02.est")
    return
endif

select 2
if !net_use("clientes",.f.,"ind01.cli","ind02.cli","ind03.cli")
    return
endif

select 3
if !net_use("prazo",.f.,"ind01.pra")
    return
endif

select 4
if !net_use("saida",.f.,"ind01.sai")
    return
endif

select 5
if !net_use("ticket",.f.,"ind01.tic")
    return
endif

select 6
if !net_use("caixa",.f.,"ind01.cai")
    return
endif
select 7
if !net_use("vendedor",.f.,"ind01.ven","ind02.ven")
    return
endif


do while .t.
    select 1
    inicia()
    select 2
    inicia()
    select 3
    inicia()
    select 4
    inicia()
    select 7
    inicia()
    select 5
    set index to ind01.tic
    go bott
    iguala()
    yticket = strzero(val(wticket)+1,6,0)
    wcodigo_cli = "00000"
    wchaves = 0
    wticket = yticket
    do telapra
    set color to w/b
    @ 23,05 say "<ESC> Para sair.<00000> para consulta"
    set color to w,n/bg,,,n/w
    save screen to telafor
    @ 04,09 get wcodigo_cli pict "99999"
    read
    if readkey() = 12 .or. lastkey() = 27
        libera()
        close all
        return
    endif
    if wcodigo_cli = "00000"
        select 2
        save screen to tela_0
        do pes_cli
        restore screen from tela_0
        iguala()
    endif
    set color to /w
    @ 04,09 say wcodigo_cli pict "99999"
    select 2
    set index to ind01.cli
    seek wcodigo_cli
    if !found()
        do mensagem with "Cliente nao cadastrado...",8
        loop
    endif
    iguala()
    limpa()
    set color to r/w
    @ 05,09 say wnome     pict "@!"
    @ 04,66 say wdata_deb pict "@d"
    @ 05,69 say wdebito   pict "@e@z 9,999.99"
    @ 06,69 say wcredito  pict "@e@z 9,999.99"
    ynome       = wnome
    ycodigo_cli = wcodigo_cli
    do while .t.
        select 3
        set index to ind01.pra
        seek ycodigo_cli
        if found()
            limpa()
            iguala()
            ydata = wdata
            select 2
            set index to ind01.cli
            seek wcodigo_cli
            iguala()
            ycodigo_cli = wcodigo_cli
            lir = 9
            wtotal = 0
            wdata = ydata
            do while wcodigo_cli = ycodigo_cli
                select 1
                set index to ind01.est
                seek wcodigo_est
                iguala()
                if wchaves = 1
                    set color to /w
                    @ lir,02 say wdata       pict "@D"
                    @ lir,13 say wcodigo_est pict "999999"
                    @ lir,20 say wdescricao pict "!!!!!!!!!!!!!!!!!!!"
                    @ lir,42 say wquantidade pict "@e@z 9999.99"
                    wprec = wvalor
                    wcompra = (wprec * wquantidade)
                    @ lir,50 say wcompra    pict "@e@z 9,999.99"
                    lir = lir + 1
                    if lir > 16
                        inkey(.5)
                        limpam()
                    endif
                endif
                if alltrim(wdescricao,1,5) != "CREDI"
                    wtotal = (wtotal + (wvalor * wquantidade))  &&----> imprime o total
                endif
                select 3
                skip
                iguala()
            enddo
            wcodigo_cli = ycodigo_cli
            set color to r/w
            lim2()
            NGrande(wtotal-wcredito, 18, 03, 52, "@E 999,999.99")
            @ 05,69 say wtotal          pict "@e@z 9,999.99"
            @ 06,69 say wcredito        pict "@e@z 9,999.99"
            @ 07,69 say wtotal-wcredito pict "@e@z 9,999.99"
            limpa()
            save screen to tela_an
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            Janela2("W",.F.,12,60,18,76,.F.)
            SetColor("N/W,W+/R")
            @ 13,61   prompt "Retorno       "
            @ 14,61   prompt "Pagamento     "
            @ 15,61   prompt "Consulta      "
            @ 16,61   prompt "Extrato       "
            @ 17,61   prompt "Credito       "
            menu to xop
            do case
                case xop = 1
                    restore screen from tela_AN
                    limpa()
                    select 2
                    set index to ind01.cli
                    seek ycodigo_cli
                    if found()
                        if rec_lock()
                            replace debito with wtotal
                            unlock
                        endif
                    endif
                    exit
                case xop = 2
                    restore screen from tela_AN
                    save screen to telapag
                    dt = quadro(14,54,18,77)
                    set color to w/r
                    @ 14,58 say "   Pagamento     "
                    set color to  /w
                    @ 15,55 say "Codigo..:"
                    @ 16,55 say "Nome....:"
                    @ 17,55 say "Total...:"
                    set color to r/w
                    @ 17,64 say wtotal-wcredito pict "@e@z 9,999.99"
                    do while .t.
                        set color to w/b
                        @ 23,05 say "<ESC> Para sair.<000> para Consulta."
                        set color to w,n/bg,,,n/w
                        @ 15,64 get wcodigo_ven pict "999" valid !empty(wcodigo_ven)
                        read
                        if readkey() = 12  .or. lastkey() = 27
                            close all
                            return
                        endif
                        if wcodigo_ven = "000"
                            select 7
                            save screen to telave
                            do pes_ven
                            restore screen from telave
                            iguala()
                        endif
                        set color to /w
                        @ 15,64 say wcodigo_ven pict "999"
                        select 7
                        set index to ind01.ven
                        seek wcodigo_ven
                        if !found()
                            do mensagem with "Desculpe ! Vendedor nao cadastrado...",8
                            loop
                        endif
                        iguala()
                        set color to r/w
                        @ 15,64 say wcodigo_ven  pict "@!"
                        @ 16,64 say wnome_ven pict "@s10"
                        ycodigo_ven = wcodigo_ven
                        ynome_ven   = wnome_ven
                        chave = 1
                        exit
                    enddo
                    limpa()
                    do while .t.
                        p = "N"
                        limpa()
                        set color to *gr+/b
                        @ 23,02 say "   Confirma Pagamento <S/N>? " get p pict "@!" valid p $ "SN"
                        read
                        limpa()
                        restore screen from telapag
                        if upper(p) = "S"
                            select 3
                            set index to ind01.pra
                            seek wcodigo_cli
                            if found()
                                ycodigo_cli = wcodigo_cli
                                do while wcodigo_cli = ycodigo_cli
                                    if ycodigo_cli = wcodigo_cli
                                        if rec_lock()
                                            delete
                                            unlock
                                        endif
                                    endif
                                    skip
                                    iguala()
                                enddo
                                select 2
                                set index to  ind01.cli,ind02.cli
                                seek ycodigo_cli
                                if found()
                                    if rec_lock()
                                        replace debito     with 0
                                        replace data_deb   with ctod("  /  /  ")
                                        replace credito    with 0
                                        unlock
                                    endif
                                endif
                                select 5
                                set index to ind01.tic
                                go bott
                                iguala()
                                wticket    = yticket
                                wdata_ti   = date()
                                wtotal_ti  = wtotal - wcredito
                                whistorico = "Recebimento :"+ynome
                                wpagamento = "P"
                                if add_rec(10)
                                    replace ticket    with wticket
                                    replace pagamento with "S"
                                    replace total_ti  with wtotal_ti
                                    replace historico with whistorico
                                    replace tipo_ti   with "D"
                                    replace tipo_tipo with "3"
                                    replace data_ti   with date()
                                    replace codigo_ven with wcodigo_ven
                                endif
                            endif
                        endif
                        exit
                    enddo
                    limpam()
                    exit
                case xop = 3
                    restore screen from tela_AN
                    save screen to tela09
                    select 3
                    index on codigo_cli to indcon.ind for codigo_cli = wcodigo_cli
                    set index to indcon.ind
                    do pescon_2 with 3
                    reg = recno()
                    locate for reg = recno()
                    set color to gr+/bg
                    dt = quadro(09,20,14,75)
                    @ 10,22 say "Codigo..:"
                    @ 11,22 say "Quantidade..:"
                    @ 12,22 say "Preco Unitario.:"
                    wcodigo_est = codigo_est
                    wquantidade = quantidade
                    select 1
                    set index to ind01.est
                    seek wcodigo_est
                    iguala()
                    wprec = wvalor
                    wvalor = (wprec * wquantidade)
                    set color to gr+/bg
                    @ 10,31      say wcodigo_est pict "999999"
                    @ 10,col()+5 say wdescricao  pict "!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
                    @ 11,34      say wquantidade pict "@e@z 9999.99"
                    @ 12,38      say wvalor      pict "@e@z 99,999.99"
                    do prazoexc
                    restore screen from tela09
                    limpam()
                    loop
                case xop = 4
                    restore screen from tela_AN
                    ycodigo_cli = wcodigo_cli
                    save screen to tela12
                    select 3
                    index on dtos(data) to &zmicro for  codigo_cli = wcodigo_cli
                    set index to &zmicro
                    limpa()
                    set color to w/b
                    set print to &zdireto
                    set devi to print
                    set cons off
                    wpag   = 1
                    geral  = 0
                    wtotal = 0
                    linh = 0
                    it = 1
                    do while inkey()#27 .and. !eof()
                        @ prow()+1,02     say zempresa  pict "@!"
                        @ prow()+1,02     say zendereco pict "@!"
                        @ prow()+1,02     say "Telefone: "
                        @ prow(),pcol()+1 say ztelefone pict "@!"
                        @ prow(),pcol()+1 say "FAX: "
                        @ prow(),pcol()+1 say ztelefone pict "@!"
                        @ prow(),70    say "Pag."
                        @ prow(),75    say strzero(wpag,3,0) pict "999"
                        @ prow()+1,02  say "Extrato de Conta Cliente  -  a Prazo"
                        @ prow()+1,02     say "Cliente...:"
                        @ prow(),pcol()+1   say wcodigo_cli pict "99999"
                        @ prow(),pcol()+2   say wnome       pict "@!"
                        @ prow()+1,02     say "Endereco..:"
                        @ prow(),pcol()+1 say wendereco pict"@!"
                        @ prow()+1,02     say "Bairro....:"
                        @ prow(),pcol()+1 say wbairro pict "@!"
                        @ prow(),pcol()+1 say "CEP.:"
                        @ prow(),pcol()+1 say wcep pict "@!"
                        @ prow()+1,02     say "Cidade....:"
                        @ prow(),pcol()+1 say wcidade pict "@!"
                        @ prow(),pcol()+1 say "UF..:"
                        @ prow(),pcol()+1 say wuf pict "@!"
                        @ prow()+1,02     say "Telefone..:"
                        @ prow(),pcol()+1 say wtelefone pict "@!"
                        @ prow(),pcol()+1 say "Celular...:"
                        @ prow(),pcol()+1 say wcelular pict "@!"
                        @ prow()+1,02  say "Data Debito :"
                        @ prow(),15    say wdata_deb   pict "99/99/99"
                        @ prow()+2,02  say "Codigo  Especificacao                   Data       Quantidade   Valor       "
                        @ prow()+1,02  say "~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~ ~~~~~~~~~~~  ~~~~~~~~~~~~"
                        linh = 13
                        do while it < 10 .and. !eof()
                            iguala()
                            @ prow()+1,02 say wcodigo_est pict "999999"
                            select 1
                            seek wcodigo_est
                            iguala()
                            @ prow(),10   say wdescricao  pict "@s30"
                            @ prow(),42   say wdata       pict "@d"
                            @ prow(),54   say wquantidade pict "@e@z 99,999.99"
                            wprec = wvalor
                            compras = (wprec * wquantidade)
                            @ prow(),66   say compras    pict "999,999.99"
                            wlin = 1
                            if alltrim(wdescricao,1,5) != "CREDI"
                                wtotal = wtotal + compras
                            endif
                            select 3
                            skip
                            it = it + 1
                            linh = linh + 1
                        enddo
                        wpag = wpag + 1
                        if !eof()
                            @ prow()+1,02  say "----------------------------------------------------------------------------"
                            @ prow()+1,02  say "                                                                 continua..."
                            linh = 32-linh
                            @ prow() + linh, 0 say ""
                        else
                            @ prow()+1,02  say "----------------------------------------------------------------------------"
                            @ prow()+1,03  say " Total de Compas..................................."
                            @ prow(),62    say wtotal pict "9,999,999,999.99"
                            @ prow()+1,03  say " Total de Credito................................."
                            @ prow(),62    say wcredito pict "9,999,999,999.99"
                            @ prow()+1,03  say " Total a Pagar...................................."
                            @ prow(),62    say wtotal-wcredito pict "9,999,999,999.99"
                            @ prow()+2,02  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
                            @ prow()+1,02  say "              Sistema de Administracao Estoque - Versao 1.0"
                            @ prow(),03    say date() pict "@d"
                            @ prow(),69    say time() pict "99:99"
                            linh = 32 - (linh+7)
                            @ prow() + linh, 0 say ""
                        endif
                        it = 0
                        setprc(0,0)
                    enddo
                    set cons on
                    set devi to screen
                    set print off
                    set print to
                    lertexto(zdireto)
                    set color to n/bg
                    restore screen from tela12
                    loop
                case xop = 5
                    restore screen from tela_AN
                    set color to gr+/bg
                    dt = quadro(08,20,14,75)
                    @ 09,22 say "Vendedor:"
                    @ 10,22 say "Codigo..:"
                    @ 11,22 say "Debito..:"
                    @ 12,22 say "Credito.:"
                    @ 13,22 say "A Pagar.:"
                    set color to n/bg
                    @ 11,32 say wtotal-wcredito pict "@e@z 999,999.99"
                    do while .t.
                        set color to w/b
                        @ 23,05 say "<ESC> Para sair.<000> para Consulta."
                        set color to w,n/bg,,,n/w
                        @ 09,32 get wcodigo_ven pict "999" valid !empty(wcodigo_ven)
                        read
                        if readkey() = 12  .or. lastkey() = 27
                            close all
                            return
                        endif
                        if wcodigo_ven = "000"
                            select 7
                            save screen to telave
                            do pes_ven
                            restore screen from telave
                            iguala()
                        endif
                        set color to /w
                        @ 09,32 say wcodigo_ven pict "999"
                        select 7
                        set index to ind01.ven
                        seek wcodigo_ven
                        if !found()
                            do mensagem with "Desculpe ! Vendedor nao cadastrado...",8
                            loop
                        endif
                        iguala()
                        set color to r/w
                        @ 09,32 say wcodigo_ven  pict "@!"
                        @ 09,36 say wnome_ven pict "@s10"
                        ycodigo_ven = wcodigo_ven
                        ynome_ven   = wnome_ven
                        exit
                    enddo
                    do while .t.
                        wcodigo_est = "000000"
                        ycredito    = 0
                        set color to w+/b
                        @ 23,05 say "<ENTER> Para Consulta. <ESC> para Sair."
                        set color to n/bg
                        @ 10,32 get wcodigo_est pict "999999"
                        read
                        if readkey() = 12 .or. lastkey() = 27
                            close all
                            return
                        endif
                        if wcodigo_est = "000000"
                            select 1
                            save screen to tela_0
                            do pes_est
                            restore screen from tela_0
                            iguala()
                        endif
                        ycodigo_est = wcodigo_est
                        set color to /w
                        @ 10,32 say ycodigo_est pict "999999"
                        select 1
                        set index to ind01.est
                        seek ycodigo_est
                        if !found()
                            do mensagem with "Codigo nao cadastrado...",8
                            loop
                        endif
                        iguala()
                        set color to /w
                        @ 10,40 say wdescricao pict "@s35"
                        exit
                    enddo
                    if alltrim(wdescricao,1,5) = "CREDI"
                        set color to w,n/bg,,,n/w
                        @ 12,32 get ycredito pict "@e@z 999,999.99"
                        read
                    else
                        do mensagem with "Codigo de produto nao de credito..",8
                        loop
                    endif
                    set color to  /w
                    @ 13,31 say wtotal-(ycredito+wcredito) pict "@e@z 999,999.99"
                    p = "S"
                    limpa()
                    set color to *gr+/b
                    @ 23,02 say " Confirma Credito <S/N>? " get p pict "@!" valid p $ "SN"
                    read
                    limpa()
                    if upper(p) = "S"
                        select 2
                        set index to  ind01.cli,ind02.cli
                        seek ycodigo_cli
                        if found()
                            iguala()
                            if rec_lock()
                                replace credito    with wcredito+ycredito
                                unlock
                            endif
                        endif
                        select 5
                        set index to ind01.tic
                        go bott
                        iguala()
                        wticket    = strzero(val(wticket)+1,6,0)
                        wdata_ti   = date()
                        wtotal_ti  = ycredito
                        whistorico = ynome
                        wpagamento = "P"
                        wtipo_ti   = "R"
                        wtipo_tipo = "3"
                        wdesconto  = 0
                        wquantidade= ycredito
                        wdata      = date()
                        wcodigo_ven= ycodigo_ven
                        wvalor     = wpreco_ven
                        if add_rec(10)
                            revar()
                        endif
                        select 4
                        set index to ind01.sai
                        if add_rec(10)
                            revar()
                        endif
                        select 3
                        set index to ind01.pra
                        if add_rec(10)
                            revar()
                        endif
                    endif
                    exit
                    limpam()
                    loop
                other
                    loop
            endcase
            loop
        else
            limpa()
            lir = 9
            t = "S"
            total = 0
            wdata = date()
            wchaves = 1
            set color to n/bg
            @ lir,02 get wdata pict "@d"
            read
            if readkey() = 12 .or. lastkey() = 27
                saida = .t.
                return
            endif
            do while .t.
                wcodigo_est = "000000"
                wquantidade = 0
                set color to w+/b
                @ 23,05 say "<ENTER> Para Consulta. <ESC> para Sair."
                set color to n/bg
                @ lir,13 get wcodigo_est pict "999999"
                read
                if readkey() = 12 .or. lastkey() = 27
                    close all
                    return
                endif
                if wcodigo_est = "000000"
                    select 1
                    save screen to tela_0
                    do pes_est
                    restore screen from tela_0
                endif
                ycodigo_est = wcodigo_est
                set color to /w
                @ lir,13 say ycodigo_est pict "999999"
                select 1
                set index to ind01.est
                seek ycodigo_est
                if !found()
                    do mensagem with "Codigo nao cadastrado...",8
                    loop
                endif
                iguala()
                limpa()
                set color to  /w
                @ lir,20 say wdescricao    pict "!!!!!!!!!!!!!!!!!!!!"
                set color to w,n/bg,,,n/w
                @ lir,42 get wquantidade   pict "@e@z 9999.99"
                read
                if readkey() = 12 .or. lastkey() = 27
                    saida = .t.
                    return
                endif
                if wc_saldo = "S"
                    if wsaldo < wquantidade
                        do mensagem with "Quantidade Solicitada Maior que a Disponivel ! ",8,1
                        limpam()
                        loop
                    endif
                endif
                wprec  = wvalor
                wpagar = (wprec * wquantidade)
                total  = total + wpagar
                set color to /w
                @ lir,50 say wpagar         pict "@e@z 9,999.99"
                @ 18,50  say total          pict "@e@z 9,999.99"
                p = "S"
                limpa()
                @ 23,03 say "Confirma Compras <S/N>? " get p pict "@!" valid p $ "SN"
                read
                limpa()
                if upper(p) = "S"
                    wnf         = "CLIEN"+wcodigo_cli
                    wpreco      = wvalor
                    wtipo_ven   = "P"
                    zcodigo_est = wcodigo_est
                    select 1
                    set index to ind01.est,ind02.est
                    seek wcodigo_est
                    iguala()
                    if wc_saldo = "S"
                        wsaldo  = wsaldo  - wquantidade
                        if wsaldo < 0
                            wsaldo = 0
                        endif
                        if rec_lock()
                            replace saldo with wsaldo
                            unlock
                        endif
                    endif
                    select 4
                    set index to ind01.sai
                    if add_rec(10)
                        revar()
                    endif
                    wcodigo_est = zcodigo_est
                    select 2
                    set index to ind01.cli,ind02.cli
                    seek wcodigo_cli
                    if rec_lock()
                        replace debito   with total
                        replace data_deb with date()
                        unlock
                    endif
                    select 3
                    set index to ind01.pra
                    if add_rec(10)
                        revar()
                    endif
                endif
                exit
            enddo
        endif
    enddo
enddo
close all
return

procedure pescon_2
para sele
select &sele
go top
sai=.f.
lp=9
tk=0
ul=22
set color to w/b
@ 23,03 say "Sobe :"+chr(24)+" Desce :"+chr(25)+" <ENTER> Confirma  <ESC> Sai "
set color to /w
do while .not. sai
    l=9   && primeira linha da tela l = 8
    do while l<=16 .and. .not. eof()
        do dados5c with l
        skip
        l=l+1
    enddo
    nl=l
    do while nl<=16
        nl=nl+1
    enddo
    l=l-1
    ul=l
    if tk<0 .or. tk=22
        goto rec
        lp=8
    else
        skip lp-(l+1)
    endif
    pl=.f.
    if l<22
        if lp>9
            lp=l
            go bottom
        endif
        pl=.t.
    endif
    set color to w+/r
    do dados5c with lp
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
                    do dados5c with lp
                    lp=lp+1
                    skip
                    set color to w+/r
                    do dados5c with lp
                    set color to /w
                    loop
                else
                    set color to /w
                    do dados5c with lp
                    skip
                    if .not. eof()
                        scroll(9,2,16,59,1)
                        @ lp,01 say "³          ³      ³                     ³       ³         ³"
                        @ 16,01 say "³          ³      ³                     ³       ³         ³"
                    endif
                    set color to w+/r
                    do dados5c with lp
                    set color to /w
                    loop
                endif
            case tk=5
                if bof()
                    do mensagem with "Inicio do Arquivo !",4
                    set color to /w
                    loop
                endif
                if lp>9
                    if eof()
                        skip -1
                    endif
                    do dados5c with lp
                    lp=lp-1
                    skip -1
                    set color to w+/r
                    do dados5c with lp
                    set color to /w
                    loop
                else
                    set color to /w
                    if eof()
                        skip -1
                    endif
                    do dados5c with lp
                    skip -1
                    if .not. bof()
                        scroll(9,2,16,59,-1)
                        @ lp,01 say "³          ³      ³                     ³       ³         ³"
                        @ 16,01 say "³          ³      ³                     ³       ³         ³"
                    endif
                    set color to w+/r
                    do dados5c with lp
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
procedure dados5c
    * Apresentacao de uma linha de dados para consulta na tela
    **********************************************************
parameters linha,valor
iguala()
@ linha,02  say wdata       pict "@D"
@ linha,13  say wcodigo_est pict "999999"
wcodigo_est = codigo_est
select 1
seek wcodigo_est
iguala()
@ linha,20  say wdescricao pict "!!!!!!!!!!!!!!!!!!!!"
@ linha,42  say wquantidade pict "@e@z 9999.99"
wtotal_urv = (wquantidade*wvalor)
@ linha,50  say wtotal_urv pict "@e@z 9,999.99"
select 3
return
    *******procedure limpa area de trabalho*********************
procedure limpam
set color to /w
@ 09,01 say "³          ³      ³                     ³       ³         ³"
@ 10,01 say "³          ³      ³                     ³       ³         ³"
@ 11,01 say "³          ³      ³                     ³       ³         ³"
@ 12,01 say "³          ³      ³                     ³       ³         ³"
@ 13,01 say "³          ³      ³                     ³       ³         ³"
@ 14,01 say "³          ³      ³                     ³       ³         ³"
@ 15,01 say "³          ³      ³                     ³       ³         ³"
@ 16,01 say "³          ³      ³                     ³       ³         ³"
lir = 9
return

procedure telapra
set color to /w
@ 03,01 say "ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿"
@ 04,01 say "³Codigo:                                                  ³"
@ 05,01 say "³Nome..:                                                  ³"
@ 06,01 say "ÃÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄ´"
@ 07,01 say "³ Data     ³ Cod. ³     Descricao       ³ Quant ³  Valor  ³"
@ 08,01 say "ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄ´"
@ 09,01 say "³          ³      ³                     ³       ³         ³"
@ 10,01 say "³          ³      ³                     ³       ³         ³"
@ 11,01 say "³          ³      ³                     ³       ³         ³"
@ 12,01 say "³          ³      ³                     ³       ³         ³"
@ 13,01 say "³          ³      ³                     ³       ³         ³"
@ 14,01 say "³          ³      ³                     ³       ³         ³"
@ 15,01 say "³          ³      ³                     ³       ³         ³"
@ 16,01 say "³          ³      ³                     ³       ³         ³"
@ 17,01 say "ÃÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄ´"
@ 18,01 say "³                                                         ³"
@ 19,01 say "³                                                         ³"
@ 20,01 say "³                                                         ³"
@ 21,01 say "³                                                         ³"
@ 22,01 say "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ"



set color to n/w
dt = quadro(03,60,08,77)
set color to w/r
@ 03,60 say "     Debito       "
set color to  /w
@ 04,61 say "Data:          "
@ 05,61 say "Valor..:        "
@ 06,61 say "Credito:        "
@ 07,61 say "A Pagar:        "
return
    *****************************************************************************
    *                          F   I   M                                        *
    *****************************************************************************
