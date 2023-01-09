*****************************************************************************
* Programa...: CADSAI1.PRG                                                  *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 02/12/98          Data da Ultima Atualizacao:                *
* Objetivo...: VENDAS CAIXA NAO FISCAL                                      *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************
select 1
if !net_use("estoque",.f.,"ind01.est","ind02.est")
    return
endif
select 2
if !net_use("saida",.f.,"ind01.sai")
    return
endif
select 3
if !net_use("ticket",.f.,"ind01.tic")
    return
endif
select 4
if !net_use("vendedor",.f.,"ind01.ven","ind02.ven")
    return
endif
select 5
if !net_use("clientes",.f.,"ind01.cli","ind02.cli","ind03.cli")
    return
endif
select 6
if !net_use("caixa",.f.,"ind01.cai")
    return
endif
select 7
if !net_use("receber",.f.,"ind01.rec")
    return
endif
select 8
if !net_use("prazo",.f.,"ind01.pra")
    return
endif

chave = 0
do while .t.
    select 1
    inicia()
    select 2
    inicia()
    select 5
    inicia()
    select 8
    inicia()
    select 4
    inicia()
    select 3
    set index to ind01.tic
    go bott
    iguala()
    yticket = strzero(val(wticket)+1,6,0)
    inicia()
    select 3
    inicia()
    wticket = yticket
    do tela3
    set color to r/w
    save screen to telafor
    @ 05,65 say wticket pict "999999"
    zticket = wticket
    do while .t.
        select 3
        set index to ind01.tic
        seek zticket
        if found()
            limpa()
            iguala()
            set color to r/w
            @ 06,65 say wdata_ti pict "@d"
            @ 08,65 say wdinheiro pict "@e@z 9,999.99"
            select 2
            set index to ind01.sai
            seek wticket
            if found()
                iguala()
                yticket     = wticket
                ydata_ti    = wdata_ti
                ycodigo_ven = wcodigo_ven
                lir = 7
                wtotal = 0
                do while wticket = yticket
                    set color to /w
                    @ lir,03 say wcodigo_est pict "999999"
                    select 1
                    set index to ind01.est
                    seek wcodigo_est
                    iguala()
                    select 2
                    @ lir,10 say wtitulo                pict "@s19"
                    @ lir,30 say wquantidade+wquant_con pict "@e@z 9,999"
                    @ lir,39 say wpreco_v               pict "@e@z 999.99"
                    wcompra = (wpreco_v * (wquantidade+wquant_con))
                    @ lir,46 say wcompra                pict "@e@z 9999.99"
                    lir = lir + 1
                    if lir > 14
                        inkey(.5)
                        do limpas
                    endif
                    wtotal = (wtotal + wcompra)  &&----> imprime o total
                    yitem  = witem
                    skip
                    iguala()
                enddo
                wticket = yticket
                ytotal  = wtotal
                wtotal  = wtotal - wdesconto
                if wdinheiro > 0
                    wtroco  = wdinheiro - wtotal
                    @ 10,65 say wtroco     pict "@e@z 9,999.99"
                endif
                vl1()
                NGrande( wtotal, 17, 03, 52, "@E 999,999.99" )
                set color to r/w
                @ 07,65 say wtotal     pict "@e@z 9,999.99"
                @ 08,65 say wdesconto  pict "@e@z 9,999.99"
            endif
            limpa()
            save screen to telaan
            Janela2("W",.F.,12,55,20,72,.F.)
            SetColor("N/W,W+/R")
            @ 13,56   prompt "    RETORNO    "
            @ 14,56   prompt "Compras        "
            @ 15,56   prompt "Cancelar Item  "
            @ 16,56   prompt "Dinheiro/Troco "
            @ 17,56   prompt "Pagamento      "
            @ 18,56   prompt "Cancelar Cupom "
            @ 19,56   prompt "Desconto       "
            xop = 5
            menu to xop
            do case
                case xop = 1
                    restore screen from telaAN
                    limpa()
                    if wpagamento = " "
                        do mensagem with " Pagamento nao Efetuado...",8
                        loop
                    endif
                    close all
                    return
                case xop = 2
                    restore screen from telaAN
                    if wpagamento !=" "
                        do mensagem with " Pagamento ja Efetuado...",8
                        loop
                    endif
                    limpa()
                    wcodigoest  = 0
                    wquantidade = 0
                    set color to w/b
                    @ 23,05 say "<ENTER> Para Consulta. <ESC> para Sair."
                    set color to n/bg
                    @ lir,03 get wcodigoest pict "999999"
                    read
                    if readkey() = 12 .or. lastkey() = 27
                        loop
                    endif
                    wcodigo_est = strzero(wcodigoest,6,0)
                    if wcodigo_est = "000000"
                        select 1
                        save screen to tela_0
                        do pes_est
                        restore screen from tela_0
                        iguala()
                    endif
                    set color to /w
                    @ lir,03 say wcodigo_est pict "999999"
                    select 1
                    set index to ind01.est
                    seek wcodigo_est
                    if !found()
                        do mensagem with "Codigo nao cadastrado...",8
                        limpas()
                        loop
                    endif
                    iguala()
                    set color to  /w
                    @ lir,10 say wtitulo     pict "@s19"
                    @ lir,39 say wpreco_v    pict "@e@z 999.99"
                    set color to w,n/bg,,,n/w
                    @ lir,30 get wquantidade pict "@e@z 9,999"
                    read
                    if wsaldo < wquantidade
                        do mensagem with "Quantidade Solicitada Maior que a Disponivel ! ",8,1
                        limpas()
                        loop
                    endif
                    wcompra = (wpreco_v * wquantidade)
                    set color to  /w
                    @ lir,46 say wcompra    pict "@e@z 9999.99"
                    p = "S"
                    limpa()
                    @ 23,03 say "Confirma Compras <S/N>? " get p pict "@!" valid p $ "SN"
                    read
                    limpa()
                    if upper(p) = "S"

                        select 1
                        set index to ind01.est
                        seek wcodigo_est
                        iguala()
                        wdata     = wdata_ti
                        witem     = strzero(val(yitem)+1,3,0)
                        wnf       = "V"+zticket
                        wcs        = "C"
                        ysaldo     = wsaldo
                        if rec_lock()
                            replace saldo with wsaldo - wquantidade
                            unlock
                        endif
                        wdata      = ydata_ti
                        wcodigo_ven = ycodigo_ven
                        select 2
                        set index to ind01.sai
                        if add_rec(10)
                            revar()
                        endif
                        store 0 to wquantidade
                        zitem = zitem + 1
                    endif
                    limpa()
                    loop
                case xop = 3
                    restore screen from telaAN
                    if wpagamento !=" "
                        do mensagem with " Pagamento Ja Efetuado...",8
                        loop
                    endif
                    save screen to tela09
                    select 2
                    index on ticket to indcon.vis for ticket = wticket
                    set index to indcon.vis
                    do pescon_66 with 2
                    reg = recno()
                    locate for reg = recno()
                    if Confirma ("Confirma Delecao do Item ?")=1    && gravacao dos dados
                        iguala()
                        qual= witem
                        wnf = "V"+zticket
                        select 2
                        set index to ind01.sai
                        seek wticket
                        if found()
                            do while !eof()
                                iguala()
                                if wcodigo_est=codigo_est .and. nf=wnf .and. item = qual
                                    if rec_lock()
                                        delete
                                    endif
                                    select 1
                                    set index to ind01.est,ind02.est
                                    seek wcodigo_est
                                    if found()
                                        iguala()
                                        wsaldo    = wsaldo    + wquantidade
                                        if rec_lock()
                                            revar()
                                            unlock
                                        endif
                                    endif
                                endif
                                select 2
                                skip
                            enddo
                            do mensagem with "Registro Deletado...",8
                        endif
                        limpas()
                        wtotal = wtotal - (wquantidade * wpreco_v)
                        if wtotal = 0
                            select 3
                            set index to ind01.tic
                            seek wticket
                            if found()
                                if rec_lock()
                                    replace pagamento with "O"
                                    replace historico with "Cupom Cancelado"
                                    unlock
                                endif
                            endif
                            do mensagem with "Cupom Cancelado...",8
                            close all
                            return
                        endif
                    endif
                    restore screen from tela09
                    limpas()
                    loop
                case xop = 4
                    restore screen from telaAN
                    if wpagamento !=" "
                        do mensagem with " Pagamento ja Efetuado...",8
                        loop
                    endif
                    wdesc = 0
                    save screen to teladesc
                    set color to w/b
                    @ 23,05 say "<ESC> para Sair"
                    dt = quadro(09,20,11,50)
                    @ 10,22 say "Dinheiro R$.:"
                    @ 10,36 get wdinheiro pict "@e@z 999.99" valid wdinheiro > wtotal
                    read
                    if readkey() = 12 .or. lastkey() = 27
                        restore screen from teladesc
                        loop
                    endif
                    p = "S"
                    limpa()
                    @ 23,03 say "Confirma Dados <S/N>? " get p pict "@!" valid p $ "SN"
                    read
                    restore screen from teladesc
                    limpa()
                    if upper(p) = "S"
                        wtroco  = wdinheiro - wtotal
                        if wtroco < 0
                            wtroco = 0
                        endif
                        set color to r/w
                        @ 09,65 say wdinheiro  pict "@e@z 9,999.99"
                        @ 10,65 say wtroco     pict "@e@z 9,999.99"
                        select 3
                        set index to ind01.tic
                        seek wticket
                        if rec_lock()
                            replace dinheiro with wdinheiro
                            unlock
                        endif
                    endif
                    loop
                case xop = 5
                    if wpagamento !=" "
                        do mensagem with " Pagamento Ja Efetuado...",8
                        loop
                    endif
                    save screen to telaxop5
                    save screen to telaop3
                    quadro(14,05,17,77)
                    @ 15,06 say "Total da Venda:"
                    @ 16,06 say "Entrada..:"
                    @ 15,22 say wtotal pict "@e@z 999,999.99"
                    @ 16,17 say wtotal pict "@e@z 999,999.99"
                    do sinal with "Forma de Pagamento    "
                    set color to w/b
                    @ 23,05 say "<ESC> Para sair."
                    do while .t.
                        set color to w/b
                        @ 23,05 say "<ESC> Para sair."
                        SetColor("N/W,W+/R")
                        @ 19,08      prompt "Dinheiro "
                        @ 19,col()+1 prompt "Prazo    "
                        @ 19,col()+1 prompt "Cliente  "
                        save screen to telarel1
                        menu to op2
                        do case
                            case op2 = 1
                                wtipo_tipo = "1"
                                p = "S"
                                limpa()
                                @ 23,03 say "Confirma Dados <S/N> ?" get p pict "@!" valid p $ "SN"
                                read
                                limpa()
                                if upper(p) = "S"
                                    select 3
                                    set index to ind01.tic
                                    seek wticket
                                    if found()
                                        if rec_lock()
                                            replace pagamento with "S"
                                            replace total_ti  with wtotal
                                            replace historico with "a Vista-Dinheiro "
                                            replace tipo_ti   with "D"
                                            replace tipo_tipo with wtipo_tipo
                                            unlock
                                        endif
                                    endif
                                    select 6
                                    set index to ind01.cai
                                    seek dtos(wcaixa)+wtipo_caixa
                                    if found()
                                        iguala()
                                        if rec_lock()
                                            replace caixa      with wdata_ti
                                            replace tipo_caixa with "C"
                                            replace dinheiro   with wdinheiro + wtotal
                                            unlock
                                        endif
                                    endif
                                else
                                    exit
                                endif
                            case op2 = 2
                                wtipo_tipo = "4"
                                set color to w,n/bg,,,n/w
                                @ 16,16  get wcartao_T   pict "@s18"
                                @ 16,50  get wcartao_N   pict "@s10"
                                @ 17,16  get wbanco      pict "@!"
                                @ 17,33  get wagencia    pict "@!"
                                @ 17,45  get wconta      pict "@!"
                                @ 17,65  get wcheque     pict "@!"
                                @ 18,17  get wtotal      pict "@e@z 999,999.99"
                                read
                                p = "S"
                                limpa()
                                @ 23,03 say "Confirma Dados <S/N> ?" get p pict "@!" valid p $ "SN"
                                read
                                limpa()
                                if upper(p) = "S"
                                    select 3
                                    set index to ind01.tic
                                    seek wticket
                                    if found()
                                        if rec_lock()
                                            replace pagamento with "S"
                                            replace total_ti  with wtotal
                                            replace historico with "a Prazo-30d    "
                                            replace tipo_ti   with "D"
                                            replace tipo_tipo with wtipo_tipo
                                            replace cartao_n  with wcartao_n
                                            replace cartao_t  with wcartao_t
                                            replace banco     with wbanco
                                            replace agencia   with wagencia
                                            replace conta     with wconta
                                            replace cheque    with wcheque
                                            unlock
                                        endif
                                    endif
                                    select 6
                                    set index to ind01.cai
                                    seek dtos(wcaixa)+wtipo_caixa
                                    if found()
                                        iguala()
                                        if rec_lock()
                                            replace caixa      with wdata_ti
                                            replace tipo_caixa with "C"
                                            if !empty(wcartao_t)
                                                if rec_lock()
                                                    replace cartao     with wcartao + wtotal
                                                    unlock
                                                endif
                                            endif
                                            if !empty(wagencia)
                                                if rec_lock()
                                                    replace cheque     with wcheque + wtotal
                                                    unlock
                                                endif
                                            endif
                                            if empty(wcartao_t) .and. empty(wagencia)
                                                if rec_lock()
                                                    replace dinheiro  with wdinheiro + wtotal
                                                    unlock
                                                endif
                                            endif
                                            unlock
                                        endif
                                    endif
                                    do while .t.
                                        select 7
                                        set index to ind01.chq
                                        go bott
                                        iguala()
                                        ycod_cheq = strzero(val(wcod_cheq)+1,5,0)
                                        inicia()
                                        wcod_cheq = ycod_cheq
                                        set color to n/w
                                        dt = quadro(08,21,18,76)
                                        set color to w/r
                                        @ 08,37 say "  CADASTRA CHEQUES PRE  "
                                        set color to /w
                                        @ 09,22 say "Documento...:"
                                        @ 11,22 say "N§ do Cheque:"
                                        @ 12,22 say "Vencimento..:"
                                        @ 13,22 say "Valor.......:"
                                        @ 14,22 say "Agencia.....:                        Banco.:"
                                        @ 15,22 say "Titular.....:"
                                        @ 16,22 say "Telefone....:"
                                        @ 17,22 say "Cliente.....:"
                                        limpa()
                                        @ 23,01 say space(78)
                                        @ 23,05 say "<ESC> Para sair."
                                        set color to r/w
                                        @ 09,36 say wcod_cheq  pict "99999"
                                        set color to w,n/bg,,,n/w
                                        @ 11,36 get wn_cheque   pict "@!"
                                        @ 12,36 get wdata_cq    pict "@d"
                                        @ 13,36 get wvalor_cq   pict "@e 99,999.99"
                                        @ 14,36 get wagencia    pict "@!"
                                        @ 14,66 get wbanco      pict "@!"
                                        @ 15,36 get wtitular    pict "@!"
                                        @ 16,36 get wtel_tit    pict "@!"
                                        read
                                        do while .t.
                                            set color to w/b
                                            @ 23,05 say "<ESC> Para sair. <00000> para Consulta"
                                            set color to w,n/bg,,,n/w
                                            @ 17,36 get wcodigo_cli pict "99999"
                                            read
                                            if readkey() = 12  .or. lastkey() = 27
                                                close all
                                                return
                                            endif
                                            if empty(wcodigo_cli) .or. wcodigo_cli = "00000"
                                                select 5
                                                save screen to t01tela
                                                do pes_cliv
                                                restore screen from t01tela
                                                iguala()
                                            endif
                                            set color to /w
                                            @ 17,36 say wcodigo_cli pict "99999"
                                            select 5
                                            set index to ind01.cli
                                            seek wcodigo_cli
                                            if !found()
                                                do mensagem with "Cliente nao Cadastrado...",8
                                                loop
                                            endif
                                            iguala()
                                            set color to r/w
                                            @ 17,36 say wcodigo_cli pict "@!"
                                            @ 17,42 say wnome       pict "@s30"
                                            exit
                                        enddo
                                        limpa()
                                        p = "S"
                                        set color to w/b
                                        @ 23,03 say "Confirma Dados <S/N> ?" get p pict "@!" valid p $ "SN"
                                        read
                                        limpa()
                                        if upper(p) = "S"
                                            select 7
                                            set index to ind01.chq
                                            if add_rec(10)
                                                revar()
                                            endif
                                            select 6
                                            set index to ind01.cai
                                            seek dtos(wdata_cq)+"R"
                                            if !found()
                                                if add_rec(10)
                                                    replace caixa      with wdata_cq
                                                    replace tipo_caixa with "R"
                                                    replace cheque_pre with wvalor_cq
                                                endif
                                            else
                                                iguala()
                                                if rec_lock()
                                                    replace cheque_pre with wvalor_cq+wcheque_pre
                                                    unlock
                                                endif
                                            endif
                                            select 3
                                            set index to ind01.tic
                                            seek wticket
                                            if found()
                                                if rec_lock()
                                                    replace d30 with wvalor_cq
                                                    unlock
                                                endif
                                            endif
                                        endif
                                        close all
                                        return
                                    enddo
                                endif
                                loop
                            other
                                restore screen from telaan
                                exit
                        endcase
                        restore screen from telaan
                        close all
                        return
                    enddo
                    restore screen from telaan
                    loop
                case xop = 6
                    restore screen from telaAN
                    if wpagamento !=" "
                        do mensagem with " Pagamento Ja Efetuado...",8
                        loop
                    endif
                    restore screen from telaAN
                    if Confirma ("Confirma Cancelamento do Cupom ?")=1    && gravacao dos dados
                        select 2
                        set index to ind01.sai
                        index on ticket to indvis.con for ticket = zticket
                        do while !eof()
                            iguala()
                            if rec_lock()
                                delete
                            endif
                            wnf = "V"+zticket
                            select 1
                            set index to ind01.est,ind02.est
                            seek wcodigo_est
                            if found()
                                iguala()
                                iguala()
                                wsaldo    = wsaldo    + wquantidade
                                if rec_lock()
                                    revar()
                                    unlock
                                endif
                            endif
                            select 2
                            skip
                        enddo
                        select 4
                        set index to ind01.tic
                        seek wticket
                        if found()
                            if rec_lock()
                                replace pagamento with "O"
                                replace historico with "Cupom Cancelado"
                                unlock
                            endif
                        endif
                        do mensagem with "Cupom Cancelado...",8
                        close all
                        return
                    endif
                    loop
                case xop = 7
                    restore screen from telaAN
                    if wpagamento !=" "
                        do mensagem with " Pagamento Ja Efetuado...",8
                        loop
                    endif
                    wdesc  = 0
                    @ 08,67 get wdesc pict "@e@z 999.99" valid wdesc >= 0.00 .and. wdesc <= wtotal
                    read
                    if readkey() = 12  .or. lastkey() = 27
                        loop
                    endif
                    wdesconto = wdesc
                    wtotal    = wtotal - wdesconto
                    select 3
                    set index to ind01.tic
                    seek wticket
                    if rec_lock()
                        replace desconto with wdesconto
                        unlock
                    endif
                    loop
                case xop = 8
                    restore screen from telaAN
                    if wpagamento !=" "
                        do mensagem with " Pagamento ja Efetuado...",8
                        loop
                    endif
                    save screen to teladesc
                    set color to /w
                    @ 23,05 say "<ESC> para Sair"
                    dt = quadro(09,20,11,75)
                    @ 10,22 say "Cliente..:"
                    select 2
                    set index to ind01.sai
                    seek zticket
                    if found()
                        iguala()
                        select 5
                        set index to ind01.cli
                        seek wcodigo_cli
                        iguala()
                        set color to r/w
                        @ 10,32 say wcodigo_cli pict "@!"
                        @ 10,39 say wnome       pict "@s25"
                    endif
                    do while .t.
                        set color to w/b
                        @ 23,05 say "<ESC> Para sair. <ENTER> para Consulta"
                        set color to w,n/bg,,,n/w
                        @ 10,32 get wcodigo_cli pict "99999"
                        read
                        if readkey() = 12  .or. lastkey() = 27
                            restore screen from teladesc
                            exit
                        endif
                        wcodigo_cli = strzero(val(wcodigo_cli),5,0)
                        if empty(wcodigo_cli) .or. wcodigo_cli = "00000"
                            select 5
                            save screen to t01tela
                            do pes_cliv
                            restore screen from t01tela
                            iguala()
                        endif
                        zcodigo_cli = wcodigo_cli
                        set color to /w
                        @ 10,32  say wcodigo_cli pict "99999"
                        select 5
                        set index to ind01.cli
                        seek wcodigo_cli
                        if !found()
                            do mensagem with "Cliente nao Cadastrado...",8
                            loop
                        endif
                        iguala()
                        select 2
                        seek wcodigo_ig
                        iguala()
                        set color to /w
                        @ 10,39 say wnome       pict "@s20"
                        exit
                    enddo
                    if Confirma ("Confirma Dados?")=1    && gravacao dos dados
                        select 2
                        set index to ind01.sai
                        index on ticket to indvis.con for ticket = zticket
                        do while !eof()
                            if rec_lock()
                                replace codigo_cli with zcodigo_cli
                                unlock
                            endif
                            skip
                        enddo
                    endif
                    restore screen from teladesc
                    loop
                other
                    loop
            endcase
            loop
        endif
        lir = 7
        wtotal   = 0
        wdata_ti = date()
        zitem    = 1
        set color to r/w
        @ 06,65 say wdata_ti pict "@d"
        if chave = 0
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
                    select 4
                    save screen to telave
                    do pes_ven
                    restore screen from telave
                    iguala()
                endif
                set color to /w
                @ 15,64 say wcodigo_ven pict "999"
                select 4
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
        else
            wcodigo_ven = ycodigo_ven
            wnome_ven   = ynome_ven
            set color to r/w
            @ 15,64 say wcodigo_ven  pict "@!"
            @ 16,64 say wnome_ven pict "@s10"
        endif
        do while .t.
            wcodigoest = 0
            wquantidade = 0
            set color to w/b
            @ 23,05 say "<ENTER> Para Consulta. <ESC> para Sair."
            set color to n/bg
            @ lir,03 get wcodigoest pict "999999"
            read
            if readkey() = 12 .or. lastkey() = 27
                if wtotal = 0
                    close all
                    return
                else
                    exit
                endif
            endif
            wcodigo_est = strzero(wcodigoest,6,0)
            if wcodigo_est = "000000"
                select 1
                save screen to tela_0
                do pes_est
                restore screen from tela_0
                iguala()
            endif
            set color to /w
            @ lir,03 say wcodigo_est pict "999999"
            select 1
            set index to ind01.est
            seek wcodigo_est
            if !found()
                do mensagem with "Codigo nao cadastrado...",8
                loop
            endif
            iguala()
            set color to  /w
            @ lir,10 say wtitulo     pict "@s19"
            @ lir,39 say wpreco_v    pict "@e@z 999.99"
            set color to w,n/bg,,,n/w
            @ lir,30 get wquantidade pict "@e@z 9,999" valid wquantidade > 0
            read
            if  wquantidade = 0
                do mensagem with "Quantidade Solicitada nao pode ser Zero! ",8,1
                loop
            endif
            if wsaldo < wquantidade
                do mensagem with "Quantidade Solicitada Maior que a Disponivel ! ",8,1
                loop
            endif
            wcompra = (wpreco_v * wquantidade)
            set color to  /w
            @ lir,46 say wcompra    pict "@e@z 9999.99"
            witem = strzero(zitem,3,0)
            wdata = wdata_ti
            select 1
            set index to ind01.est
            seek wcodigo_est
            iguala()
            wdata       = wdata_ti
            wcaixa      = wdata_ti
            wtipo_caixa = "C"
            witem       = strzero(zitem,3,0)
            wnf         = "V"+zticket
            wcs         = "C"
            s1          = 0
            ysaldo      = wsaldo
            if rec_lock()
                replace saldo with wsaldo - wquantidade
                unlock
            endif
            select 2
            set index to ind01.sai
            if add_rec(10)
                revar()
            endif
            store 0 to wquantidade
            zitem = zitem + 1
            select 3
            set index to ind01.tic
            seek wticket
            if found()
                if rec_lock()
                    revar()
                    unlock
                endif
            else
                if add_rec(10)
                    revar()
                endif
            endif
            select 6
            set index to ind01.cai
            seek dtos(wcaixa)+wtipo_caixa
            if !found()
                if add_rec(10)
                    replace caixa      with wcaixa
                    replace tipo_caixa with wtipo_caixa
                endif
            endif
            lir = lir + 1
            if lir > 14
                inkey(.5)
                limpas()
            endif
            wtotal = (wtotal + wcompra)  &&----> imprime o total
            vl1()
            NGrande( wtotal, 17,03, 52, "@E 999,999.99" )
            set color to r/w
            @ 07,65 say wtotal     pict "@e@z 9,999.99"
        enddo
    enddo
    loop
enddo
close all
return

procedure pescon_66
para sele
select &sele
go top
sai=.f.
lp=7
tk=0
ul=14
set color to w/B
@ 23,03 say "Sobe :"+chr(24)+" Desce :"+chr(25)+" <ENTER> Confirma  <ESC> Sai "
set color to /w
do while .not. sai
    l=7   && primeira linha da tela l = 8
    do while l<=14 .and. .not. eof()
        do dados66c with l
        skip
        l=l+1
    enddo
    nl=l
    do while nl<=14
        nl=nl+1
    enddo
    l=l-1
    ul=l
    if tk<0 .or. tk=14
        goto rec
        lp=7
    else
        skip lp-(l+1)
    endif
    pl=.f.
    if l<14
        if lp>7
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
                if lp<14
                    do dados66c with lp
                    lp=lp+1
                    skip
                    set color to w+/r
                    do dados66c with lp
                    set color to /w
                    loop
                else
                    set color to n/w
                    do dados66c with lp
                    skip
                    if .not. eof()
                        scroll(07,02,14,53,1)
                        @ lp,02 say "³      ³                   ³        ³      ³       ³"
                        @ 14,02 say "³      ³                   ³        ³      ³       ³"
                    endif
                    set color to w+/r
                    do dados66c with lp
                    set color to n/w
                    loop
                endif
            case tk=5
                if bof()
                    do mensagem with "Inicio do Arquivo !",4
                    set color to /w
                    loop
                endif
                if lp>7
                    if eof()
                        skip -1
                    endif
                    do dados66c with lp
                    lp=lp-1
                    skip -1
                    set color to w+/r
                    do dados66c with lp
                    set color to n/w
                    loop
                else
                    set color to n/w
                    if eof()
                        skip -1
                    endif
                    do dados66c with lp
                    skip -1
                    if .not. bof()
                        scroll(07,02,14,53,-1)
                        @ lp,02 say "³      ³                   ³        ³      ³       ³"
                        @ 14,02 say "³      ³                   ³        ³      ³       ³"
                    endif
                    set color to w+/r
                    do dados66c with lp
                    set color to n/w
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
@ linha,03 say wcodigo_est pict "999999"
select 1
set index to ind01.est
seek wcodigo_est
iguala()
@ linha,10 say wtitulo                 pict "@s19"
select 2
@ linha,30 say wquantidade+wquant_con  pict "@e@z 99999"
@ linha,39 say wpreco_v                pict "@e@z 999.99"
wcompra = (wpreco_v * (wquantidade+wquant_con))
@ linha,46 say wcompra    pict "@e@z  999.99"
select 2
return

 *****************************************************************************
 *                          F   I   M                                        *
 *****************************************************************************

