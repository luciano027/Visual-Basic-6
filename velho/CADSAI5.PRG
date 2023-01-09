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
select 9
if !net_use("juros",.f.,"ind01.jur")
    return
endif



chave = 0
wparc = space(02)
do while .t.
    select 1
    inicia()
    select 2
    inicia()
    select 5
    inicia()
    select 6
    inicia()
    select 7
    inicia()
    select 9
    inicia()
    select 4
    inicia()
    select 3
    set index to ind01.tic
    go bott
    iguala()
    yticket = strzero(val(wticket)+1,6,0)
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
            seek zticket
            if found()
                iguala()
                yticket     = zticket
                ydata_ti    = wdata_ti
                ycodigo_ven = wcodigo_ven
                lir         = 7
                wtotal      = 0
                do while wticket = zticket
                    set color to /w
                    @ lir,03 say wcodigo_est pict "999999"
                    select 1
                    set index to ind01.est
                    seek wcodigo_est
                    iguala()
                    select 2
                    @ lir,10 say wdescricao      pict "@s19"
                    @ lir,30 say wquantidade     pict "@e@z 9999.99"
                    @ lir,38 say wpreco_ven      pict "@e@z 999.999"
                    wcompra = (wPreco_ven * wquantidade)
                    @ lir,46 say wcompra                pict "@e@z 9999.99"
                    lir = lir + 1
                    if lir > 14
                        inkey(.5)
                        limpas()
                    endif
                    wtotal = (wtotal + wcompra)  &&----> imprime o total
                    yitem  = witem
                    skip
                    iguala()
                enddo
                wticket = zticket
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
            Janela2("W",.F.,12,55,21,72,.F.)
            SetColor("N/W,W+/R")
            @ 13,56   prompt "    RETORNO    "
            @ 14,56   prompt "Compras        "
            @ 15,56   prompt "Cancelar Item  "
            @ 16,56   prompt "Dinheiro/Troco "
            @ 17,56   prompt "Pagamento      "
            @ 18,56   prompt "Cancelar Cupom "
            @ 19,56   prompt "Desconto       "
            @ 20,56   prompt "Extrato        "
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
                    @ lir,10 say wdescricao  pict "@s19"
                    @ lir,38 say wPreco_ven    pict "@e@z 999.999"
                    set color to w,n/bg,,,n/w
                    @ lir,30 get wquantidade pict "@e@z 9999.99"
                    read
                    if wc_saldo = "S"
                        if wsaldo < wquantidade
                            do mensagem with "Quantidade Solicitada Maior que a Disponivel ! ",8,1
                            limpas()
                            loop
                        endif
                    endif
                    wcompra = (wPreco_ven * wquantidade)
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
                        if wc_saldo = "S"
                            if rec_lock()
                                replace saldo with wsaldo - wquantidade
                                unlock
                            endif
                        endif
                        wdata       = ydata_ti
                        wcodigo_ven = ycodigo_ven
                        wpreco      = wpreco_ven
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
                    index on ticket to &zmicro for ticket = zticket
                    set index to &zmicro
                    do pescsai1 with 2
                    reg = recno()
                    locate for reg = recno()
                    if Confirma ("Confirma Delecao do Item ?")=1    && gravacao dos dados
                        iguala()
                        qual= witem
                        wnf = "V"+zticket
                        select 2
                        set index to &zmicro
                        do while !eof()
                            iguala()
                            if wcodigo_est=codigo_est .and. wnf=nf .and. qual = item
                                if rec_lock()
                                    delete
                                endif
                                select 1
                                set index to ind01.est,ind02.est
                                seek wcodigo_est
                                if found()
                                    iguala()
                                    if wc_saldo = "S"
                                        wsaldo    = wsaldo    + wquantidade
                                        if rec_lock()
                                            revar()
                                        endif
                                    endif
                                endif
                            endif
                            select 2
                            skip
                        enddo
                        do mensagem with "Registro Deletado...",8
                        limpas()
                        wtotal = wtotal - (wquantidade * wPreco_ven)
                        if wtotal = 0
                            select 3
                            set index to ind01.tic
                            seek zticket
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
                        seek zticket
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
                    want = wtotal
                    save screen to telaxop5
                    save screen to telaop3
                    quadro(14,05,18,57)
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
                        @ 17,08      prompt "Dinheiro "
                        @ 17,col()+3 prompt "Cliente"
                        @ 17,col()+3 prompt "Financeira"
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
                                    seek zticket
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
                                    seek dtos(wdata_ti)+"C"
                                    if found()
                                        iguala()
                                        if rec_lock()
                                            replace dinheiro   with wdinheiro+wtotal
                                            unlock
                                        endif
                                    else
                                        if add_rec(10)
                                            replace caixa      with wdata_ti
                                            replace tipo_caixa with "C"
                                            replace dinheiro   with wdinheiro + wtotal
                                        endif
                                    endif
                                else
                                    exit
                                endif
                            case op2 = 2
                                wtipo_tipo = "4"
                                wcodigo_cli = "00000"
                                save screen to telaop23
                                set color to n/w
                                dt = quadro(15,40,17,78)
                                @ 16,42 say "Cliente:"
                                set color to w/b
                                @ 23,05 say "<ESC> Para sair. <00000> Pesquisa"
                                set color to w,n/bg,,,n/w
                                @ 16,51 get wcodigo_cli pict "99999"
                                read
                                if readkey() = 12  .or. lastkey() = 27
                                    restore screen from telaop23
                                    loop
                                endif
                                if empty(wcodigo_cli) .or. wcodigo_cli = "00000"
                                    select 5
                                    save screen to tela_d
                                    do pes_cli
                                    restore screen from tela_d
                                    iguala()
                                endif
                                set color to n/w
                                @ 16,51 say wcodigo_cli pict "99999"
                                select 5
                                set index to ind01.cli
                                seek wcodigo_cli
                                if !found()
                                    do mensagem with "Cliente nao cadastrado...",8
                                    restore screen from telaop23
                                    loop
                                endif
                                iguala()
                                if  wdata_deb != ctod("")
                                    if  date() > wdata_deb+30
                                        do mensagem with "Cliente com Debito...",8
                                        set color to n/w
                                        quadro(15,40,19,78)
                                        @ 16,42 say "Cliente.....:"
                                        @ 17,42 say "Data Debito.:"
                                        @ 18,42 say "Valor Debito:"
                                        set color to r/w
                                        @ 17,56 say wdata_deb   pict "@d"
                                        @ 18,56 say wdebito     pict "@e@z 99,999.99"
                                    endif
                                endif
                                set color to n/w
                                dt = quadro(15,40,19,78)
                                @ 16,42 say "Cliente:"
                                @ 17,42 say "Endereco"
                                @ 18,42 say "Telefone"
                                set color to r/w
                                @ 16,51 say wcodigo_cli pict "@!"
                                @ 16,58 say wnome       pict "@s20"
                                @ 17,51 say wendereco   pict "@s25"
                                @ 18,51 say wtelefone   pict "@!"
                                if Confirma ("Confirma Dados Cliente ?")=1    && gravacao dos dados
                                    if wdata_deb = ctod("")
                                        if rec_lock()
                                            replace data_deb with date()
                                            replace debito   with wtotal
                                            unlock
                                        endif
                                    endif
                                    zcodigo_cli = wcodigo_cli
                                    restore screen from telaop23
                                    select 3
                                    set index to ind01.tic
                                    seek zticket
                                    if found()
                                        if rec_lock()
                                            replace pagamento with "S"
                                            replace total_ti  with wtotal
                                            replace historico with wnome
                                            replace tipo_ti   with "D"
                                            replace tipo_tipo with "4"
                                            unlock
                                        endif
                                    endif
                                    select 2
                                    index on codigo_est to &zmicro for ticket = zticket
                                    set index to &zmicro
                                    do while !eof()
                                        iguala()
                                        select 8
                                        if add_rec(10)
                                            replace codigo_cli with zcodigo_cli
                                            replace codigo_est with wcodigo_est
                                            replace data       with wdata
                                            replace quantidade with wquantidade
                                            replace valor      with wpreco
                                            unlock
                                        endif
                                        select 2
                                        if rec_lock()
                                            replace codigo_cli with zcodigo_cli
                                            unlock
                                        endif
                                        skip
                                    enddo
                                    close all
                                    return
                                endif
                                restore screen from telaop23
                                loop
                            case op2 = 3
                                save screen to tela_fi
                                wparc = "00"
                                wpart = 0
                                set color to n/w
                                quadro(10,40,15,78)
                                @ 11,42 say "Financeira....:"
                                @ 12,42 say "Total Compra..:"
                                @ 13,42 say "Parcelas......:"
                                @ 14,42 say "Total Parcelas:"
                                do while .t.
                                    set color to w/b
                                    @ 23,05 say "<ESC> Para sair.<0> para consulta"
                                    set color to w,n/bg,,,n/w
                                    save screen to telafor
                                    @ 11,58 get wcodigo_ju pict "9"
                                    read
                                    if readkey() = 12 .or. lastkey() = 27
                                        exit
                                    endif
                                    if wcodigo_ju = "0"
                                        select 9
                                        save screen to tela_0
                                        do pes_ju
                                        restore screen from tela_0
                                    endif
                                    ycodigo_ju = wcodigo_ju
                                    set color to /w
                                    @ 11,58 say ycodigo_ju pict "9"
                                    select 9
                                    set index to ind01.jur
                                    seek ycodigo_ju
                                    if !found()
                                        do mensagem with "Financeira nao cadastrada...",8
                                        loop
                                    endif
                                    iguala()
                                    set color to /w
                                    @ 11,60 say wdesc_jur  pict "@s15"
                                    @ 12,58 say wtotal     pict "@e@z 999,999.99"
                                    set color to w,n/bg,,,n/w
                                    @ 13,58 get wparc      pict "99"
                                    read
                                    if wparc = "01"
                                        if wtotal >= wlimi_1
                                            wpart = (wtotal+wtaxa) * wfator_1
                                        else
                                            do mensagem with "Valor da compra e menor...",8
                                            loop
                                        endif
                                    endif
                                    if wparc = "02"
                                        if wtotal >= wlimi_2
                                            wpart = (wtotal+wtaxa) * wfator_2
                                        else
                                            do mensagem with "Valor da compra e menor...",8
                                            loop
                                        endif
                                    endif
                                    if wparc = "03"
                                        if wtotal >= wlimi_3
                                            wpart = (wtotal+wtaxa) * wfator_3
                                        else
                                            do mensagem with "Valor da compra e menor...",8
                                            loop
                                        endif
                                    endif
                                    if wparc = "04"
                                        if wtotal >= wlimi_4
                                            wpart = (wtotal+wtaxa) * wfator_4
                                        else
                                            do mensagem with "Valor da compra e menor...",8
                                            loop
                                        endif
                                    endif
                                    if wparc = "05"
                                        if wtotal >= wlimi_5
                                            wpart = (wtotal+wtaxa) * wfator_5
                                        else
                                            do mensagem with "Valor da compra e menor...",8
                                            loop
                                        endif
                                    endif
                                    if wparc = "06"
                                        if wtotal >= wlimi_6
                                            wpart = (wtotal+wtaxa) * wfator_6
                                        else
                                            do mensagem with "Valor da compra e menor...",8
                                            loop
                                        endif
                                    endif
                                    if wparc = "07"
                                        if wtotal >= wlimi_7
                                            wpart = (wtotal+wtaxa) * wfator_7
                                        else
                                            do mensagem with "Valor da compra e menor...",8
                                            loop
                                        endif
                                    endif
                                    if wparc = "08"
                                        if wtotal >= wlimi_8
                                            wpart = (wtotal+wtaxa) * wfator_8
                                        else
                                            do mensagem with "Valor da compra e menor...",8
                                            loop
                                        endif
                                    endif
                                    if wparc = "09"
                                        if wtotal >= wlimi_9
                                            wpart = (wtotal+wtaxa) * wfator_9
                                        else
                                            do mensagem with "Valor da compra e menor...",8
                                            loop
                                        endif
                                    endif
                                    if wparc = "10"
                                        if wtotal >= wlimi_10
                                            wpart = (wtotal+wtaxa) * wfator_10
                                        else
                                            do mensagem with "Valor da compra e menor...",8
                                            loop
                                        endif
                                    endif
                                    if wparc = "11"
                                        if wtotal >= wlimi_11
                                            wpart = (wtotal+wtaxa) * wfator_11
                                        else
                                            do mensagem with "Valor da compra e menor...",8
                                            loop
                                        endif
                                    endif
                                    if wparc = "12"
                                        if wtotal >= wlimi_12
                                            wpart = (wtotal+wtaxa) * wfator_12
                                        else
                                            do mensagem with "Valor da compra e menor...",8
                                            loop
                                        endif
                                    endif
                                    set color to /w
                                    @ 14,58 say wpart            pict "@e@z 999,999.99"
                                    if Confirma ("Imprime Extrato <S/N> ?")=1
                                        save screen to tela_a
                                        do tela_imp
                                        restore screen from tela_a
                                    endif
                                    if confirma("Confirma Cancelamento <S/N> ?")=1
                                        select 2
                                        set index to ind01.sai
                                        index on ticket to &zmicro for ticket = zticket
                                        set index to &zmicro
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
                                                if wc_saldo = "S"
                                                    wsaldo    = wsaldo    + wquantidade
                                                    if rec_lock()
                                                        revar()
                                                    endif
                                                endif
                                            endif
                                            select 2
                                            skip
                                        enddo
                                        select 3
                                        set index to ind01.tic
                                        seek zticket
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
                                enddo
                                restore screen from tela_fi
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
                        index on ticket to &zmicro for ticket = zticket
                        set index to &zmicro
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
                                if wc_saldo = "S"
                                    wsaldo    = wsaldo    + wquantidade
                                    if rec_lock()
                                        revar()
                                    endif
                                endif
                            endif
                            select 2
                            skip
                        enddo
                        select 3
                        set index to ind01.tic
                        seek zticket
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
                    seek zticket
                    if rec_lock()
                        replace desconto with wdesconto
                        unlock
                    endif
                    loop
                case xop = 8
                    save screen to tela12
                    do tela_imp
                    restore screen from tela12
                    loop
                other
                    loop
            endcase
            loop
        endif
        select 3
        set index to ind01.tic
        seek zticket
        if !found()
            wticket = zticket
            if add_rec(10)
                revar()
            endif
        else
            close all
            return
        endif
        close all
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
        select 9
        if !net_use("juros",.f.,"ind01.jur")
            return
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
            @ lir,10 say wdescricao  pict "@s19"
            @ lir,38 say wPreco_ven    pict "@e@z 999.999"
            set color to w,n/bg,,,n/w
            @ lir,30 get wquantidade pict "@e@z 9999.99" valid wquantidade > 0
            read
            if  wquantidade = 0
                do mensagem with "Quantidade Solicitada nao pode ser Zero! ",8,1
                loop
            endif
            if wc_saldo = "S"
                if wsaldo < wquantidade
                    do mensagem with "Quantidade Solicitada Maior que a Disponivel ! ",8,1
                    loop
                endif
            endif
            wcompra = (wPreco_ven * wquantidade)
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
            wpreco      = wpreco_ven
            if wc_saldo = "S"
                if rec_lock()
                    replace saldo with wsaldo - wquantidade
                    unlock
                endif
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
            seek zticket
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

procedure pescsai1
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
        do dadosai1 with l
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
    do dadosai1 with lp
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
                    do dadosai1 with lp
                    lp=lp+1
                    skip
                    set color to w+/r
                    do dadosai1 with lp
                    set color to /w
                    loop
                else
                    set color to n/w
                    do dadosai1 with lp
                    skip
                    if .not. eof()
                        scroll(07,02,14,53,1)
                        @ lp,02 say "³      ³                   ³        ³      ³       ³"
                        @ 14,02 say "³      ³                   ³        ³      ³       ³"
                    endif
                    set color to w+/r
                    do dadosai1 with lp
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
                    do dadosai1 with lp
                    lp=lp-1
                    skip -1
                    set color to w+/r
                    do dadosai1 with lp
                    set color to n/w
                    loop
                else
                    set color to n/w
                    if eof()
                        skip -1
                    endif
                    do dadosai1 with lp
                    skip -1
                    if .not. bof()
                        scroll(07,02,14,53,-1)
                        @ lp,02 say "³      ³                   ³        ³      ³       ³"
                        @ 14,02 say "³      ³                   ³        ³      ³       ³"
                    endif
                    set color to w+/r
                    do dadosai1 with lp
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
procedure dadosai1
                                                                                                            * Apresentacao de uma linha de dados para consulta na tela
                                                                                                            **********************************************************
parameters linha,valor
iguala()
@ linha,03 say wcodigo_est pict "999999"
select 1
set index to ind01.est
seek wcodigo_est
iguala()
@ linha,10 say wdescricao               pict "@s19"
select 2
@ linha,30 say wquantidade            pict "@e@z 9999.99"
@ linha,38 say wPreco_ven                pict "@e@z 999.999"
wcompra = (wPreco_ven * wquantidade)
@ linha,46 say wcompra    pict "@e@z 9999.99"
select 2
return


procedure tela_imp
select 2
index on ticket to &zmicro for ticket = zticket
set index to &zmicro
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
it = 1
do while inkey()#27 .and. !eof()
    @ prow()+1,02     say zempresa  pict "@!"
    @ prow()+1,02     say zendereco pict "@!"
    @ prow()+1,02     say "Telefone: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),pcol()+1 say "FAX: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),70       say "Pag."
    @ prow(),75       say strzero(wpag,3,0) pict "999"
    @ prow()+1,02     say "Extrato Cliente "
    @ prow()+1,02     say "Data..:"
    @ prow(),pcol()+1 say wdata_ti pict "@d"
    @ prow()+2,02  say "Codigo  Especificacao                            Quant    P.Venda Total      "
    @ prow()+1,02  say "~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~ ~~~~~~~ ~~~~~~~~~~ "
    do while it < 8 .and. !eof()
        iguala()
        @ prow()+1,02 say wcodigo_est pict "999999"
        select 1
        seek wcodigo_est
        iguala()
        @ prow(),10   say wdescricao             pict "@s39"
        @ prow(),51   say wquantidade            pict "@e@z 99999.99"
        @ prow(),60   say wpreco_ven             pict "@e@z 999.999"
        compras = (wpreco_ven * wquantidad)
        @ prow(),68   say compras                pict "@e@z 99,999.99"
        geral = geral + compras
        select 2
        skip
        it = it + 1
    enddo
    @ prow()+1,02   say "----------------------------------------------------------------------------"
    @ prow()+1,02   say " Total de Compas..:"
    @ prow(),pcol() say geral pict "@e@z 999,999.99"
    @ prow(),pcol() say " Desconto..:"
    @ prow(),pcol() say wdesconto pict "@e@z 999,999.99"
    @ prow(),pcol() say " Total a Pagar..:"
    @ prow(),pcol() say geral - wdesconto pict "@e@z 99,999.99"
    @ prow()+1,02   say "----------------------------------------------------------------------------"
    if !empty(wparc)
        @ prow()+1,02   say "Total de Compras..:"
        @ prow(),pcol()+1 say geral pict "@e@z 999,999.99"
        @ prow(),pcol()+1 say "Parcelas.:"
        @ prow(),pcol()+1 say wparc pict "@!"
        @ prow(),pcol()+1 say "Valor Parcela..:"
        @ prow(),pcol()+1 say wpart pict "@e@z 999,999.99"
    endif
    @ prow()+2,02   say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    @ prow()+1,02  say "              Sistema de Administracao Estoque - Versao 1.0"
    @ prow(),02    say date() pict "@d"
    @ prow(),69    say time() pict "99:99"
    if !eof()
       tsalto = zsalto + (8 - it)
       @ prow()+tsalto,01 say "."
       it = 0
       setprc(0,0)
    endif
enddo
set cons on
set devi to screen
set print off
set print to
lertexto(zdireto)
return
                                                                             *****************************************************************************
                                                                             *                          F   I   M                                        *
                                                                             *****************************************************************************

