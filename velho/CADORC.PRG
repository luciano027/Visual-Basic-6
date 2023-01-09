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
    select 6
    inicia()
    select 7
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
                    @ lir,10 say wdescricao      pict "@s19"
                    @ lir,30 say wquantidade     pict "@e@z 9999.99"
                    @ lir,39 say wPreco_ven      pict "@e@z 999.99"
                    wcompra = (wPreco_ven * wquantidade)
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
            Janela2("W",.F.,12,55,21,72,.F.)
            SetColor("N/W,W+/R")
            @ 13,56   prompt "    RETORNO    "
            @ 14,56   prompt "Compras        "
            @ 15,56   prompt "Cancelar Item  "
            @ 16,56   prompt "Dinheiro/Troco "
            @ 17,56   prompt "Fecha Orcamento"
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
                    @ lir,39 say wPreco_ven    pict "@e@z 999.99"
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
                        wdata       = wdata_ti
                        witem       = strzero(val(yitem)+1,3,0)
                        wnf         = "O"+zticket
                        wcs         = "C"
                        wdata       = ydata_ti
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
                    do pescsai1 with 2
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
                                endif
                                select 2
                                skip
                            enddo
                            do mensagem with "Registro Deletado...",8
                        endif
                        limpas()
                        wtotal = wtotal - (wquantidade * wPreco_ven)
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
                        close all
                        return
                    endif
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
                            wnf = "O"+zticket
                            select 2
                            skip
                        enddo
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
                    save screen to tela12
                    select 2
                    index on ticket to indext.vis for ticket = yticket
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
                        @ prow()+1,02     say zempresa  pict "@!"
                        @ prow()+1,02     say zendereco pict "@!"
                        @ prow()+1,02     say "Telefone: "
                        @ prow(),pcol()+1 say ztelefone pict "@!"
                        @ prow(),pcol()+1 say "FAX: "
                        @ prow(),pcol()+1 say ztelefone pict "@!"
                        @ prow(),70       say "Pag."
                        @ prow(),75       say strzero(wpag,3,0) pict "999"
                        @ prow()+1,02     say "Orcamento"
                        @ prow()+1,02     say "Data..:"
                        @ prow(),pcol()+1 say wdata_ti pict "@d"
                        @ prow()+2,02  say "Codigo  Especificacao                            Quant    P.Venda Total      "
                        @ prow()+1,02  say "~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~ ~~~~~~~ ~~~~~~~~~~ "
                        do while prow()<45 .and. !eof()
                            iguala()
                            @ prow()+1,02 say wcodigo_est pict "999999"
                            select 1
                            seek wcodigo_est
                            iguala()
                            @ prow(),10   say wdescricao             pict "@s27"
                            @ prow(),51   say wquantidade            pict "@e@z 99999.99"
                            @ prow(),60   say wpreco_ven             pict "@e@z 9999.99"
                            compras = (wpreco_ven * wquantidad)
                            @ prow(),68   say compras                pict "@e@z 99,999.99"
                            wlin = 1
                            geral = geral + compras
                            select 2
                            skip
                            esc = inkey()
                            if esc = 12
                                @ prow()+2,40 say "***********  Impressao Cancelada..."
                                close all
                                return
                            endif
                        enddo
                        if !eof()
                            eject
                            wpag = wpag + 1
                        endif
                    enddo
                    reg = recno()
                    @ prow()+1,02   say "----------------------------------------------------------------------------"
                    @ prow()+1,02   say " Total de Compas..:"
                    @ prow(),pcol() say geral pict "@e@z 999,999.99"
                    @ prow(),pcol() say " Desconto..:"
                    @ prow(),pcol() say wdesconto pict "@e@z 999,999.99"
                    @ prow(),pcol() say " Total a Pagar..:"
                    @ prow(),pcol() say geral - wdesconto pict "@e@z 99,999.99"
                    @ prow()+2,02   say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
                    @ prow()+1,02  say "              Sistema de Administracao Estoque - Versao 1.0"
                    @ prow(),02    say date() pict "@d"
                    @ prow(),69    say time() pict "99:99"
                    eject
                    set cons on
                    set devi to screen
                    set print off
                    set print to
                    lertexto(zdireto)
                    restore screen from tela12
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
            @ lir,10 say wdescricao  pict "@s19"
            @ lir,39 say wPreco_ven    pict "@e@z 999.99"
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
            wnf         = "O"+zticket
            wcs         = "C"
            s1          = 0
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

                                                                             *                          F   I   M                                        *
                                                                             *****************************************************************************

