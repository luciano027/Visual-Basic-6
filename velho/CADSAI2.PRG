*****************************************************************************
* Programa...: CADSAI1.PRG                                                  *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 02/12/98          Data da Ultima Atualizacao:                *
* Objetivo...: VENDAS CAIXA NAO FISCAL                                      *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************
save screen to telaantes
set color to w+/r
dt = quadro(11,17,13,60)
set color to *+w/r
@ 12,18 say "Aguarde... Procurando Impressora Fiscal"
tempo = 1
chave = 0
nPort := 0
do while nPort = 0
    if nPort = 0
        if COM_OPEN(1, 4000)
            COM_INIT(1, 9600, "N", 8 , 1)
            COM_RTS(1, .T.)
            COM_DTR(1, .T.)
            COM_SEND (1,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (1,29)     /* ESC */
            COM_SEND (1,5)
            COM_FLUSH(1)
            millisec(100)
            nCaract := COM_COUNT(1)
            if nCaract > 1
                rec_char = COM_READ (1, 1)
                rec_char = COM_READ (1, 1)
                if (rec_char = chr(13))
                    nPort := 1
                    veloc = '9.600'
                else
                    COM_CLOSE (1)
                endif
            else
                COM_CLOSE (1)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(2, 4000)
            COM_INIT(2, 9600, "N", 8 , 1)
            COM_RTS(2, .T.)
            COM_DTR(2, .T.)
            COM_SEND (2,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (2,29)     /* ESC */
            COM_SEND (2,5)
            COM_FLUSH(2)
            millisec(100)
            nCaract := COM_COUNT(2)
            if nCaract > 1
                rec_char = COM_READ (2, 1)
                rec_char = COM_READ (2, 1)
                if (rec_char = chr(13))
                    nPort := 2
                    veloc = '9.600'
                else
                    COM_CLOSE (2)
                endif
            else
                COM_CLOSE (2)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(1, 4000)
            COM_INIT(1, 600, "N", 8 , 1)
            COM_RTS(1, .T.)
            COM_DTR(1, .T.)
            COM_SEND (1,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (1,29)     /* ESC */
            COM_SEND (1,5)
            COM_FLUSH(1)
            millisec(400)
            nCaract := COM_COUNT(1)
            if nCaract > 1
                rec_char = COM_READ (1, 1)
                rec_char = COM_READ (1, 1)
                if (rec_char = chr(13))
                    nPort := 1
                    veloc = '  600'
                else
                    COM_CLOSE (1)
                endif
            else
                COM_CLOSE (1)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(1, 4000)
            COM_INIT(1,  300, "N", 8 , 1)
            COM_RTS(1, .T.)
            COM_DTR(1, .T.)
            COM_SEND (1,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (1,29)     /* ESC */
            COM_SEND (1,5)
            COM_FLUSH(1)
            millisec(400)
            nCaract := COM_COUNT(1)
            if nCaract > 1
                rec_char = COM_READ (1, 1)
                rec_char = COM_READ (1, 1)
                if (rec_char = chr(13))
                    nPort := 1
                    veloc = '  300'
                else
                    COM_CLOSE (1)
                endif
            else
                COM_CLOSE (1)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(2, 4000)
            COM_INIT(2, 600, "N", 8 , 1)
            COM_RTS(2, .T.)
            COM_DTR(2, .T.)
            COM_SEND (2,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (2,29)     /* ESC */
            COM_SEND (2,5)
            COM_FLUSH(2)
            millisec(400)
            nCaract := COM_COUNT(2)
            if nCaract > 1
                rec_char = COM_READ (2, 1)
                rec_char = COM_READ (2, 1)
                if (rec_char = chr(13))
                    nPort := 2
                    veloc = '  600'
                else
                    COM_CLOSE (2)
                endif
            else
                COM_CLOSE (2)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(2, 4000)
            COM_INIT(2, 300, "N", 8 , 1)
            COM_RTS(2, .T.)
            COM_DTR(2, .T.)
            COM_SEND (2,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (2,29)     /* ESC */
            COM_SEND (2,5)
            COM_FLUSH(2)
            millisec(400)
            nCaract := COM_COUNT(2)
            if nCaract > 1
                rec_char = COM_READ (2, 1)
                rec_char = COM_READ (2, 1)
                if (rec_char = chr(13))
                    nPort := 2
                    veloc = '  300'
                else
                    COM_CLOSE (2)
                endif
            else
                COM_CLOSE (2)
            endif
        endif
    endif
    if nPort = 0
        if tempo > 3
            restore screen from telaantes
            save screen to porta
            Janela2("W",.F.,11,16,14,62,.F.)
            SetColor("N/W,W+/R")
            @ 12,18 say 'NŽO FOI POSSIVEL ENCONTRAR A SUA IMPRESSORA'
            @ 13,23 say 'DIGITE ALGO PARA REPETIR A BUSCA'
            inkey(1)
            COM_CLOSE (nPort)
            keyboard chr(027)
            close all
            return
            restore screen from porta
        endif
        tempo = tempo + 1
    endif
enddo
millisec(200)
COM_FLUSH(nPort)
restore screen from telaantes


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

primeiro=.t.
sair    =.f.
total   =0.00
contite =0

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
                    @ lir,38 say wPreco_ven      pict "@e@z 999.999"
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
                        if wtotal = 0
                            COM_CLOSE (nPort)
                            close all
                            return
                        else
                            loop
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
                    if lastkey()<>27
                        if primeiro
                            enviacodigo(1) &&-------------- Abre Cupom Fiscal
                            primeiro=.f.
                        endif
                        ws  = walicota
                        wc  = strzero(val(wcodigo_est),13,0)
                        wr  = "000"
                        wd  = "0"
                        we  = "0000"
                        wp  = strtran(strzero(wpreco_ven,10,3),'.','')
                        wq  = strtran(strzero(wquantidade,6,2),'.',',')
                        wu  = wunidade
                        wcr = subst(wdescricao,1,30)
                        if wquantidade < 1000
                            par_fiscal= ws+wc+wr+wd+we+wp+wq+wu+wcr+chr(255)
                        else
                            wq = "0" + strtran(strzero(int(wquantidade),5,0),'.','')
                            par_fiscal= ws+wc+wr+wd+we+wp+wq+wu+wcr+chr(255)

                        endif
                        enviacodigo(3) &&-------------- Produto Com Descricao em duas linhas
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
                    index on ticket to &zmicro for ticket = wticket
                    set index to &zmicro
                    do pescsai1 with 2
                    reg = recno()
                    locate for reg = recno()
                    if Confirma ("Confirma Delecao do Item ?")=1    && gravacao dos dados
                        iguala()
                        qual= witem
                        wnf = "V"+zticket
                        if val(qual)>0
                            par_fiscal= strzero(val(qual),3)
                            enviacodigo(4)  &&* Cancelamento de Item do Cupom Sendo Confeccionado
                        endif
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
                                            unlock
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
                            seek wticket
                            if found()
                                if rec_lock()
                                    replace pagamento with "O"
                                    replace historico with "Cupom Cancelado"
                                    unlock
                                endif
                            endif
                            do mensagem with "Cupom Cancelado...",8
                            enviacodigo(5) &&-------------- Cancela Cupom Geral
                            saida()
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
                    want  = wtotal
                    wclie = space(42)
                    wende = space(42)
                    save screen to telaxop5
                    save screen to telaop3
                    quadro(14,05,19,67)
                    @ 15,06 say "Total da Venda:"
                    @ 16,06 say "Cliente.......:"
                    @ 17,06 say "Endereco......:"
                    @ 15,22 say wtotal pict "@e@z 999,999.99"
                    do sinal with "Forma de Pagamento    "
                    set color to w/b
                    @ 23,05 say "<ESC> Para sair."
                    do while .t.
                        set color to w/b
                        @ 23,05 say "<ESC> Para sair."
                        SetColor("N/W,W+/R")
                        @ 18,08      prompt "Dinheiro "
                        @ 18,col()+2 prompt "Cliente  "
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
                                    select 2
                                    index on ticket to &zmicro for ticket = zticket
                                    set index to &zmicro
                                    do while !eof()
                                        if rec_lock()
                                            replace pago with "S"
                                            replace codigo_cli with "00001"
                                            unlock
                                        endif
                                        skip
                                    enddo

                                    par_fiscal = chr(27)+chr(241)+"1"+strtran(strzero(wdesconto,13,2),".","")
                                    tx_direto(par_fiscal)
                                    par_fiscal = chr(27)+chr(242)+"A"+strtran(strzero(wtotal,13,2),".","")+chr(255)
                                    tx_direto(par_fiscal)
                                    par_fiscal = chr(27)+chr(243)+chr(255)
                                    tx_direto(par_fiscal)
                                    saida()
                                    close all
                                    return
                                else
                                    exit
                                endif
                            case op2 = 2
                                wtipo_tipo = "1"
                                set color to w,n/bg,,,n/w
                                @ 16,22 get wclie  pict "@!"
                                @ 17,22 get wende  pict "@!"
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
                                    par_fiscal = chr(27)+chr(241)+"1"+strtran(strzero(wdesconto,13,2),".","")
                                    tx_direto(par_fiscal)
                                    par_fiscal = chr(27)+chr(242)+"A"+strtran(strzero(wtotal,13,2),".","")+chr(255)
                                    tx_direto(par_fiscal)
                                    par_fiscal = chr(27)+chr(243)+wclie+chr(13)+chr(10)+wende+chr(13)+chr(10)+chr(255)
                                    tx_direto(par_fiscal)
                                    saida()
                                    close all
                                    return
                                else
                                    exit
                                endif

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
                        index on ticket to &zmicro for ticket = zticket
                        set index to &zmicro,ind01.sai
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
                                        unlock
                                    endif
                                endif
                            endif
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
                        enviacodigo(5) &&-------------- Cancela Cupom Geral
                        do mensagem with "Cupom Cancelado...",8
                        saida()
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
                other
                    loop
            endcase
            loop
        endif
        select 3
        set index to ind01.tic
        seek zticket
        if !found()
            if add_rec(10)
                revar()
            endif
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
                    saida()
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
                    saida()
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
            if lastkey()<>27
                if primeiro
                    enviacodigo(1) &&-------------- Abre Cupom Fiscal
                    primeiro=.f.
                endif
                ws  = walicota
                wc  = strzero(val(wcodigo_est),13,0)
                wr  = "000"
                wd  = "0"
                we  = "0000"
                wp  = strtran(strzero(wpreco_ven,10,3),'.','')
                wq  = strtran(strzero(wquantidade,6,2),'.',',')
                wu  = wunidade
                wcr = subst(wdescricao,1,30)
                if wquantidade < 1000
                    par_fiscal= ws+wc+wr+wd+we+wp+wq+wu+wcr+chr(255)
                else
                    wq = "0" + strtran(strzero(int(wquantidade),5,0),'.','')
                    par_fiscal= ws+wc+wr+wd+we+wp+wq+wu+wcr+chr(255)
                endif
                enviacodigo(3) &&-------------- Produto Com Descricao em duas linhas
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

    * - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *
    *       Envia o codigo da opera‡„o a Impressora Fiscal
    *       Se o programa foi ativado com o parametro /RESPOSTA, apresenta o
    *       codigo enviado no rodape' da tela, junto com a resposta recebida
    *       da Impressora, e aguarda uma tecla antes de continuar

procedure enviacodigo
parameters comando
do case
    case comando=1
        say_fiscal (200,"" )            &&  * Abre Cupom Fiscal
    case comando=2
        say_fiscal (201,par_fiscal )   && * Fecha Cupom Fiscal (Desconto Por Porcentagem)
    case comando=3
        say_fiscal (225,par_fiscal)      && * Produto Com Descricao de Duas Linhas
    case comando=4
        say_fiscal (205,par_fiscal)      && * Cancela Item do Cupom Aberto
    case comando=5
        say_fiscal (206,"")              && * Cancela Ultimo Cupom emitido
    case comando=6
        say_fiscal (207,"")              && * Leitura-X
    case comando=7
        say_fiscal (208,par_fiscal)      && * Reducao-Z
    case comando=8
        say_fiscal (209,par_fiscal)      && * Leitura da Memoria Fiscal
    case comando=9
        say_fiscal (220,par_fiscal)      && * Grava Nova Aliquota Fiscal
    case comando=10
        say_fiscal (213,par_fiscal)      && * Fecha Cupom Fiscal (Desconto em Valor)
    case comando=11
        say_fiscal (202,par_fiscal)      && * Produto Com Descricao de una Linha
    case comando=12
        say_fiscal (225,par_fiscal)      && * Descricao do Produto em Forma Livre
endcase
return

    * - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *
procedure saida
COM_CLOSE (nPort)
return


    * - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *
    *       Envia uma string para a impressora fiscal sem aguardar resposta

function tx_direto
parameters parametro
COM_FLUSH(nPort)
COM_SEND (nPort,13)
nRest := COM_SEND (nPort, parametro)

DO WHILE nRest > 0
    parametro := RIGHT (parametro, nRest)
    nRest := COM_SEND (nPort, parametro)
ENDDO
return

    * - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - *
    *       Envia uma string para a impressora fiscal

function say_fiscal
parameters comando,parametro
COM_FLUSH(nPort)
COM_SEND (nPort,27)
COM_SEND (nPort,comando)
nRest := COM_SEND (nPort, parametro)

DO WHILE nRest > 0
    parametro := RIGHT (parametro, nRest)
    nRest     := COM_SEND (nPort, parametro)
ENDDO
return

    *****************************************************************************
    *                          F   I   M                                        *
    *****************************************************************************
                                                                             *****************************************************************************

