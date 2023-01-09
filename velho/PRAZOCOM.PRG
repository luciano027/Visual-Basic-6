*****************************************************************************
* Programa...: PRAZOCOM.PRG                                                 *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao:                *
* Objetivo...: VENDAS A PRAZO - COMPRAS                                     *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************
do while .t.
    wcodigo_est = "000000"
    wquantidade = 0
    total       = wdebito
    wcodigo_cli = ycodigo_cli
    wdata       = date()
    set color to n/bg
    @ lir,02 get wdata   pict "@d"
    read
    set color to /w
    @ lir,02 say wdata pict "@d"
    if readkey() = 12 .or. lastkey() = 27
        return
    endif
    set color to w+/b
    @ 23,05 say "<ENTER> Para Consulta. <ESC> para Sair."
    set color to n/bg
    @ lir,13 get wcodigo_est pict "999999"
    read
    if wcodigo_est = "000000"
        select 1
        save screen to tela_0
        do pes_est
        restore screen from tela_0
        iguala()
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
    set color to /w
    @ lir,20 say wdescricao    pict "!!!!!!!!!!!!!!!!!!!!"
    set color to w,n/bg,,,n/w
    @ lir,42 get wquantidade   pict "@e@z 999.99"
    read
    if wc_saldo = "S"
        if wsaldo < wquantidade
            do mensagem with "Quantidade Solicitada Maior que a Disponivel ! ",8,1
            limpam()
            loop
        endif
    endif
    wprec = wpreco_ven
    wpagar = (wprec * wquantidade)
    total = total + wpagar
    set color to /w
    @ lir,50 say  wpagar       pict "@e@z 9,999.99"
    p = "S"
    set color to w/b
    limpa()
    set escap off
    @ 23,03 say "Confirma Compras <S/N>? " get p pict "@!" valid p $ "SN"
    read
    set escap on
    if upper(p) = "S"
        wnf         = "CLIEN"+wcodigo_cli
        wpreco      = wpreco_ven
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
            replace saldo with wsaldo
        endif
        wpreco = wpreco_ven
        select 4
        set index to ind01.sai
        append blank
        revar()
        wcodigo_est = zcodigo_est
        select 2
        set index to ind01.cli,ind02.cli
        seek wcodigo_cli
        replace debito   with total
        replace data_deb with ydata
        select 3
        set index to ind01.pra
        append blank
        revar()
        lir = lir + 1
        if lir > 16
            limpam()
        endif
    endif
    return
enddo
*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
