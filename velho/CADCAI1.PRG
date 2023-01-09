*****************************************************************************
* Programa...: CADCAI1.PRG                                                  *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 02/12/98          Data da Ultima Atualizacao:                *
* Objetivo...: RETIRADA DO ESTOQUE - COM ENTRADA NO CAIXA                   *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************

select 1
if !net_use("ticket",.f.,"ind01.tic")
    return
endif
select 2
if !net_use("caixa",.f.,"ind01.cai")
    return
endif
select 3
if !net_use("vendedor",.f.,"ind01.ven","ind02.ven")
    return
endif

do while .t.
    select 1
    set index to ind01.tic
    go bott
    iguala()
    yticket = strzero(val(wticket)+1,6,0)
    inicia()
    select 3
    inicia()
    wticket = yticket
    wdata_ti= date()
    set color to w/b
    @ 23,05 say "<ESC> Para sair."
    SetColor("W/W")
    quadro(06,21,12,76)
    Janela2("W",.F.,06,21,12,76,.F.)
    set color to w/r
    @ 06,37 say "  RETIRADA DO CAIXA "
    SetColor("N/W,W+/R")
    @ 07,23 say "Vendedor...:"
    @ 08,23 say "Ticket.....:"
    @ 09,23 say "Data.......:"
    @ 10,23 say "Historico..:"
    @ 11,23 say "Valor Pago.:"
    set color to r/w
    @ 08,36 say yticket  pict "@!"
    @ 09,36 say wdata_ti pict "@d"
    do while .t.
        set color to w/b
        @ 23,05 say "<ESC> Para sair.<000> para Consulta."
        set color to w,n/bg,,,n/w
        @ 07,34 get wcodigo_ven pict "999" valid !empty(wcodigo_ven)
        read
        if readkey() = 12  .or. lastkey() = 27
            close all
            return
        endif
        if wcodigo_ven = "000"
            select 3
            save screen to telave
            do pes_ven
            restore screen from telave
            iguala()
        endif
        set color to /w
        @ 07,34 say wcodigo_ven pict "999"
        select 3
        set index to ind01.ven
        seek wcodigo_ven
        if !found()
            do mensagem with "Desculpe ! Vendedor nao cadastrado...",8
            loop
        endif
        iguala()
        set color to r/w
        @ 07,34 say wcodigo_ven  pict "@!"
        @ 07,38 say wnome_ven pict "@s10"
        ycodigo_ven = wcodigo_ven
        ynome_ven   = wnome_ven
        exit
    enddo
    limpa()
    set color to w/b
    @ 23,05 say "<ESC> Retorna."
    set color to w,n/bg,,,n/w
    @ 10,36 get whistorico pict "@!"
    @ 11,36 get wtotal_ti  pict "@e@z 99,999.99"
    read
    if readkey() = 12  .or. lastkey() = 27
        close all
        return
    endif
    p = "S"
    limpa()
    set color to w/b
    @ 23,03 say "Confirma Dados <S/N> ?" get p pict "@!" valid p $ "SN"
    read
    limpa()
    if upper(p) = "S"
        wpagamento  = "D"
        wtipo_ti    = "T"
        wtipo_tipo  = "7"
        wcodigo_ven = ycodigo_ven
        select 1
        set index to ind01.tic
        if add_rec(10)
            revar()
        endif
        select 2
        set index to ind01.cai
        seek dtos(wdata_ti)+"A"
        if found()
            iguala()
            if rec_lock()
                replace dinheiro   with wdinheiro+wtotal_ti
                unlock
            endif
        else
            if add_rec(10)
                replace caixa      with wdata_ti
                replace tipo_caixa with "A"
                replace dinheiro   with wtotal_ti
            endif
        endif

    endif
    close all
    return
enddo
close all
return
*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
