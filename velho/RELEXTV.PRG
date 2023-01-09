***************************************************************************
* Programa Tabela de Precos - Relatorio                                   *
* Emite relatorio geral (tabela de precos) indexado por espec_est         *
* Luciano 26/08/93                                                        *
***************************************************************************
save screen to telarel
select 1
if !net_use("estoque",.f.,"ind01.est","ind02.est")
    return
endif
select 2
if !net_use("csaida",.f.)
    return
endif
select 3
if !net_use("cticket",.f.)
    return
endif
select 4
if !net_use("vendedor",.f.,"ind01.ven","ind02.ven")
    return
endif
select 5
if !net_use("clientes",.f.,"ind01.cli","ind02.cli")
    return
endif

wticket = space(06)

do while .t.
    yticket = 0
    set color to n/w
    dt = quadro(19,40,21,68)
    @ 20,42 say "Ticket Nro....:"
    set color to w,n/bg,,,n/w
    set color to w/b
    @ 23,05 say "<ESC> Para sair. "
    set color to w,n/bg,,,n/w
    @ 20,58 get yticket pict "999999"
    read
    if readkey() = 12  .or. lastkey() = 27
        close all
        return
    endif
    wticket = strzero(yticket,6,0)
    if empty(wticket) .or. wticket = "000000"
        close all
        return
    endif
    set color to r/w
    @ 20,58 say wticket pict "999999"
    select 3
    locate for ticket = wticket
    if found()
        iguala()
        wdesconto = desconto
    else
        do mensagem with "Ticket nao cadastrado...",8
        loop
    endif
    set color to /w
    @ 20,58 say wticket pict "999999"
    exit
enddo

select 2
index on ticket to &zmicro for ticket = wticket
set index to &zmicro
iguala()
select 5
seek wcodigo_cli
if found()
    iguala()
    zcodigo_cli = wcodigo_cli
    znome       = wnome
else
    zcodigo_cli = space(06)
endif
select 4
set index to ind01.ven
seek wcodigo_ven
if found()
    iguala()
    zcodigo_ven = wcodigo_ven
    znome_ven   = wnome_ven
else
    zcodigo_ven = space(03)
endif

select 2
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
total= 0
do while inkey()#27 .and. !eof()
    @ prow()+1,02     say zempresa  pict "@!"
    @ prow()+1,02     say zendereco pict "@!"
    @ prow()+1,02     say "Telefone: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),pcol()+1 say "FAX: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),70       say "Pag."
    @ prow(),75       say strzero(wpag,3,0) pict "999"
    @ prow()+1,02     say "Extrato de Venda a Vista "
    @ prow()+1,02     say "Data..:"
    @ prow(),pcol()+1 say wdata_ti pict "@d"
    if !empty(zcodigo_cli)
        select 5
        set index to ind01.cli
        seek zcodigo_cli
        iguala()
        @ prow()+1,02     say "Cliente...:"
        @ prow(),pcol()+1 say zcodigo_cli pict "@!"
        @ prow(),pcol()+2 say znome       pict "@!"
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
        select 2
    endif
    @ prow()+1,02     say "Condicoes de Pagamento:"
    @ prow(),pcol()+1 say whistorico pict "@!"
    if !empty(wcartao_t)
        @ prow()+1,02     say "Cartao....:"
        @ prow(),pcol()+1 say wcartao_t pict "@!"
    endif
    if !empty(wcheque)
        @ prow()+1,02     say "Cheque NRO:"
        @ prow(),pcol()+1 say wcheque pict "@!"
        @ prow(),pcol()+1 say "Banco..:"
        @ prow(),pcol()+1 say wbanco pict "@!"
        @ prow(),pcol()+1 say "Agencia.:"
        @ prow(),pcol()+1 say wagencia pict "@!"
    endif
    if !empty(zcodigo_ven)
        @ prow()+1,02     say "Vendedor..:"
        @ prow(),pcol()+1 say zcodigo_ven pict "@!"
        @ prow(),pcol()+2 say znome_ven   pict "@!"
    endif
    @ prow()+2,02  say "Codigo  Especificacao                            Quant    P.Venda Total      "
    @ prow()+1,02  say "~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~ ~~~~~~~ ~~~~~~~~~~ "
    do while prow()<55 .and. !eof()
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
        total = total + compras
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
@ prow(),pcol() say total pict "@e@z 999,999.99"
@ prow(),pcol() say " Desconto..:"
@ prow(),pcol() say wdesconto pict "@e@z 999,999.99"
@ prow(),pcol() say " Total a Pagar..:"
@ prow(),pcol() say total - wdesconto pict "@e@z 99,999.99"
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
return
*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
