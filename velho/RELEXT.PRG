*****************************************************************************
* Programa...: RELEXT.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao: 07/01/97       *
* Objetivo...: EXTRATO CONTAS DE CLIENTES - A PRAZO                         *
* Sistema....: SISPAN - Sistema de Panifica‡Æo                              *
*****************************************************************************
save screen to telarel
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

wcopia  = 0
wmes    = space(06)
wdata   = space(08)
wtotal  = 0
credito = 0
wdata_est = space(08)
save screen to telarel
select 2
set index to ind01.cli
seek wcodigo_cli
iguala()
select 3
index on codigo_cli to &zmicro for codigo_cli = wcodigo_cli
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
do while inkey()#27 .and. !eof()
    @ prow()+1,02     say zempresa  pict "@!"
    @ prow()+1,02     say zendereco pict "@!"
    @ prow()+1,02     say "Telefone: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),pcol()+1 say "FAX: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),70       say "Pag."
    @ prow(),75       say strzero(wpag,3,0) pict "999"
    @ prow()+1,02     say "Extrato de Conta Cliente  -  a Prazo"
    @ prow()+2,02     say "Cliente.....:"
    @ prow(),15       say wcodigo_cli pict "99999"
    @ prow(),22       say wnome       pict "@!"
    @ prow()+1,02     say "Data Debito :"
    @ prow(),15       say wdata_deb   pict "99/99/99"
    @ prow()+2,02     say "Codigo  Especificacao                   Data       Quantidade   Valor       "
    @ prow()+1,02     say "~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~ ~~~~~~~~~~~  ~~~~~~~~~~~~"
    do while prow()<55 .and. !eof()
        iguala()
        @ prow()+1,02 say wcodigo_est pict "999999"
        select 1
        seek wcodigo_est
        iguala()
        @ prow(),10   say wdescricao  pict "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
        @ prow(),42   say wdata       pict "@d"
        @ prow(),53   say wquantidade pict "999.99"
        wprec = wpreco_apr
        compras = (wprec * wquantidade)
        @ prow(),66   say compras    pict "999,999.99"
        wlin = 1
        total = total + compras
        select 3
        skip
        esc = inkey()
        if esc = 12
            @ prow()+2,40 say "***********  Impressao Cancelada..."
            close all
            return
        endif
    enddo
    eject
    wpag = wpag + 1
enddo
reg = recno()
@ prow()+1,02  say "----------------------------------------------------------------------------"
@ prow()+1,03  say " Total de Compas..................................."
@ prow(),62    say total pict "9,999,999,999.99"
@ prow()+1,03  say " Total de Credito................................."
@ prow(),62    say wcredito pict "9,999,999,999.99"
@ prow()+1,03  say " Total a Pagar...................................."
@ prow(),62    say total-wcredito pict "9,999,999,999.99"
@ prow()+2,02  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
@ prow()+1,02  say "              Sistema de Administracao Estoque - Versao 1.0"
@ prow(),03    say date() pict "@d"
@ prow(),69    say time() pict "99:99"
set cons on
set devi to screen
set print off
set print to
verfil(zdireto)
return
*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
