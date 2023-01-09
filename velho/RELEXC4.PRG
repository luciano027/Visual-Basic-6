*****************************************************************************
* Programa...: RELEXTVE.PRG                                                 *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao:                *
* Objetivo...: RELATORIO DE VENDAS POR DATA - VENDEDOR                      *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************
select 1
use estoque
if !file("ind01.est") .or. !file("ind02.est")
    index on codigo_est  to ind01.est
    index on descricao   to ind02.est
endif
set index to ind01.est,ind02.est
select 2
use clientes
if !file("ind01.cli") .or. !file("ind02.cli")
    index on codigo_cli to ind01.cli
    index on nome       to ind02.cli
endif
set index to ind01.cli,ind02.cli
select 3
use prazo
if !file("ind01.pra") .or. !file("ind02.pra") .or. !file("ind03.pra")
    index on codigo_cli to ind01.pra
    index on codigo_est to ind02.pra
    index on codigo_cli+codigo_est to ind03.pra
endif
set index to ind01.pra,ind02.pra,ind03.pra
select 4
use empresa
locate for !empty(empresa)
if found()
    wempresa  = empresa
    wendereco = endereco
else
    do mensagem with "Empresa nao Cadastrada ...",8
    return
endif

select 1
inicia()
select 2
inicia()
select 3
inicia()


wdata_i  = date()
wdata_f  = date()
wcod_ven = space(03)
save screen to telafor
set color to n/w
dt = quadro(18,40,20,78)
@ 19,42 say "Periodo....:           a  "
set color to w,n/bg,,,n/w
@ 19,54 get wdata_i  pict "@d"
read
if readkey() = 12 .or. lastkey()=27
    close all
    return
endif
@ 19,67 get wdata_f  pict "@d"
read

set color to w+/r
dt = quadro(15,18,17,35)
set color to *+w/r
@ 16,20 say "Aguarde..."

select 3
index on codigo_cli to indrel.sai for data >= wdata_i .and. data <=  wdata_f
set index to indrel.sai

limpa()
set color to w/b
set print to arq.txt
set devi to print
set cons off
wpag = 1
ydata_ti = wdata_i
total = 0
t = 0
do while !eof()
    @ prow()+1,02     say zempresa  pict "@!"
    @ prow()+1,02     say zendereco pict "@!"
    @ prow()+1,02     say "Telefone: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),pcol()+1 say "FAX: "
    @ prow(),pcol()+1 say ztelefone pict "@!"

    @ prow(),70       say "Pag."
    @ prow(),75       say strzero(wpag,3,0) pict "999"
    @ prow()+1,02     say "Extrato de Venda a Prazo - Geral"
    @ prow()+1,02     say "Periodo :"
    @ prow(),pcol()+1 say wdata_i pict "@d"
    @ prow(),pcol()+1 say " a "
    @ prow(),pcol()+1 say wdata_f pict "@d"
    @ prow()+2,02     say "Valores em R$"
    @ prow()+1,02     say "Data       Cliente       Descricao             Quantidade    Valor       "
    @ prow()+1,02     say "~~~~~~~~~~ ~~~~~~~~~~~~~ ~~~~~~~~~~~~~~~~~~~~  ~~~~~~~~~~~~~ ~~~~~~~~~~~~"
    pagina = 0
    do while prow()<55 .and. !eof()
        iguala()
        select 2
        seek wcodigo_cli
        iguala()
        select 1
        seek wcodigo_est
        iguala()
        select 3
        @ prow()+1,02 say wdata                  pict "@d"
        @ prow(),13   say wnome                  pict "@s13"
        @ prow(),27   say wdescricao             pict "@s20"
        @ prow(),49   say wquantidade            pict "@e 99,999.99"
        @ prow(),63   say wpreco_ven*wquantidade pict "@e 99,999.99"
        pagina  = pagina  + (wquantidade * wpreco_ven)
        skip
    enddo
    total = total + pagina
    @ prow()+1,01  say "-------------------------------------------------------------------------"
    @ prow()+1,03  say " Total na Pagina........."
    @ prow(),62   say pagina     pict "@e@z 999,999.99"
    @ prow()+1,03  say " Total a Geral .........."
    @ prow(),62   say total     pict "@e@z 999,999.99"
    @ prow()+1,01  say "-------------------------------------------------------------------------"
    @ prow()+2,02  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    @ prow()+1,02  say "              Sistema de Administracao Estoque - Versao 1.0"
    @ prow(),02    say date() pict "@d"
    @ prow(),69    say time() pict "99:99"
    wpag = wpag + 1
    eject
enddo
set cons on
set devi to screen
set print off
set print to
verfil("arq.txt")
close all
return
*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
