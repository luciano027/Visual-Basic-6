*****************************************************************************
* Programa...: RELEXTVE.PRG                                                 *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao:                *
* Objetivo...: RELATORIO DE VENDAS POR DATA - VENDEDOR                      *
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

set color to 0/7
wcopia   = 0
wcodigo_cli = "00000"
g        = 0
total = 0
save screen to telafor
set color to n/w
dt = quadro(17,40,19,78)
@ 18,42 say "Cliente....:"
do while .t.
    set color to w/b
    @ 23,05 say "<ESC> Para sair. <00000> Pesquisa"
    set color to w,n/bg,,,n/w
    @ 18,57 get wcodigo_cli pict "99999"
    read
    if readkey() = 12  .or. lastkey() = 27
        close all
        return
    endif
    if empty(wcodigo_cli) .or. wcodigo_cli = "00000"
        select 2
        save screen to tela_d
        do pes_cli
        restore screen from tela_d
        iguala()
    endif
    set color to n/w
    @ 18,57 say wcodigo_cli pict "99999"
    select 2
    set index to ind01.cli
    seek wcodigo_cli
    if !found()
        do mensagem with "Cliente nao cadastrado...",8
        loop
    endif
    iguala()
    exit
enddo
set color to w+/b
dt = quadro(15,18,17,35)
set color to *+w/b
@ 16,20 say "Aguarde..."
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
it = 0
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
    do while it < 6 .and. !eof()
        iguala()
        @ prow()+1,02 say wcodigo_est pict "999999"
        select 1
        seek wcodigo_est
        iguala()
        @ prow(),10   say wdescricao  pict "@s30"
        @ prow(),42   say wdata       pict "@d"
        @ prow(),53   say wquantidade pict "@e@z 99,999.99"
        wprec = wvalor
        compras = (wprec * wquantidade)
        @ prow(),66   say compras    pict "999,999.99"
        wlin = 1
        wtotal = wtotal + compras
        select 3
        skip
    enddo
    wpag = wpag + 1
    if !eof()
        @ prow()+1,02  say "----------------------------------------------------------------------------"
        @ prow()+1,02  say "                                                                 continua..."
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
    endif
    tsalto = zsalto + (8 - it)
    @ prow()+tsalto,01 say "."
    it = 0
    setprc(0,0)
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
