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
else
    inicia()
endif
select 2
if !net_use("clientes",.f.,"ind01.cli","ind02.cli","ind03.cli")
    return
else
    inicia()
endif
select 3
if !net_use("prazo",.f.,"ind01.pra")
    return
else
    inicia()
endif
set color to 0/7
wcopia      = 0
wcodigo_cli = "00000"
g           = 0
total       = 0
save screen to telafor
set color to n/w
dt = quadro(17,40,19,78)
@ 18,42 say "Produto....:"
do while .t.
    set color to w/b
    @ 23,05 say "<ESC> Para sair. <000000> Pesquisa"
    set color to w,n/bg,,,n/w
    @ 18,57 get wcodigo_est pict "999999"
    read
    if readkey() = 12  .or. lastkey() = 27
        close all
        return
    endif
    if empty(wcodigo_est) .or. wcodigo_est = "000000"
        select 1
        save screen to tela_d
        do pes_est
        restore screen from tela_d
        iguala()
    endif
    set color to n/w
    @ 18,57 say wcodigo_est pict "999999"
    select 1
    set index to ind01.est
    seek wcodigo_est
    if !found()
        do mensagem with "Produto nao cadastrado...",8
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
index on codigo_cli to &zmicro for  codigo_est = wcodigo_est
set index to &zmicro
limpa()
set color to w/b
set print to &zdireto
set devi to print
set cons off
wpag   = 1
geral  = 0
wtotal = 0
do while inkey()#27 .and. !eof()
    @ prow()+1,02     say zempresa  pict "@!"
    @ prow()+1,02     say zendereco pict "@!"
    @ prow()+1,02     say "Telefone: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),pcol()+1 say "FAX: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),70       say "Pag."
    @ prow(),75       say strzero(wpag,3,0) pict "999"
    @ prow()+1,02     say "Extrato de Conta Cliente  -  Produtos"
    @ prow()+2,02     say "Produto...:"
    @ prow(),pcol()+1 say wcodigo_est pict "99999"
    @ prow(),pcol()+2 say wdescricao  pict "@!"
    @ prow()+2,02     say "Cliente                                      Data Debit Quant     P Unitario"
    @ prow()+1,02     say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~ ~~~~~~~~~ ~~~~~~~~~~"
    do while prow()<55 .and. !eof()
        iguala()
        select 2
        seek wcodigo_cli
        iguala()
        @ prow()+1,02 say wcodigo_cli pict "99999"
        @ prow(),pcol() say " - "
        @ prow(),10   say wnome       pict "@s35"
        @ prow(),47   say wdata_deb   pict "@d"
        @ prow(),58   say wquantidade pict "@e@z 99,999.99"
        @ prow(),68   say wvalor      pict "@e@z 99,999.99"
        select 3
        skip
        esc = inkey()
        if esc = 12
            @ prow()+2,40 say "***********  Impressao Cancelada..."
            close all
            return
        endif
    enddo
    wpag = wpag + 1
    if !eof()
        @ prow()+1,02  say "----------------------------------------------------------------------------"
    else
        @ prow()+2,02  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        @ prow()+1,02  say "              Sistema de Administracao Estoque - Versao 1.0"
        @ prow(),03    say date() pict "@d"
        @ prow(),69    say time() pict "99:99"
    endif
    eject
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
