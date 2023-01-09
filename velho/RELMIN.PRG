*****************************************************************************
* Programa...: RELMINI.PRG                                                  *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......:                   Data da Ultima Atualizacao:                *
* Objetivo...: RELATORIO SALDO MINIMO                                       *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************
select 1
if !net_use("estoque",.f.,"ind01.est","ind02.est")
    return
endif

select 1
index on descricao   to &zmicro for saldo < minino
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
geral= 0
do while inkey()#27 .and. !eof()
    @ prow()+1,02     say zempresa  pict "@!"
    @ prow()+1,02     say zendereco pict "@!"
    @ prow()+1,02     say "Telefone: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),pcol()+1 say "FAX: "
    @ prow(),pcol()+1 say ztelefone pict "@!"

    @ prow(),70       say "Pag."
    @ prow(),75       say strzero(wpag,3,0) pict "999"
    @ prow()+1,02     say "Relatorio SALDO MININO"
    @ prow()+2,02     say " Codigo  Produto                 Unid. P.Compra  Saldo Atual  Saldo Minimo "
    @ prow()+1,02     say "~~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~ ~~~~~~~~~ ~~~~~~~~~~~  ~~~~~~~~~~~~"
    do while prow()<55 .and. !eof()
        iguala()
        @ prow()+1,02 say wcodigo_est  pict "999999"
        @ prow(),11   say wdescricao   pict "@s23"
        @ prow(),35   say wunidade     pict "@!"
        @ prow(),41   say wpreco_com   pict "@e@z 99,999.99"
        @ prow(),52   say wsaldo       pict "@e@z 99,999.99"
        @ prow(),65   say wminino      pict "@e@z 99,999.99"
        wlin = 1
        skip
        esc = inkey()
        if esc = 12
            @ prow()+2,40 say "***********  Impressao Cancelada..."
            close all
            return
        endif
    enddo
    if eof()
        do while prow()<55
            @ prow()+1,00 say ""
        enddo
    endif
    @ prow()+2,01  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    @ prow()+1,01  say "              Sistema de Administracao Estoque - Versao 1.0"
    @ prow(),02    say date() pict "@d"
    @ prow(),69    say time() pict "99:99"
    wpag = wpag +1
    eject
enddo
set cons on
set devi to screen
set print off
set print to
lertexto(zdireto)
close all
return
*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
