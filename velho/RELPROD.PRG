*****************************************************************************
* Programa...: RELPROD.PRG                                                  *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......:                   Data da Ultima Atualizacao:                *
* Objetivo...: RELATORIO DE PRODUTOS CADASTRADOS                            *
* Sistema....: SISPAN - Sistema de Panifica‡Æo                              *
*****************************************************************************
select 1
if !net_use("estoque",.f.,"ind02.est","ind01.est")
    return
endif

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
    @ prow()+1,02     say "Relatorio PRODUTOS CADASTRADOS "
    @ prow()+2,02     say "Codigo  Produto                         Uni  Saldo     Venda    Total "
    @ prow()+1,02     say"~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~ ~~~~~~~~~ ~~~~~~~~ ~~~~~~~~~"
    pagina = 0
    do while prow()<55 .and. !eof()
        iguala()
        @ prow()+1,02 say wcodigo_est  pict "999999"
        @ prow(),10   say wdescricao   pict "@s30"
        @ prow(),42   say wunidade     pict "@!"
        @ prow(),47   say wsaldo       pict "@e@z 99999.99"
        @ prow(),57   say wpreco_ven   pict "@e@z 9,999.99"
        wtotal = wsaldo * wpreco_ven
        @ prow(),66   say wtotal       pict "@e@z 99,999.99"
        wlin = 1
        pagina= pagina+ wtotal
        skip
        esc = inkey()
        if esc = 12
            @ prow()+2,40 say "***********  Impressao Cancelada..."
            close all
            return
        endif
    enddo
    geral = geral + pagina
    if eof()
        do while prow()<55
            @ prow()+1,00 say ""
        enddo
    endif
    @ prow()+1,02  say "-------------------------------------------------------------------------"
    @ prow()+1,03  say "Total na Pagina"
    @ prow(),65    say pagina pict "@e@z 999,999.99"
    @ prow()+1,03  say "Total Geral"
    @ prow(),65    say geral  pict "@e@z 999,999.99"
    @ prow()+1,02  say "-------------------------------------------------------------------------"
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
lertexto(zdireto)
close all
return
*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
