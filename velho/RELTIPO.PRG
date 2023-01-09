***************************************************************************
* Programa Tabela de Precos - Relatorio                                   *
*                                                                         *
* Emite relatorio geral (tabela de precos) indexado por espec_est         *
*                                                                         *
* Luciano 26/08/93                                                        *
*                                                                         *
***************************************************************************

select 1
if !net_use("tipo",.f.,"ind01.tip")
    return
endif

limpa()
set color to w+/b
dt = quadro(15,18,17,35)
set color to *+w/b
@ 16,20 say "Aguarde..."
set color to w/b
set print to &zdireto
set devi to print
set cons off
wpag = 1
do while inkey()#27 .and. !eof()
    @ prow()+1,02     say zempresa  pict "@!"
    @ prow()+1,02     say zendereco pict "@!"
    @ prow(),pcol()+2 say ztelefone pict "@!"
    @ prow(),70       say "Pag."
    @ prow(),75       say strzero(wpag,3,0)  pict "999"
    @ prow()+1,02     say "Relatorio TIPO DE DOCUMENTOS"
    @ prow()+2,02     say "Codigo       Documento                              "
    @ prow()+1,02     say "~~~~~~~~~~~  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    do while prow()<55 .and. !eof()
        iguala()
        @ prow()+1,03 say wcodigo    pict "@!"
        @ prow(),14   say wdescricao pict "@s32"
        wlin = 1
        skip
    enddo
    if eof()
        do while prow()<55
            @ prow()+1,00 say ""
        enddo
    endif
    @ prow()+2,01  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    @ prow()+1,01  say "             Sistema de Estoque e Vendas   - Versao 2.1"
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
***************************************************************************
*                                                                         *
*                                                                         *
*                        f i m                                            *
*                                                                         *
*                                                                         *
*                                                                         *
***************************************************************************
