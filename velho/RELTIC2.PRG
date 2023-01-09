*****************************************************************************
* Programa...: RELENT.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao:                *
* Objetivo...: RELATORIO DE ENTRADA DE MERCADORIAS NO ESTOQUE               *
* Sistema....: SISPAN - Sistema de Panifica‡Æo                              *
*****************************************************************************
select 1
use empresa
locate for !empty(empresa)
if found()
    wempresa  = empresa
    wendereco = endereco
else
    do mensagem with "Empresa nao Cadastrada ...",8
    return
endif
select 2
use ticket
if !file("ind01.tic")
    index on ticket  to ind01.tic
endif
set index to ind01.tic

set color to 0/7
wcopia = 0
wmes   = space(01)
wdata_i= date()
wdata_f= date()
g      = 0
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

select 2
index on dtos(data_ti) to indrel.ent for data_ti >= wdata_i .and. data_ti <=  wdata_f .and. pagamento = "D"
set index to indrel.ent

limpa()
set color to w+/r
dt = quadro(15,18,17,35)
set color to *+w/r
@ 16,20 say "Aguarde..."
set color to w/b
set print to arq.txt
set devi to print
set cons off
wpag = 1
gera_c = 0
gera_d = 0
do while inkey()#27 .and. !eof()
    @ prow()+1,02     say zempresa  pict "@!"
    @ prow()+1,02     say zendereco pict "@!"
    @ prow()+1,02     say "Telefone: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),pcol()+1 say "FAX: "
    @ prow(),pcol()+1 say ztelefone pict "@!"

    @ prow(),70       say "Pag."
    @ prow(),75       say strzero(wpag,3,0) pict "999"
    @ prow()+1,02     say "Relatorio TICKET CAIXA - Saida no Caixa"
    @ prow()+1,02     say "Periodo :"
    @ prow(),pcol()+1 say wdata_i pict "@d"
    @ prow(),pcol()+1 say " a "
    @ prow(),pcol()+1 say wdata_f pict "@d"
    @ prow()+2,02     say "Data       Ticket Historico                                Debito  "
    @ prow()+1,02     say "~~~~~~~~~~ ~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~ "
    store 0 to pagina_c,pagina_d
    do while prow()<55 .and. !eof()
        iguala()
        @ prow()+1,02 say wdata_ti    pict "@d"
        @ prow(),13   say wticket     pict "999999"
        @ prow(),20   say whistorico  pict "@s40"
        @ prow(),61   say wtotal_ti   pict "@e@z 99,999.99"
        pagina_d = pagina_d +  wtotal_ti
        skip
        esc = inkey()
        if esc = 12
            @ prow()+2,40 say "***********  Impressao Cancelada..."
            close all
            return
        endif
    enddo
    gera_d = gera_d + pagina_d
    if eof()
        do while prow()<50
            @ prow()+1,00 say ""
        enddo
    endif
    @ prow()+1,01  say "-------------------------------------------------------------------------"
    @ prow()+1,03  say "Total na Pagina:"
    @ prow(),69    say pagina_d   pict "@e@z 99,999.99"
    @ prow()+1,03  say "Total a Geral:"
    @ prow(),69    say gera_d     pict "@e@z 99,999.99"
    @ prow()+2,01  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    @ prow()+1,01  say "              Sistema de Administracao Estoque - Versao 1.0"
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
