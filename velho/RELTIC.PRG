*****************************************************************************
* Programa...: RELENT.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao:                *
* Objetivo...: RELATORIO DE ENTRADA DE MERCADORIAS NO ESTOQUE               *
* Sistema....: SISPAN - Sistema de Panifica‡Æo                              *
*****************************************************************************
select 1
if !net_use("cticket",.f.)
    return
endif
select 2
if !net_use("vendedor",.f.,"ind01.ven","ind02.ven")
    return
endif

set color to 0/7
wcopia = 0
wmes   = space(01)
wdata_i= date()
wdata_f= date()
wcodigo_ven = space(03)
g      = 0
save screen to telafor
set color to n/w
dt = quadro(17,40,20,78)
@ 18,42 say "Vendedor.....:"
@ 19,42 say "Periodo....:           a  "
do while .t.
    set color to w/b
    @ 23,05 say "<ESC> Para sair. <000> Pesquisa"
    set color to w,n/bg,,,n/w
    @ 18,57 get wcodigo_ven pict "999"
    read
    if readkey() = 12  .or. lastkey() = 27
        close all
        return
    endif
    if empty(wcodigo_ven) .or. wcodigo_ven = "000"
        select 2
        save screen to tela_d
        do pes_ven
        restore screen from tela_d
        iguala()
    endif
    set color to n/w
    @ 18,57 say wcodigo_ven pict "999"
    select 2
    set index to ind01.ven
    seek wcodigo_ven
    if !found()
        do mensagem with "Vendedor nao cadastrado...",8
        loop
    endif
    iguala()
    exit
enddo
set color to w,n/bg,,,n/w
@ 19,54 get wdata_i  pict "@d"
read
if readkey() = 12 .or. lastkey()=27
    close all
    return
endif
@ 19,67 get wdata_f  pict "@d"
read
set color to w+/b
dt = quadro(15,18,17,35)
set color to *+w/b
@ 16,20 say "Aguarde..."

select 1
index on dtos(data_ti) to &zmicro for data_ti >= wdata_i .and. data_ti <=  wdata_f;
.and. codigo_ven = wcodigo_ven
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
store 0 to wdin,wche,wcre,w30d,w60d,w90d,wret
do while inkey()#27 .and. !eof()
    @ prow()+1,02     say zempresa  pict "@!"
    @ prow()+1,02     say zendereco pict "@!"
    @ prow()+1,02     say "Telefone: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),pcol()+1 say "FAX: "
    @ prow(),pcol()+1 say ztelefone pict "@!"

    @ prow(),70       say "Pag."
    @ prow(),75       say strzero(wpag,3,0) pict "999"
    @ prow()+1,02     say "Relatorio TICKET CAIXA"
    @ prow()+1,02     say "Periodo :"
    @ prow(),pcol()+1 say wdata_i pict "@d"
    @ prow(),pcol()+1 say " a "
    @ prow(),pcol()+1 say wdata_f pict "@d"
    @ prow()+2,02     say"Data       Ticket Historico                  Dinheiro   Credito   A Prazo "
    @ prow()+1,02     say"~~~~~~~~~~ ~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~ ~~~~~~~~~ ~~~~~~~~~"
    do while prow()<55 .and. !eof()
        iguala()
        @ prow()+1,02 say wdata_ti    pict "@d"
        @ prow(),13   say wticket     pict "999999"
        @ prow(),20   say whistorico  pict "@s24"
        if wtipo_tipo     = "1"
            @ prow(),47   say wtotal_ti   pict "@e@z 99,999.99"
            wdin = wdin + wtotal_ti
        elseif wtipo_tipo = "3"
            @ prow(),58   say wtotal_ti   pict  "@e@z 99,999.99"
            wcre = wcre + wtotal_ti
        elseif wtipo_tipo = "4"
            @ prow(),68    say wtotal_ti   pict "@e@z 99,999.99"
            w30d = w30d + wtotal_ti
        elseif wtipo_tipo = "7"
            @ prow(),47   say wtotal_ti   pict "@e@z 99,999.99"
            wret = wret + wtotal_ti
        endif
        skip
        esc = inkey()
        if esc = 12
            @ prow()+2,40 say "***********  Impressao Cancelada..."
            close all
            return
        endif
    enddo
    if eof()
        @ prow()+1,01  say "----------------------------------------------------------------------------"
        @ prow()+1,02  say "Total Dinheiro..:"
        @ prow(),pcol()+1 say wdin pict "@e@z 999,999.99"
        @ prow()+1,02  say "Total Credito...:"
        @ prow(),pcol()+1 say wcre pict "@e@z 999,999.99"
        @ prow()+1,02  say "Total a Prazo...:"
        @ prow(),pcol()+1 say w30d pict "@e@z 999,999.99"
        @ prow()+1,01  say "----------------------------------------------------------------------------"
        @ prow()+1,02  say "Total Retirada..:"
        @ prow(),pcol()+1 say wret pict "@e@z 999,999.99"
        @ prow()+1,01  say "----------------------------------------------------------------------------"
        @ prow()+1,01  say "Total no Caixa.R$"
        wtotal_c = (wdin+wcre)-wret
        @ prow(),20    say wtotal_c pict "@e@z 99,999.99"
        @ prow()+1,01  say "----------------------------------------------------------------------------"
    endif
    @ prow()+2,01  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
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
lertexto(zdireto)
close all
return
    *****************************************************************************
    *                          F   I   M                                        *
    *****************************************************************************
