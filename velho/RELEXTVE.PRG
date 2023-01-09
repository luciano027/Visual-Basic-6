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
if !net_use("vendedor",.f.,"ind01.ven","ind02.ven")
    return
endif

select 3
if !net_use("cticket",.f.)
    return
endif

set color to 0/7
wcopia   = 0
wmes     = space(01)
wdata_i  = date()
wdata_f  = date()
wcodigo  = space(03)
wcodigo_ven = "000"
g        = 0
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
select 3
index on dtos(data_ti) to &zmicro for data_ti >= wdata_i .and. data_ti <=  wdata_f;
.and. codigo_ven = wcodigo_ven
set index to &zmicro
limpa()
set color to w/b
set print to &zdireto
set devi to print
set cons off
wpag = 1
geral = 0
do while inkey()#27 .and. !eof()
    @ prow()+1,02     say zempresa  pict "@!"
    @ prow()+1,02     say zendereco pict "@!"
    @ prow()+1,02     say "Telefone: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),pcol()+1 say "FAX: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),70       say "Pag."
    @ prow(),76       say strzero(wpag,3,0)  pict "999"
    @ prow()+1,02     say "Relatorio EXTRATO DE VENDAS"
    @ prow()+2,02     say "Vendedor...:"
    @ prow(),pcol()+1 say wcodigo_ven pict "@!"
    @ prow(),pcol()+1 say wnome_ven   pict "@!"
    @ prow()+1,02     say "Periodo :"
    @ prow(),pcol()+1 say wdata_i pict "@d"
    @ prow(),pcol()+1 say " a "
    @ prow(),pcol()+1 say wdata_f pict "@d"
    @ prow()+2,02     say " Data    Historico                     Valor Total   "
    @ prow()+1,02     say "~~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~~~~ "
    pagina= 0
    do while prow()<55 .and. !eof()
        iguala()
        ydata_ti  = wdata_ti
        wtotalv  = 0
        do while ydata_ti = wdata_ti
            wtotalv = wtotalv + wtotal_ti
            skip
            iguala()
        enddo
        whist = " Vendas do Dia................."
        @ prow()+1,02 say ydata_ti    pict "@d"
        @ prow(),13   say whist       pict "@s27"
        @ prow(),44   say wtotalv     pict "@e@z 99,999.99"
        wlin = 1
        pagina= pagina + wtotalv
    enddo
    geral = geral + pagina
    if eof()
        do while prow()<55
            @ prow()+1,00 say ""
        enddo
    endif
    @ prow()+1,02  say "-----------------------------------------------------------------------"
    @ prow()+1,03  say "             Total na Pagina........."
    @ prow(),44    say pagina pict "@e@z 999,999.99"
    @ prow()+1,03  say "             Total Geral............."
    @ prow(),44    say geral  pict "@e@z 999,999.99"
    @ prow()+1,02  say "-----------------------------------------------------------------------"
    @ prow()+1,03  say " Percentual.......................... "
    @ prow(),41    say wcomissao pict "@e@z 999.99"
    @ prow()+1,03  say " Total a Receber..................... "
    wtotalg = (geral*wcomissao)/100
    @ prow(),41    say wtotalg  pict "@e@z 99,999.99"
    @ prow()+2,02  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    @ prow()+1,02  say "             Sistema de Estoque e Vendas   - Versao 2.1"
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
***************************************************************************
*                        f i m                                            *
***************************************************************************
