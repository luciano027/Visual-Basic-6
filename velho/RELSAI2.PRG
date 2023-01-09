*****************************************************************************
* Programa...: RELSAI.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......:                   Data da Ultima Atualizacao:                *
* Objetivo...: RELATORIO DE SAIDA DE MERCADORIAS NO PERIODO                 *
* Sistema....: SISPAN - Sistema de Panifica‡Æo                              *
*****************************************************************************
select 1
if !net_use("estoque",.f.,"ind01.est","ind02.est")
    return
endif
select 2
if !net_use("csaida",.f.)
    return
endif
select 3
if !net_use("vendedor",.f.,"ind01.ven","ind02.ven")
    return
endif


set color to 0/7
wcopia = 0
wmes   = space(01)
wdata_i= date()
wdata_f= date()
wcodigo_ven = space(03)
zcodigo_ven = space(03)
g      = 0
save screen to telafor
set color to n/w
dt = quadro(18,40,21,78)
@ 19,42 say "Periodo....:           a  "
@ 20,42 say "Cod. Vendedor.:"
set color to w,n/bg,,,n/w
@ 19,54 get wdata_i  pict "@d"
read
if readkey() = 12 .or. lastkey()=27
    close all
    return
endif
@ 19,67 get wdata_f  pict "@d"
read
do while .t.
    set color to w/b
    @ 23,05 say "<ESC> Para sair.<000> para consulta"
    set color to w,n/bg,,,n/w
    save screen to telafor
    @ 20,57 get wcodigo_ven pict "999" valid !empty(wcodigo_ven)
    read
    if readkey() = 12 .or. lastkey() = 27
        select 3
        libera()
        close all
        return
    endif
    if wcodigo_ven = "000"
        select 3
        save screen to tela_0
        do pes_ven
        restore screen from tela_0
        iguala()
    endif
    ycodigo_ven = wcodigo_ven
    set color to /w
    @ 20,57 say wcodigo_ven pict "999"
    select 3
    set index to ind01.ven
    seek wcodigo_ven
    if !found()
        do mensagem with "Codigo nao cadastrao ..",8
        loop
    endif
    iguala()
    zcodigo_ven = wcodigo_ven
    znome_ven   = wnome_ven
    set color to /w
    @ 20,65 say wnome_ven  pict "@s10"
    exit
enddo
select 2
index on codigo_est to &zmicro for data >= wdata_i .and. data <=  wdata_f .and.;
codigo_ven = wcodigo_ven
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
geral = 0
do while inkey()#27 .and. !eof()
    @ prow()+1,02     say zempresa  pict "@!"
    @ prow()+1,02     say zendereco pict "@!"
    @ prow()+1,02     say "Telefone: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),pcol()+1 say "FAX: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),70       say "Pag."
    @ prow(),75       say strzero(wpag,3,0) pict "999"
    @ prow()+1,02     say "Relatorio SAIDA MERCADORIAS NO ESTOQUE"
    @ prow()+2,02     say "Codigo Vendedor.:"
    @ prow(),pcol()+1 say zcodigo_ven pict "@!"
    @ prow(),pcol()+1 say znome_ven   pict "@!"
    @ prow()+1,02     say "Periodo :"
    @ prow(),pcol()+1 say wdata_i pict "@d"
    @ prow(),pcol()+1 say " a "
    @ prow(),pcol()+1 say wdata_f pict "@d"
    @ prow()+2,02     say "Codigo     Mercadoria                         Quant.    Preco      Total"
    @ prow()+1,02     say "~~~~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~  ~~~~~~~~  ~~~~~~~~~~"
    pagina = 0
    do while prow()<55 .and. !eof()
        iguala()
        select 1
        seek wcodigo_est
        iguala()
        select 2
        @ prow()+1,02 say wcodigo_est pict "999999"
        @ prow(),13   say wdescricao  pict "@s34"
        @ prow(),48   say wquantidade pict "@e@z 99999.99"
        @ prow(),57   say wpreco_ven  pict "@e@z 99,999.99"
        wvalor = wquantidade * wpreco_ven
        @ prow(),68   say wvalor      pict "@e@z 99,999.99"
        wlin = 1
        pagina = pagina + wvalor
        esc = inkey()
        if esc = 12
            @ prow()+2,40 say "***********  Impressao Cancelada..."
            close all
            return
        endif
        skip
    enddo
    geral = geral + pagina
    @ prow()+1,02  say "----------------------------------------------------------------------------"
    @ prow()+1,03  say "Total na Pagina...."
    @ prow(),63    say pagina pict "@e@z 999,999.99"
    if eof()
        @ prow()+1,03  say "Total Geral........"
        @ prow(),63    say geral  pict "@e@z 999,999.99"
    endif
    @ prow()+1,02  say "----------------------------------------------------------------------------"
    @ prow()+2,02  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    @ prow()+1,02  say "              Sistema de Administracao Estoque - Versao 1.0"
    @ prow(),02    say date() pict "@d"
    @ prow(),69    say time() pict "99:99"
    eject
    wpag = wpag + 1
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
