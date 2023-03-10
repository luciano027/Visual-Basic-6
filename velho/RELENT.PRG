*****************************************************************************
* Programa...: RELENT.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao:                *
* Objetivo...: RELATORIO DE ENTRADA DE MERCADORIAS NO ESTOQUE               *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************
select 1
if !net_use("estoque",.f.,"ind01.est","ind02.est")
    return
endif
select 2
if !net_use("fornece",.f.,"ind01.for","ind02.for")
    return
endif
select 3
if !net_use("entrada",.f.,"ind01.ent")
    return
endif

set color to 0/7
wcopia = 0
wmes   = space(01)
wdata_i= date()
wdata_f= date()
wcodigo_est = space(06)
zcodigo_est = space(06)
g      = 0
save screen to telafor
set color to n/w
dt = quadro(18,40,21,78)
@ 19,42 say "Periodo....:           a  "
@ 20,42 say "Codigo Estoque:"
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
    @ 23,05 say "<ESC> Para sair.<000000> para consulta"
    set color to w,n/bg,,,n/w
    save screen to telafor
    @ 20,57 get wcodigo_est pict "999999"
    read
    if readkey() = 12 .or. lastkey() = 27
        libera()
        close all
        return
    endif
    if empty(wcodigo_est)
        exit
    endif
    if wcodigo_est = "000000"
        select 1
        save screen to tela_0
        do pes_est
        restore screen from tela_0
        iguala()
    endif
    ycodigo_est = wcodigo_est
    set color to /w
    @ 20,57 say wcodigo_est pict "999999"
    select 1
    set index to ind01.est
    seek wcodigo_est
    if !found()
        do mensagem with "Codigo nao cadastrao ..",8
        loop
    endif
    iguala()
    zcodigo_est = wcodigo_est
    zdescricao  = wdescricao
    set color to /w
    @ 20,65 say wdescricao pict "@s10"
    exit
enddo

select 3
if empty(wcodigo_est)
    index on dtos(data) to &zmicro for data >= wdata_i .and. data <=  wdata_f
else
    index on dtos(data) to &zmicro for data >= wdata_i .and. data <=  wdata_f .and.;
    codigo_est = wcodigo_est
endif
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
    @ prow()+1,02     say "Relatorio ENTRADA MERCADORIAS NO ESTOQUE"
    if !empty(zcodigo_est)
        @ prow()+2,02 say "Codigo Estoque..:"
        @ prow(),pcol()+1 say zcodigo_est pict "@!"
        @ prow(),pcol()+1 say zdescricao  pict "@!"
    endif
    @ prow()+1,02 say "Periodo.........:"
    @ prow(),pcol()+1 say wdata_i pict "@d"
    @ prow(),pcol()+1 say " a "
    @ prow(),pcol()+1 say wdata_f pict "@d"
    @ prow()+2,02     say "Data       Codigo  Mercadoria           Quant.   NF         Preco  "
    @ prow()+1,02     say "~~~~~~~~~~ ~~~~~~~ ~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~ ~~~~~~~~~~~ ~~~~~~~~~"
    pagina= 0
    do while prow()<55 .and. !eof()
        iguala()
        select 1
        seek wcodigo_est
        iguala()
        select 3
        @ prow()+1,02 say wdata       pict "@d"
        @ prow(),13   say wcodigo_est pict "999999"
        @ prow(),21   say wdescricao  pict "@s19"
        @ prow(),42   say wquantidade pict "@e@z 99999.99"
        @ prow(),51   say wnf         pict "@!"
        wvalor = wquantidade * wpreco_com
        @ prow(),63   say wvalor      pict "@e@z 99999.99"
        wlin = 1
        pagina= pagina +  wvalor
        skip
        esc = inkey()
        if esc = 12
            @ prow()+2,40 say "***********  Impressao Cancelada..."
            close all
            return
        endif
    enddo
    geral = geral + pagina
    @ prow()+1,01  say "-------------------------------------------------------------------------"
    @ prow()+1,03  say "Total na Pagina...."
    @ prow(),63    say pagina pict "@e@z 999,999.99"
    @ prow()+1,03  say "Total a Geral......"
    @ prow(),63    say geral  pict "@e@z 999,999.99"
    @ prow()+1,01  say "-------------------------------------------------------------------------"
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
lertexto(zdireto)
close all
return
*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
