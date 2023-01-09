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
if !net_use("clientes",.f.,"ind01.cli","ind02.cli","ind03.cli")
    return
endif


set color to 0/7
wcopia = 0
wmes   = space(01)
wdata_i= date()
wdata_f= date()
wcodigo_cli = space(05)
zcodigo_cli = space(05)
g      = 0
save screen to telafor
set color to n/w
dt = quadro(18,40,20,78)
@ 19,42 say "Codigo Cliente:"
do while .t.
    set color to w/b
    @ 23,05 say "<ESC> Para sair.<00000> para consulta"
    set color to w,n/bg,,,n/w
    save screen to telafor
    @ 19,57 get wcodigo_cli pict "99999" valid !empty(wcodigo_cli)
    read
    if readkey() = 12 .or. lastkey() = 27
        libera()
        close all
        return
    endif
    if wcodigo_cli = "00000"
        select 3
        save screen to tela_0
        do pes_cli
        restore screen from tela_0
        iguala()
    endif
    ycodigo_cli = wcodigo_cli
    set color to /w
    @ 19,57 say wcodigo_cli pict "99999"
    select 3
    set index to ind01.cli
    seek wcodigo_cli
    if !found()
        do mensagem with "Codigo nao cadastrado ..",8
        loop
    endif
    iguala()
    zcodigo_cli = wcodigo_cli
    znome       = wnome
    set color to /w
    @ 19,65 say wnome      pict "@s10"
    exit
enddo
select 2
index on dtos(data) to &zmicro for codigo_cli = wcodigo_cli
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
it = 8
setprc(0,0)
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
    @ prow()+2,02 say "Codigo Cliente..:"
    @ prow(),pcol()+1 say zcodigo_cli pict "@!"
    @ prow(),pcol()+1 say znome       pict "@!"
    @ prow()+2,02     say "Data       Mercadoria                         Quant.    Preco      Total"
    @ prow()+1,02     say "~~~~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~  ~~~~~~~~  ~~~~~~~~~~"
    pagina = 0             
    do while it < 8 .and. !eof()
        iguala()
        select 1
        seek wcodigo_est
        iguala()
        select 2
        @ prow()+1,02 say wdata       pict "@d"    
        @ prow(),13   say wdescricao  pict "@s34"
        @ prow(),48   say wquantidade pict "@e@z 99999.99"
        @ prow(),57   say wpreco_ven  pict "@e@z 99,999.99"
        wvalor = wquantidade * wpreco_ven
        @ prow(),68   say wvalor      pict "@e@z 99,999.99"
        wlin = 1
        pagina = pagina + wvalor
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
    wpag = wpag + 1
    if !eof()
       tsalto = zsalto + (8 - it)
       @ prow()+tsalto,01 say "."
       it = 0
       setprc(0,0)
    endif
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
