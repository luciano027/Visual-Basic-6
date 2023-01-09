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
if !net_use("cticket",.f.)
    return
endif
select 4
if !net_use("vendedor",.f.,"ind01.ven","ind02.ven")
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
        select 4
        save screen to tela_d
        do pes_ven
        restore screen from tela_d
        iguala()
    endif
    set color to n/w
    @ 18,57 say wcodigo_ven pict "999"
    select 4
    set index to ind01.ven
    seek wcodigo_ven
    if !found()
        do mensagem with "Vendedor nao cadastrado...",8
        loop
    endif
    iguala()
    ynome_ven = wnome_ven
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
index on dtos(data_ti)+historico to &zmicro for data_ti >= wdata_i .and. data_ti <=  wdata_f  .and.;
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
    @ prow()+1,02     say zempresa pict "@!"
    @ prow(),70       say "Pag."
    @ prow(),75       say strzero(wpag,3,0) pict "999"
    @ prow()+1,02     say "Relatorio - Recebimentos"
    @ prow()+1,02     say "Vendedor:"
    @ prow(),pcol()   say  ynome_ven pict "@!"
    @ prow()+1,02     say "Periodo :"
    @ prow(),pcol()+1 say wdata_i pict "@d"
    @ prow(),pcol()+1 say " a "
    @ prow(),pcol()+1 say wdata_f pict "@d"
    @ prow()+2,02     say "Data Pgto  Descricao                                       Valor"
    @ prow()+1,02     say "~~~~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~~~~"
    pagina = 0
    do while prow()<45 .and. !eof()
        iguala()
        zhistorico = subst(whistorico,1,25)
        yhistorico = zhistorico
        qhistorico = whistorico
        wtotalt = 0
        zdata_ti = wdata_ti
        do while whistorico = yhistorico
            iguala()
            wtotalt = wtotalt + wtotal_ti
            skip
            iguala()
            zhistorico = subst(whistorico,1,25)
        enddo
        @ prow()+1,02 say zdata_ti    pict "@d"
        @ prow(),13   say qhistorico  pict "@!"
        @ prow(),62   say wtotalt     pict "@e@z 99,999.99"
        pagina = pagina + wtotalt
    enddo
    geral = geral + pagina
    @ prow()+1,02  say "----------------------------------------------------------------------------"
    @ prow()+1,03 say "Total na Pagina..:"
    @ prow(),pcol()+1 say pagina      pict "@e@z 99,999.99"
    if eof()
        @ prow()+1,03 say "Total Geral.....:"
        @ prow(),pcol()+1 say geral       pict "@e@z 99,999.99"
    endif
    @ prow()+1,02  say "----------------------------------------------------------------------------"
    @ prow()+2,02  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    @ prow()+1,02  say "              Sistema de Administracao Estoque - Versao 5.3"
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
