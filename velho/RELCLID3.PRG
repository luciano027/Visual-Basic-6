*****************************************************************************
* Programa...: RELCLI.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao:                *
* Objetivo...: Relatorio de Clientes                                        *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************
select 1
if !net_use("estoque",.f.,"ind01.est","ind02.est")
    return
endif

select 2
if !net_use("clientes",.f.,"ind01.cli","ind02.cli","ind03.cli")
    return
endif

select 3
if !net_use("prazo",.f.,"ind01.pra")
    return
endif

select 4
if !net_use("csaida",.f.,)
    return
endif

select 5
if !net_use("cticket",.f.,)
    return
endif

save screen to telagu
set color to w+/r
dt = quadro(15,18,17,35)
set color to *+w/r
@ 16,20 say "Aguarde..."


select 2
iguala()
do while !eof()
    iguala()
    select 3
    set index to ind01.pra
    seek wcodigo_cli
    if found()
        iguala()
        ycodigo_cli = wcodigo_cli
        wtotal = 0
        do while wcodigo_cli = ycodigo_cli
            select 1
            set index to ind01.est
            seek wcodigo_est
            iguala()
            wprec = wvalor
            wcompra = (wprec * wquantidade)
            if alltrim(wdescricao,1,5) != "CREDI"
                wtotal = (wtotal + (wvalor * wquantidade))  &&----> imprime o total
            endif
            select 3
            skip
            iguala()
        enddo
        select 2
        set index to ind01.cli
        seek ycodigo_cli
        if found()
            if rec_lock()
                replace debito with wtotal
                unlock
            endif
        endif
    endif
    select 2
    skip
enddo


select 2
index on nome to &zmicro for debito > 0 .and. data_deb+30 < date()
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
    @ prow(),75       say strzero(wpag,3,0) pict "999"
    @ prow()+1,02     say "Relatorio Clientes    "
    @ prow()+2,02     say "Data Debito Codigo   Nome                                       Debito"
    @ prow()+1,02     say "~~~~~~~~~~~ ~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~~~"
    pagina = 0
    do while prow()<55 .and. !eof()
        iguala()
        @ prow()+1,02   say wdata_deb        pict "@d"
        @ prow(),14     say wcodigo_cli      pict "@!"
        @ prow(),22     say wnome            pict "@s20"
        @ prow(),66     say wdebito-wcredito pict "@e@z 9,999.99"
        @ prow()+1,02     say "----------------------------------------------------------------------------"
        wlin = 1
        skip
        esc = inkey()
        if esc = 12
            @ prow()+2,40 say "***********  Impressao Cancelada..."
            close all
            return
        endif
        pagina = pagina + (wdebito-wcredito)
    enddo
    geral = geral + pagina
    if eof()
        do while prow()<55
            @ prow()+1,00 say ""
        enddo
    endif
    @ prow()+1,02     say "----------------------------------------------------------------------------"
    @ prow()+1,02     say "Total na Pagina..:"
    @ prow(),pcol()+1 say pagina pict "@e@z 99,999.99"
    @ prow()+1,02     say "Total Geral......:"
    @ prow(),pcol()+1 say geral  pict "@e@z 99,999.99"
    @ prow()+2,02     say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    @ prow()+1,01  say "                        Sistema de Vendas  - Versao 1.0"
    @ prow(),02    say date() pict "@d"
    @ prow(),69    say time() pict "99:99"
    wpag = wpag +1
    eject
enddo
restore screen from telagu
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
