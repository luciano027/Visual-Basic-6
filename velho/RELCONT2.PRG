***************************************************************************
* Programa Tabela de Precos - Relatorio                                   *
* Emite relatorio geral (tabela de precos) indexado por espec_est         *
* Luciano 26/08/93                                                        *
***************************************************************************
select 1
if !net_use("apagar",.f.,"ind01.apa")
    return
endif
select 2
if !net_use("fornece",.t.,"ind01.for","ind02.for")
    return
endif

do while .t.
    set color to 0/7
    wcopia = 0
    wmes   = space(01)
    wdata_i= date()
    wdata_f= date()
    g      = 0
    set color to n/w
    dt = quadro(18,40,20,78)
    @ 19,42 say  "Periodo..:            a  "
    set color to w,n/bg,,,n/w
    @ 19,53 get wdata_i  pict "@d"
    read
    if readkey()=12 .or. lastkey() = 27
        close all
        return
    endif
    set color to w,n/bg,,,n/w
    @ 19,66 get wdata_f  pict "@d"
    read
    select 1
    index on codigo_for to ind05.apa for;
    data_p1 != ctod("  /  /  ") .or. data_p2 != ctod("  /  /  ") .or.;
    data_p3 != ctod("  /  /  ") .or. data_p4 != ctod("  /  /  ") .or.;
    data_p5 != ctod("  /  /  ") .or. data_p6 != ctod("  /  /  ")
    set index to ind05.apa
    go top
    limpa()
    set color to w+/r
    dt = quadro(15,18,17,35)
    set color to *+w/r
    @ 16,20 say "Aguarde..."
    set print to &zdireto
    set devi to print
    set cons off
    wpag = 1
    totav = 0
    totap = 0

    do while inkey()#27 .and. !eof()
        iguala()
        @ prow()+1,02     say zempresa  pict "@!"
        @ prow()+1,02     say zendereco pict "@!"
        @ prow()+1,02     say "Telefone: "
        @ prow(),pcol()+1 say ztelefone pict "@!"
        @ prow(),pcol()+1 say "FAX: "
        @ prow(),pcol()+1 say ztelefone pict "@!"
        @ prow(),65       say "Pag."
        @ prow(),72       say strzero(wpag,3,0)  pict "999"
        @ prow()+1,03     say "Relatorio CONTAS PAGAS"
        @ prow()+1,03     say "Emissao :"
        @ prow(),pcol()   say date()    pict "@d"
        @ prow(),pcol()+2 say time()    pict "99:99:99"
        @ prow()+1,03  say "-------------------------------------------------------------------------"
        @ prow()+1,03     say "Periodo..:          a"
        @ prow(),13       say wdata_i pict "@d"
        @ prow(),25       say wdata_f pict "@d"
        select 2
        seek wcodigo_for
        iguala()
        @ prow()+1,03     say "Fornecedor:"
        @ prow(),pcol()+1 say wcodigo_for pict "@!"
        @ prow(),pcol()+1 say wfornecedor pict "@!"
        @ prow()+1,03  say "-------------------------------------------------------------------------"
        @ prow()+2,03     say "NF          Vencimento  Valor       Data Pgto  Valor Pago "
        @ prow()+1,03     say "~~~~~~~~~~~ ~~~~~~~~~~~ ~~~~~~~~~~~ ~~~~~~~~~~ ~~~~~~~~~~~"
        select 1
        zcodigo_for = wcodigo_for
        do while prow()<55 .and. !eof() .and. zcodigo_for = wcodigo_for
            iguala()
            if  wdata_p1 >= wdata_i .and. wdata_p1 <= wdata_f
                @ prow()+1,03   say wnf         pict "@!"
                @ prow(),15     say wdata_v1    pict "@d"
                @ prow(),27     say wvalor_v1   pict "@e@z 999,999.99"
                @ prow(),39     say wdata_p1    pict "@d"
                @ prow(),51     say wvalor_p1   pict "@e@z 999,999.99"
                totav = totav + wvalor_v1
                totap = totap + wvalor_p1
            endif
            if  wdata_p2 >= wdata_i .and. wdata_p2 <= wdata_f
                @ prow()+1,03   say wnf         pict "@!"
                @ prow(),15     say wdata_v2    pict "@d"
                @ prow(),27     say wvalor_v2   pict "@e@z 999,999.99"
                @ prow(),39     say wdata_p2    pict "@d"
                @ prow(),51     say wvalor_p2   pict "@e@z 999,999.99"
                totav = totav + wvalor_v2
                totap = totap + wvalor_p2
            endif
            if  wdata_p3 >= wdata_i .and. wdata_p3 <= wdata_f
                @ prow()+1,03   say wnf         pict "@!"
                @ prow(),15     say wdata_v3    pict "@d"
                @ prow(),27     say wvalor_v3   pict "@e@z 999,999.99"
                @ prow(),39     say wdata_p3    pict "@d"
                @ prow(),51     say wvalor_p3   pict "@e@z 999,999.99"
                totav = totav + wvalor_v3
                totap = totap + wvalor_p3
            endif
            if  wdata_p4 >= wdata_i .and. wdata_p4 <= wdata_f
                @ prow()+1,03   say wnf         pict "@!"
                @ prow(),15     say wdata_v4    pict "@d"
                @ prow(),27     say wvalor_v4   pict "@e@z 999,999.99"
                @ prow(),39     say wdata_p4    pict "@d"
                @ prow(),51     say wvalor_p4   pict "@e@z 999,999.99"
                totav = totav + wvalor_v4
                totap = totap + wvalor_p4
            endif
            if  wdata_p5 >= wdata_i .and. wdata_p5 <= wdata_f
                @ prow()+1,03   say wnf         pict "@!"
                @ prow(),15     say wdata_v5    pict "@d"
                @ prow(),27     say wvalor_v5   pict "@e@z 999,999.99"
                @ prow(),39     say wdata_p5    pict "@d"
                @ prow(),51     say wvalor_p5   pict "@e@z 999,999.99"
                totav = totav + wvalor_v5
                totap = totap + wvalor_p5
            endif
            if  wdata_p6 >= wdata_i .and. wdata_p6 <= wdata_f
                @ prow()+1,03   say wnf         pict "@!"
                @ prow(),15     say wdata_v6    pict "@d"
                @ prow(),27     say wvalor_v6   pict "@e@z 999,999.99"
                @ prow(),39     say wdata_p6    pict "@d"
                @ prow(),51     say wvalor_p6   pict "@e@z 999,999.99"
                totav = totav + wvalor_v6
                totap = totap + wvalor_p6
            endif
            select 1
            skip
            iguala()
            if wcodigo_for != zcodigo_for .or. eof()
                @ prow()+1,03  say "-------------------------------------------------------------------------"
                @ prow()+1,03  say "Total .........."
                @ prow(),27    say totav  pict "@e@z 999,999.99"
                @ prow(),51    say totap  pict "@e@z 999,999.99"
                @ prow()+1,03  say "-------------------------------------------------------------------------"
                totav = 0
                totap = 0
                exit
            endif
        enddo
        wpag = wpag + 1
        @ prow()+1,01  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        @ prow()+1,01  say "              Sistema de Administracao Estoque - Versao 1.0"
        eject
    enddo
    exit
enddo
set cons on
set devi to screen
set print off
set print to
verfil(zdireto)
close all
return
