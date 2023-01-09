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
if !net_use("fornece",.f.,"ind01.for","ind02.for")
    return
endif

do while .t.
    set color to 0/7
    wcopia = 0
    wcodigo_for = space(03)
    wdata_i= date()
    wdata_f= date()
    g      = 0
    set color to n/w
    dt = quadro(19,40,21,78)
    @ 20,42 say  "Fornecedor:"
    do while .t.
        set color to w/b
        @ 23,05 say "<ESC> Para sair. <000> para Consulta"
        set color to w,n/bg,,,n/w
        @ 20,55 get wcodigo_for pict "999"
        read
        if readkey() = 12  .or. lastkey() = 27
            close all
            return
        endif
        if empty(wcodigo_for) .or. wcodigo_for = "000"
            select 2
            save screen to tela_01
            do pes_for
            restore screen from tela_01
            iguala()
        endif
        if empty(wcodigo_for)
            close all
            return
        endif
        set color to /w
        @ 20,55 say wcodigo_for pict "999"
        select 2
        set index to ind01.for
        seek wcodigo_for
        if !found()
            do mensagem with "Codigo do Fornecedor nao cadastrado..",8
            loop
        endif
        iguala()
        set color to w/r
        @ 20,60 say wfornecedor pict "@s10"
        exit
    enddo
    select 1
    index on codigo_for to ind05.apa for  codigo_for = wcodigo_for
    set index to ind05.apa
    go top
    limpa()
    set color to w+/r
    dt = quadro(15,18,17,35)
    set color to *+w/r
    @ 16,20 say "Aguarde..."
    set print to arq.txt
    set devi to print
    set cons off
    wpag = 1
    totav = 0
    totap = 0
    do while inkey()#27 .and. !eof()
        iguala()
        @ prow()+1,02     say zempresa  pict "@!"
        @ prow()+1,02     say zendereco pict "@!"
        @ prow(),65       say "Pag."
        @ prow(),72       say strzero(wpag,3,0)  pict "999"
        @ prow()+1,03     say "Relatorio CONTAS A PAGAR - Geral"
        @ prow()+1,03     say "Emissao :"
        @ prow(),pcol()   say date()    pict "@d"
        @ prow(),pcol()+2 say time()    pict "99:99:99"
        @ prow()+1,03     say "Periodo..:          a"
        @ prow(),13       say wdata_i pict "@d"
        @ prow(),25       say wdata_f pict "@d"
        @ prow()+2,03     say "Fornecedor         NF          Vencimento  Valor       Data Pgto  Valor Pago "
        @ prow()+1,03     say "~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~~ ~~~~~~~~~~~ ~~~~~~~~~~~ ~~~~~~~~~~ ~~~~~~~~~~~"
        select 1
        zcodigo_for = wcodigo_for
        do while prow()<55 .and. !eof()
            iguala()
            select 2
            seek wcodigo_for
            iguala()
            @ prow()+1,03   say wfornecedor pict "@s17"
            @ prow(),22     say wnf         pict "@!"
            @ prow(),34     say wdata_v1    pict "@d"
            @ prow(),46     say wvalor_v1   pict "@e@z 999,999.99"
            @ prow(),58     say wdata_p1    pict "@d"
            @ prow(),69     say wvalor_p1   pict "@e@z 999,999.99"
            if !empty(wdata_v2)
                @ prow()+1,34     say wdata_v2    pict "@d"
                @ prow(),46     say wvalor_v2   pict "@e@z 999,999.99"
                @ prow(),58     say wdata_p2    pict "@d"
                @ prow(),69     say wvalor_p2   pict "@e@z 999,999.99"
            endif
            if !empty(wdata_v3)
                @ prow()+1,34     say wdata_v3    pict "@d"
                @ prow(),46     say wvalor_v3   pict "@e@z 999,999.99"
                @ prow(),58     say wdata_p3    pict "@d"
                @ prow(),69     say wvalor_p3   pict "@e@z 999,999.99"
            endif
            if !empty(wdata_v4)
                @ prow()+1,34     say wdata_v4    pict "@d"
                @ prow(),46     say wvalor_v4   pict "@e@z 999,999.99"
                @ prow(),58     say wdata_p4    pict "@d"
                @ prow(),69     say wvalor_p4   pict "@e@z 999,999.99"
            endif
            if !empty(wdata_v5)
                @ prow()+1,34     say wdata_v5    pict "@d"
                @ prow(),46     say wvalor_v5   pict "@e@z 999,999.99"
                @ prow(),58     say wdata_p5    pict "@d"
                @ prow(),69     say wvalor_p5   pict "@e@z 999,999.99"
            endif
            if !empty(wdata_v6)
                @ prow()+1,34     say wdata_v6    pict "@d"
                @ prow(),46     say wvalor_v6   pict "@e@z 999,999.99"
                @ prow(),58     say wdata_p6    pict "@d"
                @ prow(),69     say wvalor_p6   pict "@e@z 999,999.99"
            endif
            totav = totav + (wvalor_v1+wvalor_v2+wvalor_v3+wvalor_v4+wvalor_v5+wvalor_v6)
            totap = totap + (wvalor_p1+wvalor_p2+wvalor_p3+wvalor_p4+wvalor_p5+wvalor_p6)
            select 1
            skip
            iguala()
            if eof()
                @ prow()+1,03     say "-----------------------------------------------------------------------------"
                @ prow()+1,03  say "Total .........."
                @ prow(),27    say totav  pict "@e@z 999,999.99"
                @ prow(),51    say totap  pict "@e@z 999,999.99"
                @ prow()+1,03     say "-----------------------------------------------------------------------------"
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
verfil("arq.txt")
close all
return
