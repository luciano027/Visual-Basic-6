*****************************************************************************
* Programa...: RELCONTA.PRG                                                 *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao:                *
* Objetivo...: RELATORIO DE CONTAS A PAGAR - GERAL                          *
* Sistema....: SISPAN - Sistema de Panifica‡Æo                              *
*****************************************************************************
select 1
if !net_use("fornece",.f.,"ind01.for","ind02.for")
    return
endif

select 2
if !net_use("tipo",.f.,"ind01.tip")
    return
endif

select 3
if !net_use("apagar",.f.,"ind01.apa")
    return
endif
do while .t.
    set color to 0/7
    wcopia = 0
    wmes   = space(01)
    wdata_i= date()
    wdata_f= date()
    wfornec= space(10)
    g      = 0
    save screen to telarel
    set color to n/w
    dt = quadro(18,40,20,78)
    @ 19,42 say "Periodo..:           a  "
    set color to w,n/bg,,,n/w
    @ 19,52 get wdata_i  pict "@d"
    read
    if readkey()=12 .or. lastkey() = 27
        close all
        return
    endif
    set color to w,n/bg,,,n/w
    @ 19,65 get wdata_f  pict "@d"
    read
    select 3
    index on dtos(venci) to &zmicro for venci >= wdata_i .and. venci <= wdata_f
    set index to &zmicro
    go top
    restore screen from telarel
    limpa()
    set color to w+/r
    dt = quadro(15,18,17,35)
    set color to *+w/r
    @ 16,20 say "Aguarde..."
    set color to w/b
    set print to &zdireto
    set devi to print
    set cons off
    wpag  = 1
    total = 0
    tota  = 0
    do while inkey()#27 .and. !eof()
        @ prow()+1,02     say zempresa  pict "@!"
        @ prow()+1,02     say zendereco pict "@!"
        @ prow()+1,02     say "Telefone: "
        @ prow(),pcol()+1 say ztelefone pict "@!"
        @ prow(),pcol()+1 say "FAX: "
        @ prow(),pcol()+1 say ztelefone pict "@!"

        @ prow(),70    say "Pag."
        @ prow(),75    say strzero(wpag,3,0) pict "999"
        @ prow()+1,02  say "Relatorio de Contas a Pagar - GERAL"
        @ prow()+1,02  say "Periodo..:          a"
        @ prow(),13    say wdata_i pict "@d"
        @ prow(),25    say wdata_f pict "@d"
        @ prow()+1,02  say "Documento   Tipo          Fornecedor     Vencimento Valor      pago em"
        @ prow()+1,02  say "~~~~~~~~~~~ ~~~~~~~~~~~~~ ~~~~~~~~~~~~~~ ~~~~~~~~~~ ~~~~~~~~~~ ~~~~~~~~~"
        do while prow()<55 .and. !eof()
            iguala()
            select 2
            set index to ind01.tip
            seek wtipo
            iguala()
            select 1
            set index to ind01.for
            seek wcodigo_for
            iguala()
            wfornece = wfornecedor
            @ prow()+1,02   say wdocumento pict "@!"
            @ prow(),14     say wdescricao pict "@s13"
            @ prow(),28     say wfornece   pict "@s13"
            @ prow(),43     say wvenci     pict "@d"
            if !empty(wpagto)
                total = total + wvalor_p
                @ prow(),54     say wvalor_p   pict "@e 99,999.99"
            else
                tota  = tota  + wvalor_v
                @ prow(),54     say wvalor_v   pict "@e 99,999.99"
            endif
            @ prow(),65     say wpagto   pict "@d"
            select 3
            skip
        enddo
        @ prow()+1,02  say repl("-",78)
        @ prow()+1,02  say "   Total de Contas Pagas....:"
        @ prow(),54    say total pict "@e 99,999.99"
        @ prow()+1,02  say repl("-",78)
        @ prow()+1,02  say "   Total de Contas a Pagar..:"
        @ prow(),54    say tota  pict "@e 99,999.99"
        @ prow()+1,02  say repl("-",78)
        if eof()
            do while prow()<55
                @ prow()+1,00 say ""
            enddo
        endif
        @ prow()+2,01  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        @ prow()+1,01  say "             Sistema de Estoque e Vendas   - Versao 2.1"
        @ prow(),02    say date() pict "@d"
        @ prow(),69    say time() pict "99:99"
        wpag = wpag +1
        eject
    enddo
    exit
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
