*****************************************************************************
* Programa...: RELPAG1.PRG                                                  *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......:                   Data da Ultima Atualizacao:                *
* Objetivo...: RELATORIO CONTAS A PAGAR - PERIODO                           *
* Sistema....: SISPAN - Sistema de Panifica??o                              *
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
    wdata_i = date()
    wdata_f = date()
    g      = 0
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
    index on dtos(venci) to &zmicro for pagto >= wdata_i .and. pagto <= wdata_f .and. ;
    pagto != ctod("  /  /  ")
    set index to &zmicro
    go top
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
    total = 0
    do while inkey()#27 .and. !eof()
        @ prow()+1,02     say zempresa  pict "@!"
        @ prow()+1,02     say zendereco pict "@!"
        @ prow()+1,02     say "Telefone: "
        @ prow(),pcol()+1 say ztelefone pict "@!"
        @ prow(),pcol()+1 say "FAX: "
        @ prow(),pcol()+1 say ztelefone pict "@!"
        @ prow(),70   say "Pag."
        @ prow(),75   say strzero(wpag,3,0) pict "999"
        @ prow()+1,02  say "Relatorio Data de Pagamento "
        @ prow()+1,02  say "Periodo..:          a"
        @ prow(),13    say wdata_i pict "@d"
        @ prow(),25    say wdata_f pict "@d"
        @ prow()+2,03     say "Documento  Parcela empresa                     Tipo   Data Venc. Valor   "
        @ prow()+1,03     say "~~~~~~~~~~ ~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~  ~~~~~~~~~~ ~~~~~~~~"
        do while prow()<50 .and. !eof()
            iguala()
            select 1
            set index to ind01.for
            seek wcodigo_for
            iguala()
            @ prow()+1,03   say wdocumento  pict "@!"
            @ prow(),15     say wparcela    pict "@!"
            @ prow(),22     say wfornecedor pict "@s20"
            @ prow(),51     say wtipo       pict "@!"
            @ prow(),58     say wvenci      pict "@d"
            @ prow(),68     say wvalor_v    pict "@e@z 99,999.99"
            total = total + wvalor_v
            select 3
            skip
        enddo
        wpag = wpag + 1
        @ prow()+1,03  say "-------------------------------------------------------------------------"
        @ prow()+1,03  say "                  Total ......................."
        @ prow(),68    say total pict "@e@z 99,999.99"
        @ prow()+1,03  say "-------------------------------------------------------------------------"
        @ prow()+1,03  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        @ prow()+1,01  say "              Sistema de Administracao Estoque - Versao 1.0"
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
