***************************************************************************
* Programa Tabela de Precos - Relatorio                                   *
* Emite relatorio geral (tabela de precos) indexado por espec_est         *
* Luciano 26/08/93                                                        *
***************************************************************************
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
    index on dtos(venci) to &zmicro for venci >= wdata_i .and. venci <= wdata_f .and.;
    pagto != ctod("  /  /  ")
    set index to &zmicro
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
    total = 0
    do while inkey()#27 .and. !eof()
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
        @ prow()+2,03     say "Periodo..:          a"
        @ prow(),13       say wdata_i pict "@d"
        @ prow(),25       say wdata_f pict "@d"
        @ prow()+2,03     say "Documento  Parcela Empresa                     Tipo   Data Pgto  Valor Pago"
        @ prow()+1,03     say "~~~~~~~~~~ ~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~  ~~~~~~~~~~ ~~~~~~~~~~"
        pagina = 0
        do while prow()<50 .and. !eof()
            iguala()
            select 1
            set index to ind01.for
            seek wcodigo_for
            iguala()
            @ prow()+1,03   say wboleto     pict "@!"
            @ prow(),15     say wparcela    pict "@!"
            @ prow(),22     say wfornecedor pict "@s20"
            @ prow(),51     say wtipo       pict "@!"
            @ prow(),58     say wpagto      pict "@d"
            @ prow(),68     say wvalor_v    pict "@e@z 999,999.99"
            pagina = pagina + wvalor_v
            select 3
            skip
        enddo
        total = total + pagina
        wpag = wpag + 1
        @ prow()+1,03  say "-------------------------------------------------------------------------"
        @ prow()+1,03  say " Total na Pagina..............................."
        @ prow(),68    say pagina pict "@e@z 999,999.99"
        if eof()
            @ prow()+1,03  say " Total Geral..................................."
            @ prow(),68    say total pict "@e@z 999,999.99"
        endif
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
