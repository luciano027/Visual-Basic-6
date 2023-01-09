*****************************************************************************
* Programa...: RELPROD.PRG                                                  *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......:                   Data da Ultima Atualizacao:                *
* Objetivo...: RELATORIO DE PRODUTOS CADASTRADOS                            *
* Sistema....: SISPAN - Sistema de Panifica‡Æo                              *
*****************************************************************************
select 1
if !net_use("estoque",.f.,"ind01.est","ind02.est")
    return
endif
select 2
if !net_use("fornece",.f.,"ind01.for","ind02.for")
    return
endif

do while .t.
    set color to 0/7
    wcodigo_for = "000"
    set color to n/w
    dt = quadro(17,40,19,70)
    @ 18,42 say "Fornecedor :"
    set color to w,n/bg,,,n/w
    @ 18,54 get wcodigo_for pict "999"
    read
    if readkey()=27 .or. lastkey() = 12
        close all
        return
    endif
    if wcodigo_for = "000"
        select 2
        save screen to tela_0
        do pes_for
        restore screen from tela_0
        iguala()
    endif
    ycodigo_for = wcodigo_for
    set color to /w
    @ 18,54 say ycodigo_for pict "999"
    select 2
    set index to ind01.for
    seek ycodigo_for
    if !found()
        do mensagem with "Fornecedor nao cadastrado...",8
        loop
    endif
    iguala()
    set color to w/r
    @ 18,60 say wfornecedor pict "@s10"
    exit
enddo

select 1
index on descricao   to &zmicro for codigo_for = wcodigo_for .and. saldo <= minino
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
do while inkey()#27 .and. !eof()
    @ prow()+1,02     say zempresa  pict "@!"
    @ prow()+1,02     say zendereco pict "@!"
    @ prow()+1,02     say "Telefone: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),pcol()+1 say "FAX: "
    @ prow(),pcol()+1 say ztelefone pict "@!"

    @ prow(),70       say "Pag."
    @ prow(),75       say strzero(wpag,3,0) pict "999"
    @ prow()+1,02     say "Relatorio PEDIDO"
    @ prow()+2,02     say "Fornecedor.:"
    @ prow(),pcol()+1 say wfornecedor pict "@!"
    @ prow()+1,02     say "Telefone...:"
    @ prow(),14       say wtelefone pict "@!"
    @ prow(),43       say "FAX..:"
    @ prow(),50       say wfax      pict "@!"
    @ prow()+1,02     say "C.G.C......:"
    @ prow(),14       say wcgc        pict "@!"
    @ prow(),43       say "Insc.Es:"
    @ prow(),50       say winscr      pict "@!"
    @ prow()+2,02     say "Codigo  Produto                         Uni  Saldo      Minimo    Pedido"
    @ prow()+1,02     say "~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~ ~~~~~~~~~~ ~~~~~~~~~ ~~~~~~~"
    do while prow()<55 .and. !eof()
        iguala()
        @ prow()+1,02 say wcodigo_est  pict "999999"
        @ prow(),10   say wdescricao   pict "@s30"
        @ prow(),42   say wunidade     pict "@!"
        @ prow(),47   say wsaldo       pict "@e@z 99,999.99"
        @ prow(),58   say wminino      pict "@e@z 99,999.99"
        @ prow()+1,02 say "----------------------------------------------------------------- -------"
        skip
        esc = inkey()
        if esc = 12
            @ prow()+2,40 say "***********  Impressao Cancelada..."
            close all
            return
        endif
    enddo
    if eof()
        do while prow()<55
            @ prow()+1,00 say ""
        enddo
    endif
    @ prow()+2,02  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    @ prow()+1,02  say "              Sistema de Administracao Estoque - Versao 1.0"
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
