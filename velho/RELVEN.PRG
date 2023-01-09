***************************************************************************
* Programa Tabela de Precos - Relatorio                                   *
*                                                                         *
* Emite relatorio geral (tabela de precos) indexado por espec_est         *
*                                                                         *
* Luciano 26/08/93                                                        *
*                                                                         *
***************************************************************************
select 2
use empresa
locate for !empty(empresa)
if found()
    wempresa  = empresa
    wendereco = endereco
else
    do mensagem with "Empresa nao Cadastrada ...",8
    return
endif
select 1
use vendedor
if !file("ind01.ven") .or. !file("ind02.ven")
    index on cod_ven   to ind01.ven
    index on nome_v      to ind02.ven
endif
set index to ind01.ven,ind02.ven
limpa()
set color to w+/r
dt = quadro(15,18,17,35)
set color to *+w/r
@ 16,20 say "Aguarde..."
set print to arq.txt
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
    @ prow(),75       say strzero(wpag,3,0)  pict "999"
    @ prow()+1,03     say "Relatorio VENDEDORES CADASTRADOS "
    @ prow()+2,03     say "  Codigo       Vendedor                                Telefone    Celular    "
    @ prow()+1,03     say "  ~~~~~~~~~~~  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~~ ~~~~~~~~~~~"
    do while prow()<55 .and. !eof()
        iguala()
        @ prow()+1,05 say wcod_ven    pict "@!"
        @ prow(),18   say wnome_v     pict "@s32"
        @ prow(),60   say wtelefone   pict "@!"
        @ prow(),73   say wcelular    pict "@!"
        wlin = 1
        skip
    enddo
    if eof()
        do while prow()<55
            @ prow()+1,00 say ""
        enddo
    endif
    @ prow()+1,01  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
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
verfil("arq.txt")
close all
return
***************************************************************************
*                        f i m                                            *
***************************************************************************
