*****************************************************************************
* Programa...: RELCLI.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao:                *
* Objetivo...: Relatorio de Clientes                                        *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************
select 1
if !net_use("clientes",.f.,"ind02.cli","ind01.cli","ind03.cli")
    return
endif

limpa()
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
    @ prow()+1,02     say "Relatorio Clientes    "
    @ prow()+2,02     say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    do while prow()<55 .and. !eof()
        iguala()
        @ prow()+1,06   say wcodigo_cli pict "99999"
        @ prow(),11     say " - "
        @ prow(),14     say wnome       pict "@s40"
        @ prow()+1,06   say "Endereco.:"
        @ prow(),16     say wendereco pict "@!"
        @ prow()+1,06   say "Cidade...:"
        @ prow(),16     say wcidade pict "@s20"
        @ prow(),41     say "Bairro...:"
        @ prow(),51     say wbairro pict "@s15"
        @ prow(),72     say "UF.:"
        @ prow(),76     say wuf     pict "@!"
        @ prow()+1,06   say "CEP......:"
        @ prow(),16     say wcep pict "@!"
        @ prow(),41     say "Telefone.:"
        @ prow(),51     say wtelefone pict "@!"
        @ prow()+1,06   say "CGC/CPF..:"
        @ prow(),16     say wcgc pict "@!"
        @ prow(),41     say "Insc/CI..:"
        @ prow(),51     say winscr pict "@!"
        @ prow()+1,02     say "----------------------------------------------------------------------------"
        wlin = 1
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
    wpag = wpag +1
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
