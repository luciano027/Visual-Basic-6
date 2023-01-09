*********************************************************************
* Programa cadastro de estnos
* Data : 11/07/94
* Programador : Luciano Moreira Araujo
*********************************************************************
*****> Abre arquivo
select 1
if !net_use("estoque",.f.,"ind01.est","ind02.est")
    return
endif
index on descricao to &zmicro for preco_ven > 0
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
    iguala()
    @ prow()+1,02     say zempresa  pict "@!"
    @ prow()+1,02     say zendereco pict "@!"
    @ prow()+1,02     say "Telefone: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),pcol()+1 say "FAX: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),67       say "Pag."
    @ prow(),72       say strzero(wpag,3,0) pict "999"
    @ prow()+1,01     say "Relatorio TABELA DE PRECO DE VENDA"
    @ prow()+1,01     say "Emissao :"
    @ prow(),pcol()   say date()    pict "@d"
    @ prow(),pcol()+2 say time()    pict "99:99:99"
    @ prow()+2,01     say "Codigo          Descricao                                  UND  Pr. Venda "
    @ prow()+1,01     say "~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~ ~~~~~~~~~~"
    do while prow()<55 .and. !eof()
        iguala()
        wdescricao = alltrim(wdescricao)+repl(".",48-len(alltrim(wdescricao)))
        @ prow()+1,01 say wcodigo_est pict "@!"
        @ prow(),09   say wdescricao  pict "@!"
        @ prow(),60   say wunidade    pict "@!"
        @ prow(),67   say wpreco_ven  pict "@e 9,999.99"
        @ prow()+1,01 say "--------------------------------------------------------------------------"
        select 1
        wlin = 1
        skip
        iguala()
    enddo
    wpag = wpag + 1
    if eof()
        do while prow()<55
            @ prow()+1,00 say ""
        enddo
    endif
    @ prow()+1,03  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    @ prow()+1,02  say "              Sistema de Administracao Estoque - Versao 1.0"
    eject
enddo
set cons on
set devi to screen
set print off
set print to
lertexto(zdireto)
close all
return
***************************************************************************
*                        f i m                                            *
***************************************************************************
