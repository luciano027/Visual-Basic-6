*****************************************************************************
* Programa...: RELENT.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao:                *
* Objetivo...: RELATORIO DE ENTRADA DE MERCADORIAS NO ESTOQUE               *
* Sistema....: SISPAN - Sistema de Panifica‡Æo                              *
*****************************************************************************
select 1
if !net_use("caixa",.f.,"ind01.cai")
    return
endif

set color to 0/7
wcopia = 0
wmes   = space(01)
wdata_i= date()
wdata_f= date()
g      = 0
save screen to telafor
set color to n/w
dt = quadro(18,40,20,78)
@ 19,42 say "Periodo....:           a  "
set color to w,n/bg,,,n/w
@ 19,54 get wdata_i  pict "@d"
read
if readkey() = 12 .or. lastkey()=27
    close all
    return
endif
@ 19,67 get wdata_f  pict "@d"
read

select 1
index on dtos(caixa)+tipo_caixa to indrel.ent for caixa >= wdata_i .and. caixa <=  wdata_f
set index to indrel.ent

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
gera_a = 0
gera_r = 0
do while inkey()#27 .and. !eof()
    @ prow()+1,02     say zempresa  pict "@!"
    @ prow()+1,02     say zendereco pict "@!"
    @ prow()+1,02     say "Telefone: "
    @ prow(),pcol()+1 say ztelefone pict "@!"
    @ prow(),pcol()+1 say "FAX: "
    @ prow(),pcol()+1 say ztelefone pict "@!"

    @ prow(),70       say "Pag."
    @ prow(),75       say strzero(wpag,3,0) pict "999"
    @ prow()+1,02     say "Relatorio Fluxo Caixa"
    @ prow()+1,02     say "Periodo :"
    @ prow(),pcol()+1 say wdata_i pict "@d"
    @ prow(),pcol()+1 say " a "
    @ prow(),pcol()+1 say wdata_f pict "@d"
    @ prow()+2,02     say"Data       Historico         Dinheiro     Chq/Cart   A Pagar    A Receber "
    @ prow()+1,02     say"~~~~~~~~~~ ~~~~~~~~~~~~~~~~~ ~~~~~~~~~~~~ ~~~~~~~~~~ ~~~~~~~~~~ ~~~~~~~~~~"
    store 0 to pagina_a,pagina_r
    do while prow()<55 .and. !eof()
        iguala()
        @ prow()+1,02 say wcaixa      pict "@d"
        if wtipo_caixa = "A"
            @ prow(),13   say "A Pagar...."
            @ prow(),55   say wdinheiro   pict "@e@z 999,999.99"
            pagina_r = pagina_r + wdinheiro
        elseif wtipo_caixa = "R"
            @ prow(),13   say "A Receber.."
            @ prow(),66   say wdinheiro+wcheque_pre   pict "@e@z 999,999.99"
            pagina_a = pagina_a + (wdinheiro+wcheque_pre)
        elseif wtipo_caixa = "C"
            @ prow(),13   say "Caixa do Dia.."
            @ prow(),31   say wdinheiro       pict "@e@z 999,999.99"
            @ prow(),44   say wcartao+wcheque pict "@e@z 999,999.99"
            pagina_a = pagina_a + (wdinheiro+wcartao+wcheque)
        endif
        skip
        esc = inkey()
        if esc = 12
            @ prow()+2,40 say "***********  Impressao Cancelada..."
            close all
            return
        endif
    enddo
    gera_a = gera_a + pagina_a
    gera_r = gera_r + pagina_r
    @ prow()+1,01  say "----------------------------------------------------------------------------"
    @ prow()+1,01  say "Total no Caixa.R$"
    wtotal_c = gera_a - gera_r
    @ prow(),20    say wtotal_c pict "@e@z 99,999.99"
    @ prow()+1,01  say "----------------------------------------------------------------------------"
    @ prow()+2,01  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
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
verfil(zdireto)
close all
return
*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
