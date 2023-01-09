*****************************************************************************
* Programa...: MENUMOV.PRG                                                  *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 02/12/98          Data da Ultima Atualizacao:                *
* Objetivo...: MENU COM OPCOES MOVIMENTO NO ESTOQUE - ENTRADA/SAIDA/CAIXA   *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************

do while .t.
    limpa()
    set color to w/b
    @ 23,05 say "<ESC> Para sair."
    Janela2("W",.F.,03,30,06,45,.F.)
    SetColor("N/W,W+/R")
    @ 04,31 prompt "Venda a Vista "
    @ 05,31 prompt "Venda a Prazo "
    save screen to telabas
    menu to op
    do case
        case op = 1
            save screen to telaop31
            do sinal with "Venda a Vista         "
            close all
            do cadsai1
            close all
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            restore screen from telaop31
            loop
        case op = 2
            save screen to telaop31
            do sinal with "Venda a Prazo         "
            close all
            do cadprazo
            close all
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            restore screen from telaop31
            loop

        other
            close all
            exit
    endcase
enddo
close all
return




















