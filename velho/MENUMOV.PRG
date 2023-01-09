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
    Janela2("W",.F.,03,16,10,39,.F.)
    SetColor("N/W,W+/R")
    @ 04,17 prompt "Entrada no Estoque   "+chr(016)
    @ 05,17 prompt "Saida no Estoque     "+chr(016)
    @ 06,17 prompt "Caixa                "+chr(016)
    @ 07,17 prompt "Contas a Pagar       "+chr(016)
    @ 08,17 prompt "Contas a Receber     "+chr(016)
    @ 09,17 prompt "Boleto Bancario      "
    save screen to telabas
    menu to op
    do case
        case op = 1
            save screen to telaop31
            do sinal with "Entrada no Estoque    "
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            do while .t.
                set color to w/b
                @ 23,05 say "<ESC> Para sair."
                Janela2("W",.F.,05,39,08,70,.F.)
                SetColor("N/W,W+/R")
                @ 06,40 prompt "Entrada no Estoque           "
                @ 07,40 prompt "Relatorio Entrada no Periodo "
                menu to op1
                do case
                    case op1 = 1
                        save screen to tela_ca1
                        do sinal with "Entrada no Estoque    "
                        do cadent
                        close all
                        restore screen from tela_ca1
                        loop
                    case op1 = 2
                        save screen to tela_ca1
                        do sinal with "Entrada no Periodo    "
                        do relent
                        close all
                        restore screen from tela_ca1
                        loop
                    other
                        exit
                endcase
            enddo
            restore screen from telaop31
            loop
        case op = 2
            save screen to telaop32
            do sinal with "Saida no Estoque      "
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            do while .t.
                set color to w/b
                @ 23,05 say "<ESC> Para sair."
                Janela2("W",.F.,06,39,11,67,.F.)
                SetColor("N/W,W+/R")
                @ 07,40 prompt "Saida no Estoque           "
                @ 08,40 prompt "Relatorio Saida Geral      "
                @ 09,40 prompt "Relatorio Saida Cliente    "
                @ 10,40 prompt "Relatorio Saida Vendedor   "
                save screen to telarel1
                menu to op2
                do case
                    case op2 = 1
                        save screen to tela_ca1
                        do sinal with "Saida no Estoque      "
                        do cadsai
                        close all
                        restore screen from tela_ca1
                        loop
                    case op2 = 2
                        save screen to tela_ca1
                        do sinal with "Saida no Estoque Perio"
                        do relsai
                        close all
                        restore screen from tela_ca1
                        loop
                    case op2 = 3
                        save screen to tela_ca1
                        do sinal with "Saida no Estoque Clien"
                        do relsai1
                        close all
                        restore screen from tela_ca1
                        loop
                    case op2 = 4
                        save screen to tela_ca1
                        do sinal with "Saida no Estoque Vende"
                        do relsai2
                        close all
                        restore screen from tela_ca1
                        loop
                    other
                        exit
                endcase
            enddo
            restore screen from telaop32
            loop
        case op = 3
            save screen to telaop32
            do sinal with "Caixa                "
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            do while .t.
                set color to w/b
                @ 23,05 say "<ESC> Para sair."
                Janela2("W",.F.,06,39,10,59,.F.)
                SetColor("N/W,W+/R")
                @ 07,40 prompt "Retirada do Caixa"
                @ 08,40 prompt "Entrada no Caixa "
                @ 09,40 prompt "Relatorio        "+chr(016)
                menu to op3
                do case
                    case op3 = 1
                        save screen to tela_ca1
                        do sinal with "Retirada do Caixa     "
                        do cadcai1
                        close all
                        restore screen from tela_ca1
                        loop
                    case op3 = 2
                        save screen to tela_ca2
                        do sinal with "Entrada no Caixa      "
                        do cadcai2
                        close all
                        restore screen from tela_ca2
                        loop
                    case op3 = 3
                        save screen to telaop3
                        do sinal with "Relatorio Diversos    "
                        set color to w/b
                        @ 23,05 say "<ESC> Para sair."
                        do while .t.
                            set color to w/b
                            @ 23,05 say "<ESC> Para sair."
                            Janela2("W",.F.,10,55,14,78,.F.)
                            SetColor("N/W,W+/R")
                            @ 11,56 prompt "Relatorio Caixa       "
                            @ 12,56 prompt "Extrato Venda a Vista "
                            @ 13,56 prompt "Historico do Caixa    "
                            save screen to telarel1
                            menu to op4
                            do case
                                case op4 = 1
                                    save screen to tela_ca2
                                    do sinal with "Relatorio Caixa       "
                                    do reltic
                                    close all
                                    restore screen from tela_ca2
                                    loop
                                case op4 = 2
                                    save screen to tela_ca2
                                    do sinal with "Extrato Venda a Vista "
                                    do relextv
                                    close all
                                    restore screen from tela_ca2
                                    loop
                                case op4 = 3
                                    save screen to tela_ca2
                                    do sinal with "Historio Caixa      a "
                                    do relhist
                                    close all
                                    restore screen from tela_ca2
                                    loop

                                other
                                    exit
                            endcase
                        enddo
                        restore screen from telaop3
                        loop
                    other
                        exit
                endcase
            enddo
            restore screen from telaop32
            loop
            close all
        case op = 4
            save screen to tela_ca
            do sinal with "Contas a Pagar       "
            save screen to telaop32
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            do while .t.
                set color to w/b
                @ 23,05 say "<ESC> Para sair."
                Janela2("W",.F.,07,39,10,65,.F.)
                SetColor("N/W,W+/R")
                @ 08,40 prompt "Conta a Pagar           "
                @ 09,40 prompt "Relatorio               "+chr(016)
                menu to op3
                do case
                    case op3 = 1
                        save screen to tela_ca2
                        do sinal with "Contas a Pagar        "
                        do cadpaga
                        close all
                        restore screen from tela_ca2
                        loop
                    case op3 = 2
                        do sinal with "Relatorios          "
                        save screen to telaop41333
                        do while .t.
                            set color to w/b
                            @ 23,05 say "<ESC> Para sair."
                            Janela2("W",.F.,08,57,14,76,.F.)
                            SetColor("N/W,W+/R")
                            @ 09,58 prompt "Contas a Pagar    "
                            @ 10,58 prompt "Contas Gerais     "
                            @ 11,58 prompt "Tipo de Conta     "
                            @ 12,58 prompt "Data Pagamento    "
                            @ 13,58 prompt "Contas Pagas      "
                            save screen to telaopc41333
                            menu to op41333
                            do case
                                case op41333=1
                                    do sinal with "Rel. Contas a Pagar"
                                    do relvence
                                    limpa()
                                    restore screen from telaopc41333
                                    loop
                                case op41333=2
                                    do sinal with "Rel. Contas Gerais  "
                                    do relconta
                                    limpa()
                                    restore screen from telaopc41333
                                    loop
                                case op41333=3
                                    do sinal with "Rel. Tipo de Contas "
                                    do relcont1
                                    limpa()
                                    restore screen from telaopc41333
                                    loop
                                case op41333=4
                                    do sinal with "Rel. Data Pagamento "
                                    do relpag1
                                    limpa()
                                    restore screen from telaopc41333
                                    loop
                                case op41333=5
                                    do sinal with "Rel. Contas   Pagas"
                                    do relpaga
                                    limpa()
                                    restore screen from telaopc41333
                                    loop
                                other
                                    close all
                                    exit
                            endcase
                            restore screen from telaop41333
                        enddo
                        limpa()
                        restore screen from telaop41333
                        loop
                    other
                        exit
                endcase
            enddo
            restore screen from telaop32
            loop
            close all
        case op = 5
            save screen to tela_ca
            do sinal with "Contas a Receber     "
            save screen to telaop32
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            do while .t.
                set color to w/b
                @ 23,05 say "<ESC> Para sair."
                Janela2("W",.F.,07,39,10,65,.F.)
                SetColor("N/W,W+/R")
                @ 08,40 prompt "Baixa Contas a Receber  "
                @ 09,40 prompt "Relatorio               "+chr(016)
                menu to op3
                do case
                    case op3 = 1
                        save screen to tela_ca2
                        do sinal with "Baixa Contas a receber"
                        do cadrec
                        close all
                        restore screen from tela_ca2
                        loop
                    case op3 = 2
                        save screen to telaop3
                        do sinal with "Relatorio Diversos    "
                        set color to w/b
                        @ 23,05 say "<ESC> Para sair."
                        do while .t.
                            set color to w/b
                            @ 23,05 say "<ESC> Para sair."
                            Janela2("W",.F.,10,55,13,72,.F.)
                            SetColor("N/W,W+/R")
                            @ 11,56 prompt "Contas a Receber"
                            @ 12,56 prompt "Contas Recebidas"
                            menu to op4
                            do case
                                case op4 = 1
                                    save screen to tela_ca1
                                    do sinal with "Relat. Contas a Recebe"
                                    do relcont3
                                    close all
                                    restore screen from tela_ca1
                                    loop
                                case op4 = 2
                                    save screen to tela_ca2
                                    do sinal with "Relat. Contas Recebida"
                                    do relcont4
                                    close all
                                    restore screen from tela_ca2
                                    loop
                                other
                                    exit
                            endcase
                        enddo
                        restore screen from telaop3
                        loop
                    other
                        exit
                endcase
            enddo
            restore screen from telaop32
            loop
            close all
        case op = 6
            save screen to telabb
            select 1
            if !net_use("clientes",.f.,"ind01.cli","ind02.cli")
                return
            endif
            primeiro = "1"
   red_1:= ""
   red_2:= "M"
   tir_r1:= ""
   tir_r2:= "P"
   red_sal:= "0"
   tir_sal:= "2"
   col_gra:= "t"
   tir_gra:= "t" + Chr(0)
   col_exp:= "W"
   tir_exp:= "W" + Chr(0)
   col_enf:= "E"
   tir_enf:= "F"
   inic_imp:= "@"
   col_roman:= "x"
   col_ace:= "6"
   tir_ace:= "7"
   col_sub:= "S"
   tir_sub:= "T"

            do while .t.
                wdatav      = date()
                wdatad      = date()
                wcodigo_cli = "00000"
                wndocu      = space(10)
                waceite     = space(02)
                wvalorDoc   = 0
                wdesconto   = 0
                wobse1      = space(65)
                wobse2      = space(65)
                wobse3      = space(65)
                set color to n/w
                quadro(10,30,21,73)
                set color to w/r
                @ 10,44 say "  Imprime Boleto "
                set color to /w
                @ 11,31 say "Cliente...:"
                @ 12,31 say "Vencimento:"
                @ 13,31 say "Data Docum:"
                @ 14,31 say "N§ Docum. :"
                @ 15,31 say "Aceite....:"
                @ 16,31 say "Valor Docu:"
                @ 17,31 say "Desconto..:"
                @ 18,31 say "Observacao:"
                limpa()
                set color to w/b
                @ 23,05 say "<ESC> Para sair. <00000> para Consulta"
                set color to w,n/bg,,,n/w
                @ 11,42 get wcodigo_cli pict "99999"
                read
                if readkey() = 12  .or. lastkey() = 27
                    close all
                    return
                endif
                if empty(wcodigo_cli) .or. wcodigo_cli = "00000"
                    select 1
                    save screen to tela_01
                    do pes_cli with "clientes"
                    restore screen from tela_01
                    iguala()
                endif
                set color to /w
                @ 11,41 say wcodigo_cli pict "99999"
                select 1
                set index to ind01.cli
                seek wcodigo_cli
                if !found()
                    do mensagem with "Codigo do Cliente nao Cadastrao...",8
                    loop
                endif
                iguala()
                @ 11,49 say wnome pict "@s24"
                wender = subsT(wendereco,1,30) + subs("Bairro " + wbairro,1,15) + subst(" CEP " + wcep,1,15)
                set color to w,n/bg,,,n/w
                @ 12,42 get wdatav    pict "@d"
                @ 13,42 get wdatad    pict "@d"
                @ 14,42 get wndocu    pict "@!"
                @ 15,42 get waceite   pict "@!"
                @ 16,42 get wvalordoc pict "@e@z 999,999.99"
                @ 17,42 get wdesconto pict "@e@z 999,999.99"
                @ 18,42 get wobse1    pict "@s20"
                @ 19,42 get wobse2    pict "@s20"
                @ 20,42 get wobse3    pict "@s20"
                read
                SetColor("N/W,W+/R")
                @ 21,33      prompt "Retorno "
                @ 21,col()+5 prompt "Imprimir"
                menu to bb
                do case
                    case bb = 1
                        restore screen from telabb
                        close all
                        return
                    case bb = 2
                        wdatad = subst(dtoc(wdatad),1,2)+"/"+subst(dtoc(wdatad),4,2)+"/"+subst(dtoc(wdatad),9,2)
                        setprc(0,0)
                        set color to w/b
                        set devi to print
                        set cons off

                        pg:= 0
                        @ PRow(),  0 say inic_imp
                        @ PRow(),  0 say ""
                        @ PRow() + 1, 51 say wdatav pict "@d"
                        @ PRow(),  0 say red_1
                        @ PRow() + 4,  3 say wdatad pict "@d"
                        @ PRow(), PCol() + 9 say wndocu pict "@d"
                        @ PRow(), PCol() + 10 say "REC."
                        @ PRow(), PCol() + 15 say waceite pict "@!"
                        @ PRow(),  0 say tir_r1
                        @ PRow() + 1, 54 say valordoc picture "@E 9,999.99"
                        @ PRow() + 1,  0 say red_1
                        @ PRow() + 1,  3 say wobse1 pict "@!"
                        @ PRow() + 1,  3 say wobse2 pict "@!"
                        @ PRow() + 1,  3 say wobse3 pict "@!"
                        @ 15, 15 say wnome pict "@!"
                        @ 16, 15 say wender pict "@!"
                        @ PRow() + 8,  0 say ""
                        @ PRow(),  0 say tir_r1
                        setprc(0, 0)
                        set cons on
                        set devi to screen
                        set print off
                        set print to
                    other
                        restore screen from telabb
                        loop
                endcase
            enddo
            restore screen from telabb
            loop

        other
            close all
            exit
    endcase
enddo
close all
return




















