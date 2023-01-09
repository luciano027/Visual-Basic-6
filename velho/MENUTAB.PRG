*****************************************************************************
* Programa...: MENUTAB.PRG                                                  *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao:                *
* Objetivo...: MENU OPCOES MANUTENCAO E CADASTRO DE INFORMACOES BANCO DADOS *
* Sistema....: SISTEMA DE CAIXA E CONTROLE DE VENDAS IF                     *
*****************************************************************************
save screen to telaopcao1
do sinal with "  Tabelas           "
limpa()
do while .t.
    set color to w/b
    @ 23,05 say "<ESC> Para sair."
    SetColor("W/W")
    quadro(03,04,10,26)
    Janela2("W",.F.,03,04,11,26,.F.)
    SetColor("N/W,W+/R")
    @ 04,05 prompt "Produtos no Estoque "+chr(016)
    @ 05,05 prompt "Fornecedor          "+chr(016)
    @ 06,05 prompt "Vendedor            "+chr(016)
    @ 07,05 prompt "Clientes            "+chr(016)
    @ 08,05 prompt "Cadastro Tipo Conta "+chr(016)
    @ 09,05 prompt "Cadastro de Juros   "
    @ 10,05 prompt "Agenda              "
    save screen to telaop1
    menu to op1
    do case
        case op1 = 1
            save screen to telaop31
            do sinal with "Cadastro no Estoque   "
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            do while .t.
                set color to w/b
                @ 23,05 say "<ESC> Para sair."
                Janela2("W",.F.,05,26,08,55,.F.)
                SetColor("N/W,W+/R")
                @ 06,27 prompt "Cadastra/Altera Estoque  "
                @ 07,27 prompt "Relatorios Diversos      "+chr(016)
                save screen to telarel1
                menu to op1
                do case
                    case op1 = 1
                        save screen to tela_ca1
                        do sinal with "Cadastra   Estoque    "
                        do cadest
                        close all
                        restore screen from tela_ca1
                        loop
                    case op1 = 2
                        save screen to telaop3
                        do sinal with "Relatorio Diversos    "
                        set color to w/b
                        @ 23,05 say "<ESC> Para sair."
                        do while .t.
                            set color to w/b
                            @ 23,05 say "<ESC> Para sair."
                            SetColor("W/W")
                            Janela2("W",.F.,07,55,15,77,.F.)
                            SetColor("N/W,W+/R")
                            @ 08,56 prompt "Inventario           "
                            @ 09,56 prompt "Produtos Cadastrados "
                            @ 10,56 prompt "Tabela de Precos     "
                            @ 11,56 prompt "Tabela de Compras    "
                            @ 12,56 prompt "Saldo Minimo         "
                            @ 13,56 prompt "Tabela de Precos Desc"
                            @ 14,56 prompt "Extrato Produtos Clie"
                            save screen to telarel1
                            menu to op2
                            do case
                                case op2 = 1
                                    save screen to tela_ca2
                                    do sinal with "Inventario            "
                                    do relinv
                                    close all
                                    restore screen from tela_ca2
                                    loop
                                case op2 = 2
                                    save screen to tela_ca2
                                    do sinal with "Produtos Cadastrados  "
                                    do relprod
                                    close all
                                    restore screen from tela_ca2
                                    loop
                                case op2 = 3
                                    save screen to tela_ca2
                                    do sinal with "Tabela de Precos      "
                                    do reltab
                                    close all
                                    restore screen from tela_ca2
                                    loop
                                case op2 = 4
                                    save screen to tela_ca2
                                    do sinal with "Tabela de Precos Compr"
                                    do reltab1
                                    close all
                                    restore screen from tela_ca2
                                    loop
                                case op2 = 5
                                    save screen to tela_ca2
                                    do sinal with "Saldo Minimo          "
                                    do relmin
                                    close all
                                    restore screen from tela_ca2
                                    loop
                                case op2 = 6
                                    save screen to tela_ca2
                                    do sinal with "Tabela de Precos      "
                                    do reltab3
                                    close all
                                    restore screen from tela_ca2
                                    loop
                                case op2 = 7
                                    save screen to tela_ca2
                                    do sinal with "Extrato Produto Client"
                                    do relextp
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
            restore screen from telaop31
            loop
        case op1 = 2
            save screen to telaop3
            do sinal with "Cadastro Fornecedor   "
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            do while .t.
                set color to w/b
                @ 23,05 say "<ESC> Para sair."
                SetColor("W/W")
                Janela2("W",.F.,05,26,10,48,.F.)
                SetColor("N/W,W+/R")
                @ 06,27 prompt "Cadastro Fornecedor  "
                @ 07,27 prompt "Relatorio Produtos   "
                @ 08,27 prompt "Relatorio Fornecedor "
                @ 09,27 prompt "Relatorio Pedido     "
                save screen to telarel1
                menu to op2
                do case
                    case op2 = 1
                        save screen to tela_ca1
                        do sinal with "Cadastro Fornecedor   "
                        do cadfor
                        close all
                        restore screen from tela_ca1
                        loop
                    case op2 = 2
                        save screen to tela_ca2
                        do sinal with "Relatorio Produtos    "
                        do relprod2
                        close all
                        restore screen from tela_ca2
                        loop
                    case op2 = 3
                        save screen to tela_ca2
                        do sinal with "Relatorio Fornecedor  "
                        do relfor
                        close all
                        restore screen from tela_ca2
                        loop
                    case op2 = 4
                        save screen to tela_ca2
                        do sinal with "Relatorio Pedido      "
                        do relped
                        close all
                        restore screen from tela_ca2
                        loop
                    other
                        exit
                endcase
            enddo
            restore screen from telaop3
            loop
        case op1 = 3
            save screen to telaop3
            do sinal with "Cadastro Vendedor     "
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            do while .t.
                set color to w/b
                @ 23,05 say "<ESC> Para sair."
                SetColor("W/W")
                Janela2("W",.F.,07,26,10,48,.F.)
                SetColor("N/W,W+/R")
                @ 08,27 prompt "Cadastro Vendedor    "
                @ 09,27 prompt "Extrato de Vendas    "
                save screen to telarel1
                menu to op2
                do case
                    case op2 = 1
                        save screen to tela_ca1
                        do sinal with "Cadastro Vendedor     "
                        do cadven
                        close all
                        restore screen from tela_ca1
                        loop
                    case op2 = 2
                        save screen to tela_ca2
                        do sinal with "Extrato de Vendas     "
                        do relextve
                        close all
                        restore screen from tela_ca2
                        loop
                    other
                        exit
                endcase
            enddo
            restore screen from telaop3
            loop
        case op1 = 4
            save screen to telaop3
            do sinal with "Cadastro Clientes     "
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            do while .t.
                set color to w/b
                @ 23,05 say "<ESC> Para sair."
                SetColor("W/W")
                Janela2("W",.F.,08,26,14,48,.F.)
                SetColor("N/W,W+/R")
                @ 09,27 prompt "Cadastro Clientes    "
                @ 10,27 prompt "Extrato do Cliente   "
                @ 11,27 prompt "Cliente em Debito    "
                @ 12,27 prompt "Clientes em Debito 30"
                @ 13,27 prompt "Compras do Cliente   "
                save screen to telarel1
                menu to op2
                do case
                    case op2 = 1
                        save screen to tela_ca1
                        do sinal with "Cadastro Clientes     "
                        do cadcli
                        close all
                        restore screen from tela_ca1
                        loop
                    case op2 = 2
                        save screen to tela_ca2
                        do sinal with "Extrato de Vendas     "
                        do relextvc
                        close all
                        restore screen from tela_ca2
                        loop
                    case op2 = 3
                        save screen to tela_ca2
                        do sinal with "Cliente em Debito     "
                        do relclid
                        close all
                        restore screen from tela_ca2
                        loop
                    case op2 = 4
                        save screen to tela_ca2
                        do sinal with "Cliente em Debito     "
                        do relclid3
                        close all
                        restore screen from tela_ca2
                        loop
                        case op2 = 5
                        save screen to tela_ca1
                        do sinal with "Saida no Estoque Clien"
                        do relsai5
                        close all
                        restore screen from tela_ca1
                        loop


                    other
                        exit
                endcase
            enddo
            restore screen from telaop3
            loop
        case op1 = 5
            save screen to telaop3
            do sinal with "Cadastro Tipo de Conta"
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            do while .t.
                set color to w/b
                @ 23,05 say "<ESC> Para sair."
                SetColor("W/W")
                Janela2("W",.F.,07,26,10,48,.F.)
                SetColor("N/W,W+/R")
                @ 08,27 prompt "Cadastra Tipo Conta  "
                @ 09,27 prompt "Relatorio Tipo Conta "
                save screen to telarel1
                menu to op2
                do case
                    case op2 = 1
                        save screen to tela_ca1
                        do sinal with "Cadastro Tipo Conta   "
                        do cadtipo
                        close all
                        restore screen from tela_ca1
                        loop
                    case op2 = 2
                        save screen to tela_ca2
                        do sinal with "Relatorio Tipo Conta  "
                        do reltipo
                        close all
                        restore screen from tela_ca2
                        loop
                    other
                        exit
                endcase
            enddo
            restore screen from telaop3
            loop
        case op1 = 6
            save screen to telaop3
            do sinal with "Cadastro de Juros     "
            close all
            do cadjur
            close all
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            restore screen from telaop3
            loop
        case op1 = 7
            save screen to telaop3
            do sinal with "Agenda                "
            close all
            do cadage
            close all
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            restore screen from telaop3
            loop

        other
            exit
    endcase
    close all
    exit
enddo
close all
return
                                                                                                                                                    *****************************************************************************
                                                                                                                                                    *                          F   I   M                                        *
                                                                                                                                                    *****************************************************************************
procedure calcula
select 1
if !net_use("estoque",.f.,"ind01.est","ind02.est")
    return
endif

select 2
if !net_use("clientes",.f.,"ind01.cli","ind02.cli","ind03.cli")
    return
endif

select 3
if !net_use("prazo",.f.,"ind01.pra")
    return
endif

select 4
if !net_use("saida",.f.,"ind01.sai")
    return
endif

select 5
if !net_use("ticket",.f.,"ind01.tic")
    return
endif

select 3
index on codigo_cli to &zmicro
set index to &zmicro
do while !eof()
    iguala()
    ycodigo_cli = wcodigo_cli
    wtotal = 0
    ydata  = data
    do while wcodigo_cli = ycodigo_cli
        select 1
        set index to ind01.est
        seek wcodigo_est
        iguala()
        select 3
        wtotal = (wtotal + (wvalor * wquantidade))  &&----> imprime o total
        skip
        iguala()
    enddo
    select 2
    set index to ind01.cli
    seek ycodigo_cli
    if found()
        wdata_deb = data_deb
        if rec_lock()
            replace debito with wtotal
            unlock
            if wdata_deb = ctod("")
                if rec_lock()
                    replace data_deb with ydata
                    unlock
                endif
            endif
        endif
    endif
    select 3
enddo

close all
return


