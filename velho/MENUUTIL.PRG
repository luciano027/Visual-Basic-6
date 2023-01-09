*****************************************************************************
* Programa...: MENUUTIL.PRG                                                 *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao:                *
* Objetivo...: MENU UTILITARIOS                                             *
* Sistema....: SISTEMA DE CAIXA E CONTROLE DE VENDAS IF                     *
*****************************************************************************
limpa()
set color to w/b
@ 23,05 say "<ESC> Para sair."
do sinal with "Utilitarios         "
save screen to telautil1
do while .t.
    set color to w/b
    @ 23,05 say "<ESC> Para sair."
    Janela2("W",.F.,03,43,13,69,.F.)
    SetColor("N/W,W+/R")
    @ 04,44 prompt "Manutencao Arquivos     "
    @ 05,44 prompt "Restaura Arquivos       "
    @ 06,44 prompt "Backup do Arquivos      "
    @ 07,44 prompt "Cadastro Empresa        "
    @ 08,44 prompt "Fecha Caixa             "
    @ 09,44 prompt "Mapa Resumo do Caixa    "
    @ 10,44 prompt "Abertura Caixa   (X)    "
    @ 11,44 prompt "Fechamento Caixa (Z)    "

    menu to op6
    do case
        case op6 =1
            do sinal with "Atualiza Arquivos   "
            limpa()
            save screen to telaopc6
            do atualiza
            testa_arq()
            limpa()
            restore screen from telaopc6
            loop
        case op6 =2
            do sinal with "Restaura Arquivos   "
            limpa()
            save screen to telaoc6
            if testa_drive("A")
                do rest
            endif
            limpa()
            restore screen from telaoc6
            loop
        case op6 =3
            do sinal with "Backu Arquivos      "
            limpa()
            save screen to telac6
            if testa_drive("A")
                do backp
            endif
            limpa()
            restore screen from telac6
            loop
        case op6 =4
            do sinal with "Cadastro Empresa    "
            limpa()
            save screen to telaopc6
            do empresa
            limpa()
            restore screen from telaopc6
            loop
        case op6 =5
            save screen to telaop65
            if Confirma ("Confirma Fechamento do Caixa ?")=1    && gravacao dos dados
                close all
                excluir()
            atualiza()
            testa_arq()
            endif
            limpa()
            restore screen from telaop65
            loop
        case op6 =6
            save screen to tela_ca2
            do sinal with "Resumo impressora     "
            do naimp
            close all
            restore screen from tela_ca2
            loop
        case op6 = 7
            save screen to telaop65
            save screen to porta
            Janela2("W",.F.,11,16,19,71,.F.)
            set color to w/r
            @ 11,28 say "ABERTURA DO CUPOM FISCAL DIARIO"
            SetColor("N/W,W+/R")
            @ 12,18 say "A Leitura X e o cupom que apresenta o conteudo  atual"
            @ 13,18 say "de todos os registradores, acumuladores e  contadores"
            @ 14,18 say "internos, sem a alteracao de seus valores, e pode ser"
            @ 15,18 say "solicitado a qualquer momento."
            @ 16,18 say "A emissao do cupom de leitura  X  e  obrigratoria  no"
            @ 17,18 say "inicio de cada dia, no inicio e no final de cada nova"
            @ 18,18 say "bobina de papel."
            inkey(1)
            do sinal with "Leitura X           "
            if Confirma ("Confirma Leitura X ?")=1    && gravacao dos dados
                close all
                leiturax()
            endif
            restore screen from telaop65
            loop
        case op6 = 8
            save screen to telaop65
            do sinal with "Reducao Z           "
            Janela2("W",.F.,11,16,15,68,.F.)
            set color to w/r
            @ 11,25 say "FECHAMENTO DO CUPOM FISCAL DIARIO"
            SetColor("N/W,W+/R")
            @ 12,18 say "A reducao Z e uma operacao obrigatoria e deve ser "
            @ 13,18 say "executado uma vez ao dia, ao final do horario de  "
            @ 14,18 say "funcionamento do estabelecimento.                 "
            inkey(1)
            if Confirma ("Confirma Reducao Z ?")=1    && gravacao dos dados
                close all
                reducaoz()
            endif
            limpa()
            restore screen from telaop65
            loop

        other
            exit
    endcase
    restore screen from telaopcao6
    exit
enddo
close all
return


***************************************************************************
procedure empresa
select 1
if !net_use("empresa",.f.,"","")
    return
endif
do while .t.
    do telaemp
    inicia()
    locate for !empty(empresa)
    if found()
        iguala()
        set color to /w
        @ 07,24 say  wempresa   pict "@!"
        @ 09,24 say  wendereco  pict "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
        @ 10,24 say  wtelefone  pict "(9999)999-9999"
        @ 10,51 say  wfax       pict "(9999)999-9999"
        @ 11,24 say  wcgc       pict "@!"
        @ 12,24 say  winscricao pict "@!"
        @ 13,24 say  wdiretorio pict "@!"
        @ 14,24 say  wmicro     pict "@!"
        @ 15,24 say wdireto     pict "@!"
        limpa()
        save screen to tela_an
        set color to w/b
        @ 23,05 say "<ESC> Para sair."
        SetColor("W/W")
        Janela2("W",.F.,03,61,07,70,.F.)
        SetColor("N/W,W+/R")
        @ 04,62   prompt "RETORNO "
        @ 05,62   prompt "ALTERAR "
        @ 06,62   prompt "EXCLUIR "
        menu to xop
        restore screen from tela_an
        do case
            case xop = 1
                exit
            case xop = 2
                if !senha1()
                    close all
                    return
                endif
                set color to w,n/bg,,,n/w
                @ 07,24 get  wempresa   pict "@!"
                @ 09,24 get  wendereco  pict "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
                @ 10,24 get  wtelefone  pict "(9999)999-9999"
                @ 10,51 get  wfax       pict "(9999)999-9999"
                @ 11,24 get  wcgc       pict "@!"
                @ 12,24 get  winscricao pict "@!"
                @ 13,24 get wdiretorio  pict "@!"
                @ 14,24 get wmicro      pict "@!"
                @ 15,24 get wdireto     pict "@!"
                read
                p = "S"
                set color to w/b
                @ 23,01 say space(79)
                @ 23,05 say "Confirma Alteracao <s/n> ? " get p pict "@!" valid p $"SN"
                read
                @ 23,03 say space(77)
                if upper(p) = "S"
                    if rec_lock()
                        revar()
                        unlock
                    endif
                endif
                loop
            case xop = 3
                if !senha1()
                    close all
                    return
                endif
                p = "S"
                set color to w/b
                @ 23,01 say space(79)
                @ 23,05 say "Confirma Exclusao <s/n> ? " get p pict "@!" valid p $"SN"
                read
                @ 23,03 say space(77)
                if upper(p) = "S"
                    if rec_lock()
                        delete
                    endif
                endif
            other
                loop
        endcase
    else
        set color to w,n/bg,,,n/w
        @ 07,24 get  wempresa   pict "@!"
        @ 09,24 get  wendereco  pict "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
        @ 10,24 get  wtelefone  pict "(9999)999-9999"
        @ 10,51 get  wfax       pict "(9999)999-9999"
        @ 11,24 get  wcgc       pict "@!"
        @ 12,24 get  winscricao pict "@!"
        @ 13,24 get  wdiretorio pict "@!"
        @ 14,24 get  wmicro     pict "@!"
        @ 15,24 get  wdireto    pict "@!"
        read
        p = "S"
        set color to w/b
        limpa()
        @ 23,05 say "Confirma Inclusao <s/n> ? " get p pict "@!" valid p $"SN"
        read
        limpa()
        if upper(p) = "S"
            if add_rec(10)
                revar()
            endif
            exit
        else
            exit
        endif
    endif
enddo
return

***************************************************************************
procedure telaemp
set color to /w
dt = quadro(06,09,16,74)
@ 07,10 say "Empresa .....:"
@ 09,10 say "Endereco.....:"
@ 10,10 say "Telefone.....:                   FAX...:"
@ 11,10 say "C.G.C........:"
@ 12,10 say "Insc. Estad..:"
@ 13,10 say "Diretorio....:"
@ 14,10 say "Micro........:"
@ 15,10 say "Diretorio....:"
return


*****************************************************************************
*       Inicia a porta serial para leuitura X
*****************************************************************************

procedure leiturax
nPort := 0
tempo := 0
do while nPort = 0

    if nPort = 0
        if COM_OPEN(1, 4000)
            COM_INIT(1, 9600, "N", 8 , 1)
            COM_RTS(1, .T.)
            COM_DTR(1, .T.)
            COM_SEND (1,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (1,29)     /* ESC */
            COM_SEND (1,5)
            COM_FLUSH(1)
            millisec(100)
            nCaract := COM_COUNT(1)
            if nCaract > 1
                rec_char = COM_READ (1, 1)
                rec_char = COM_READ (1, 1)
                if (rec_char = chr(13))
                    nPort := 1
                    veloc = '9.600'
                else
                    COM_CLOSE (1)
                endif
            else
                COM_CLOSE (1)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(2, 4000)
            COM_INIT(2, 9600, "N", 8 , 1)
            COM_RTS(2, .T.)
            COM_DTR(2, .T.)
            COM_SEND (2,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (2,29)     /* ESC */
            COM_SEND (2,5)
            COM_FLUSH(2)
            millisec(100)
            nCaract := COM_COUNT(2)
            if nCaract > 1
                rec_char = COM_READ (2, 1)
                rec_char = COM_READ (2, 1)
                if (rec_char = chr(13))
                    nPort := 2
                    veloc = '9.600'
                else
                    COM_CLOSE (2)
                endif
            else
                COM_CLOSE (2)
            endif
        endif
    endif


    if nPort = 0
        if COM_OPEN(1, 4000)
            COM_INIT(1, 600, "N", 8 , 1)
            COM_RTS(1, .T.)
            COM_DTR(1, .T.)
            COM_SEND (1,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (1,29)     /* ESC */
            COM_SEND (1,5)
            COM_FLUSH(1)
            millisec(400)
            nCaract := COM_COUNT(1)
            if nCaract > 1
                rec_char = COM_READ (1, 1)
                rec_char = COM_READ (1, 1)
                if (rec_char = chr(13))
                    nPort := 1
                    veloc = '  600'
                else
                    COM_CLOSE (1)
                endif
            else
                COM_CLOSE (1)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(1, 4000)
            COM_INIT(1,  300, "N", 8 , 1)
            COM_RTS(1, .T.)
            COM_DTR(1, .T.)
            COM_SEND (1,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (1,29)     /* ESC */
            COM_SEND (1,5)
            COM_FLUSH(1)
            millisec(400)
            nCaract := COM_COUNT(1)
            if nCaract > 1
                rec_char = COM_READ (1, 1)
                rec_char = COM_READ (1, 1)
                if (rec_char = chr(13))
                    nPort := 1
                    veloc = '  300'
                else
                    COM_CLOSE (1)
                endif
            else
                COM_CLOSE (1)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(2, 4000)
            COM_INIT(2, 600, "N", 8 , 1)
            COM_RTS(2, .T.)
            COM_DTR(2, .T.)
            COM_SEND (2,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (2,29)     /* ESC */
            COM_SEND (2,5)
            COM_FLUSH(2)
            millisec(400)
            nCaract := COM_COUNT(2)
            if nCaract > 1
                rec_char = COM_READ (2, 1)
                rec_char = COM_READ (2, 1)
                if (rec_char = chr(13))
                    nPort := 2
                    veloc = '  600'
                else
                    COM_CLOSE (2)
                endif
            else
                COM_CLOSE (2)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(2, 4000)
            COM_INIT(2, 300, "N", 8 , 1)
            COM_RTS(2, .T.)
            COM_DTR(2, .T.)
            COM_SEND (2,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (2,29)     /* ESC */
            COM_SEND (2,5)
            COM_FLUSH(2)
            millisec(400)
            nCaract := COM_COUNT(2)
            if nCaract > 1
                rec_char = COM_READ (2, 1)
                rec_char = COM_READ (2, 1)
                if (rec_char = chr(13))
                    nPort := 2
                    veloc = '  300'
                else
                    COM_CLOSE (2)
                endif
            else
                COM_CLOSE (2)
            endif
        endif
    endif
    if nPort = 0
        if tempo > 3
            save screen to porta
            Janela2("W",.F.,11,16,14,62,.F.)
            SetColor("N/W,W+/R")
            @ 12,18 say 'NŽO FOI POSSIVEL ENCONTRAR A SUA IMPRESSORA'
            @ 13,23 say 'DIGITE ALGO PARA REPETIR A BUSCA'
            inkey(1)
            COM_CLOSE (nPort)
            keyboard chr(027)
            return
            restore screen from porta
        endif
        tempo = tempo + 1
    endif
enddo
millisec(200)
COM_FLUSH(nPort)
enviacodigo(6)
COM_CLOSE (nPort)
return

*****************************************************************************
*       Inicia a porta serial reducao Z
*****************************************************************************

procedure reducaoz
nPort := 0
tempo := 0
do while nPort = 0

    if nPort = 0
        if COM_OPEN(1, 4000)
            COM_INIT(1, 9600, "N", 8 , 1)
            COM_RTS(1, .T.)
            COM_DTR(1, .T.)
            COM_SEND (1,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (1,29)     /* ESC */
            COM_SEND (1,5)
            COM_FLUSH(1)
            millisec(100)
            nCaract := COM_COUNT(1)
            if nCaract > 1
                rec_char = COM_READ (1, 1)
                rec_char = COM_READ (1, 1)
                if (rec_char = chr(13))
                    nPort := 1
                    veloc = '9.600'
                else
                    COM_CLOSE (1)
                endif
            else
                COM_CLOSE (1)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(2, 4000)
            COM_INIT(2, 9600, "N", 8 , 1)
            COM_RTS(2, .T.)
            COM_DTR(2, .T.)
            COM_SEND (2,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (2,29)     /* ESC */
            COM_SEND (2,5)
            COM_FLUSH(2)
            millisec(100)
            nCaract := COM_COUNT(2)
            if nCaract > 1
                rec_char = COM_READ (2, 1)
                rec_char = COM_READ (2, 1)
                if (rec_char = chr(13))
                    nPort := 2
                    veloc = '9.600'
                else
                    COM_CLOSE (2)
                endif
            else
                COM_CLOSE (2)
            endif
        endif
    endif
                                                                                                                                                                                                               *  Teste das velocidades adicionais 300bps e 600
    if nPort = 0
        if COM_OPEN(1, 4000)
            COM_INIT(1, 600, "N", 8 , 1)
            COM_RTS(1, .T.)
            COM_DTR(1, .T.)
            COM_SEND (1,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (1,29)     /* ESC */
            COM_SEND (1,5)
            COM_FLUSH(1)
            millisec(400)
            nCaract := COM_COUNT(1)
            if nCaract > 1
                rec_char = COM_READ (1, 1)
                rec_char = COM_READ (1, 1)
                if (rec_char = chr(13))
                    nPort := 1
                    veloc = '  600'
                else
                    COM_CLOSE (1)
                endif
            else
                COM_CLOSE (1)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(1, 4000)
            COM_INIT(1,  300, "N", 8 , 1)
            COM_RTS(1, .T.)
            COM_DTR(1, .T.)
            COM_SEND (1,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (1,29)     /* ESC */
            COM_SEND (1,5)
            COM_FLUSH(1)
            millisec(400)
            nCaract := COM_COUNT(1)
            if nCaract > 1
                rec_char = COM_READ (1, 1)
                rec_char = COM_READ (1, 1)
                if (rec_char = chr(13))
                    nPort := 1
                    veloc = '  300'
                else
                    COM_CLOSE (1)
                endif
            else
                COM_CLOSE (1)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(2, 4000)
            COM_INIT(2, 600, "N", 8 , 1)
            COM_RTS(2, .T.)
            COM_DTR(2, .T.)
            COM_SEND (2,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (2,29)     /* ESC */
            COM_SEND (2,5)
            COM_FLUSH(2)
            millisec(400)
            nCaract := COM_COUNT(2)
            if nCaract > 1
                rec_char = COM_READ (2, 1)
                rec_char = COM_READ (2, 1)
                if (rec_char = chr(13))
                    nPort := 2
                    veloc = '  600'
                else
                    COM_CLOSE (2)
                endif
            else
                COM_CLOSE (2)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(2, 4000)
            COM_INIT(2, 300, "N", 8 , 1)
            COM_RTS(2, .T.)
            COM_DTR(2, .T.)
            COM_SEND (2,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (2,29)     /* ESC */
            COM_SEND (2,5)
            COM_FLUSH(2)
            millisec(400)
            nCaract := COM_COUNT(2)
            if nCaract > 1
                rec_char = COM_READ (2, 1)
                rec_char = COM_READ (2, 1)
                if (rec_char = chr(13))
                    nPort := 2
                    veloc = '  300'
                else
                    COM_CLOSE (2)
                endif
            else
                COM_CLOSE (2)
            endif
        endif
    endif
    if nPort = 0
        if tempo > 3
            save screen to porta
            Janela2("W",.F.,11,16,14,62,.F.)
            SetColor("N/W,W+/R")
            @ 12,18 say 'NŽO FOI POSSIVEL ENCONTRAR A SUA IMPRESSORA'
            @ 13,23 say 'DIGITE ALGO PARA REPETIR A BUSCA'
            inkey(1)
            COM_CLOSE (nPort)
            keyboard chr(027)
            restore screen from porta
            return
        endif
        tempo = tempo +1
    endif
enddo
millisec(200)
COM_FLUSH(nPort)
verdata=date()
verhora=time()
par_fiscal= strtran(strtran((dtoc(verdata)+verhora),'/',''),':','')
enviacodigo(7)
COM_CLOSE (nPort)

*****************************************************************************
*       Mapa Resumo Caixa na Impressora
*****************************************************************************
procedure naimp
save screen to telaantes
set color to w+/r
dt = quadro(11,17,13,60)
set color to *+w/r
@ 12,18 say "Aguarde... Procurando Impressora Fiscal"
nPort := 0
tempo := 0
do while nPort = 0

    if nPort = 0
        if COM_OPEN(1, 4000)
            COM_INIT(1, 9600, "N", 8 , 1)
            COM_RTS(1, .T.)
            COM_DTR(1, .T.)
            COM_SEND (1,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (1,29)     /* ESC */
            COM_SEND (1,5)
            COM_FLUSH(1)
            millisec(100)
            nCaract := COM_COUNT(1)
            if nCaract > 1
                rec_char = COM_READ (1, 1)
                rec_char = COM_READ (1, 1)
                if (rec_char = chr(13))
                    nPort := 1
                    veloc = '9.600'
                else
                    COM_CLOSE (1)
                endif
            else
                COM_CLOSE (1)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(2, 4000)
            COM_INIT(2, 9600, "N", 8 , 1)
            COM_RTS(2, .T.)
            COM_DTR(2, .T.)
            COM_SEND (2,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (2,29)     /* ESC */
            COM_SEND (2,5)
            COM_FLUSH(2)
            millisec(100)
            nCaract := COM_COUNT(2)
            if nCaract > 1
                rec_char = COM_READ (2, 1)
                rec_char = COM_READ (2, 1)
                if (rec_char = chr(13))
                    nPort := 2
                    veloc = '9.600'
                else
                    COM_CLOSE (2)
                endif
            else
                COM_CLOSE (2)
            endif
        endif
    endif
                                                                                                                                                                                                               *  Teste das velocidades adicionais 300bps e 600
    if nPort = 0
        if COM_OPEN(1, 4000)
            COM_INIT(1, 600, "N", 8 , 1)
            COM_RTS(1, .T.)
            COM_DTR(1, .T.)
            COM_SEND (1,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (1,29)     /* ESC */
            COM_SEND (1,5)
            COM_FLUSH(1)
            millisec(400)
            nCaract := COM_COUNT(1)
            if nCaract > 1
                rec_char = COM_READ (1, 1)
                rec_char = COM_READ (1, 1)
                if (rec_char = chr(13))
                    nPort := 1
                    veloc = '  600'
                else
                    COM_CLOSE (1)
                endif
            else
                COM_CLOSE (1)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(1, 4000)
            COM_INIT(1,  300, "N", 8 , 1)
            COM_RTS(1, .T.)
            COM_DTR(1, .T.)
            COM_SEND (1,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (1,29)     /* ESC */
            COM_SEND (1,5)
            COM_FLUSH(1)
            millisec(400)
            nCaract := COM_COUNT(1)
            if nCaract > 1
                rec_char = COM_READ (1, 1)
                rec_char = COM_READ (1, 1)
                if (rec_char = chr(13))
                    nPort := 1
                    veloc = '  300'
                else
                    COM_CLOSE (1)
                endif
            else
                COM_CLOSE (1)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(2, 4000)
            COM_INIT(2, 600, "N", 8 , 1)
            COM_RTS(2, .T.)
            COM_DTR(2, .T.)
            COM_SEND (2,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (2,29)     /* ESC */
            COM_SEND (2,5)
            COM_FLUSH(2)
            millisec(400)
            nCaract := COM_COUNT(2)
            if nCaract > 1
                rec_char = COM_READ (2, 1)
                rec_char = COM_READ (2, 1)
                if (rec_char = chr(13))
                    nPort := 2
                    veloc = '  600'
                else
                    COM_CLOSE (2)
                endif
            else
                COM_CLOSE (2)
            endif
        endif
    endif
    if nPort = 0
        if COM_OPEN(2, 4000)
            COM_INIT(2, 300, "N", 8 , 1)
            COM_RTS(2, .T.)
            COM_DTR(2, .T.)
            COM_SEND (2,13)     /* Este <Enter> n„o atrapalha   */
            COM_SEND (2,29)     /* ESC */
            COM_SEND (2,5)
            COM_FLUSH(2)
            millisec(400)
            nCaract := COM_COUNT(2)
            if nCaract > 1
                rec_char = COM_READ (2, 1)
                rec_char = COM_READ (2, 1)
                if (rec_char = chr(13))
                    nPort := 2
                    veloc = '  300'
                else
                    COM_CLOSE (2)
                endif
            else
                COM_CLOSE (2)
            endif
        endif
    endif
    if nPort = 0
        if tempo > 3
            save screen to porta
            Janela2("W",.F.,11,16,14,62,.F.)
            SetColor("N/W,W+/R")
            @ 12,18 say 'NŽO FOI POSSIVEL ENCONTRAR A SUA IMPRESSORA'
            @ 13,23 say 'DIGITE ALGO PARA REPETIR A BUSCA'
            inkey(1)
            COM_CLOSE (nPort)
            keyboard chr(027)
            restore screen from porta
            return
        endif
        tempo = tempo +1
    endif
enddo
millisec(200)
COM_FLUSH(nPort)
data_inicial := data_final := space( 08 )
set color to
set color to w+/gr
@ 11,30 clear to 14,54
@ 11,30 to 14,54
@ 12,31 say "Data Inicial:" get data_inicial pict "99/99/99"
@ 13,31 say "Data Final..:" get data_final   pict "99/99/99"
read
data_inicial := substr( data_inicial, 1, 2 )+substr( data_inicial, 4, 2 )+substr( data_inicial, 7, 2 )
data_final   := substr( data_final, 1, 2 )+substr( data_final, 4, 2 )+substr( data_final, 7, 2 )
par_fiscal   = "X"+data_inicial+data_final
enviacodigo(8)
COM_CLOSE (nPort)
close all
return

*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
