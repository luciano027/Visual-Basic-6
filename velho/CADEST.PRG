*****************************************************************************
* Programa...: CADEST.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 05/06/95          Data da Ultima Atualizacao: 05/06/95       *
* Objetivo...: Cadastrar produtos no estoque                                *
* Sistema....: DO SUL (c_saldo de Estoque)                                 *
*****************************************************************************
*****> Abre arquivo de fornecedores
select 1
if !net_use("estoque",.f.,"ind01.est","ind02.est")
    return
endif
select 2
if !net_use("fornece",.f.,"ind01.for","ind02.for")
    return
endif

do while .t.
    select 2
    inicia()
    select 1
    go bott
    iguala()
    ycodigo_est = strzero(val(wcodigo_est)+1,6,0)
    inicia()
    wcodigo_est = ycodigo_est
    Janela2("W",.F.,06,01,19,75,.F.)
    set color to w/r
    @ 06,27 say "CADASTRO PRODUTO ESTOQUE"
    set color to /w
    @ 07,03 say "Codigo........:"
    @ 08,03 say "Descricao.....:"
    @ 09,03 say "Unidade.......:"
    @ 10,03 say "Fabrica.......:                            ÚÄÄÄÄÄÄÄ   RESUMO    ÄÄÄÄÄÄ¿"
    @ 11,03 say "Saldo.Minimo..:           Saldo:           ³                          ³"
    @ 12,03 say "Controlar Saldo:                           ³                          ³"
    @ 13,03 say "Preco de Venda.......R$                    ³                          ³"
    @ 14,03 say "Fornecedor............:                    ³                          ³"
    @ 15,03 say "Data Ultima Compra....:                    ³                          ³"
    @ 16,03 say "Quantidade Comprada...:                    ³                          ³"
    @ 17,03 say "Preco de Compra......R$                    ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ"
    @ 18,03 say "Aliquota Fiscal.......:"
    set color to w/b
    @ 23,05 say "<ESC> Para sair.<000000> para consulta"
    set color to w,n/bg,,,n/w
    save screen to telafor
    @ 07,18 get wcodigo_est pict "999999"
    read
    if readkey() = 12 .or. lastkey() = 27
        libera()
        close all
        return
    endif
    if wcodigo_est = "000000"
        save screen to tela_0
        do pes_est
        restore screen from tela_0
    endif
    ycodigo_est = wcodigo_est
    set color to /w
    @ 07,18 say ycodigo_est pict "999999"
    select 1
    set index to ind01.est
    seek ycodigo_est
    if found()
        iguala()
        select 2
        set index to ind01.for
        seek wcodigo_for
        iguala()
        set color to n/w
        @ 08,18 say wdescricao  pict "@!"
        @ 09,18 say wunidade    pict "@!"
        @ 10,18 say wfabrica    pict "@!"
        @ 11,35 say wsaldo      pict "@e@z 99,999.99"
        @ 11,18 say wminino     pict "@e@z 99,999.99"
        @ 12,19 say wc_saldo    pict "!"
        @ 13,27 say wpreco_ven  pict "@e 999,999.999"
        @ 18,26 say walicota    pict "!!"
        set color to gr/w
        @ 14,26 say wcodigo_for     pict "@!"
        @ 14,32 say wfornecedor     pict "@s10"
        @ 15,26 say wdata_com       pict "@d"
        @ 16,26 say wquant_com      pict "@e@z 99,999.99"
        @ 17,26 say wpreco_com      pict "@e@z 99,999.99"
        select 1
        set color to r/w
        wespecifica = memoedit(especifica,11,47,16,72,.f.,.f.)
        limpa()
        save screen to tela_an
        set color to w/b
        @ 23,05 say "<ESC> Para sair."
        Janela2("W",.F.,03,61,08,70,.F.)
        SetColor("N/W,W+/R")
        @ 04,62   prompt "RETORNO "
        @ 05,62   prompt "ALTERAR "
        @ 06,62   prompt "EXCLUIR "
        @ 07,62   prompt "RESUMO  "
        menu to xop
        restore screen from tela_an
        do case
            case xop = 1
                limpa()
                loop
            case xop = 2
                limpa()
                set color to w/b
                @ 23,05 say "<ESC> Retorna."
                set color to w,n/bg,,,n/w
                @ 08,18 get wdescricao  pict "@s25" valid !empty(wdescricao)
                @ 09,18 get wunidade        pict "@!"
                @ 10,18 get wfabrica        pict "@!"
                @ 11,18 get wminino         pict "@e@z 99,999.99"
                @ 12,19 get wc_saldo        pict "!"
                @ 13,27 get wpreco_ven      pict "@e@z 999,999.999"
                read
                if lastkey() = 27 .or. readkey() = 12
                    limpa()
                    close all
                    return
                endif
                do while .t.
                    set color to w/b
                    @ 23,05 say "<ESC> Para sair. <000> para Consulta"
                    set color to w,n/bg,,,n/w
                    @ 14,26 get wcodigo_for pict "999"
                    read
                    if readkey() = 12  .or. lastkey() = 27 .or. empty(wcodigo_for)
                        close all
                        return
                    endif
                    if wcodigo_for = "000"
                        select 2
                        save screen to tela_01
                        do pes_for
                        restore screen from tela_01
                        iguala()
                    endif
                    if empty(wcodigo_for)
                        close all
                        return
                    endif
                    set color to /w
                    @ 14,26 say wcodigo_for pict "999"
                    select 2
                    set index to ind01.for
                    seek wcodigo_for
                    if !found()
                        do mensagem with "Fornecedor nao cadastrado...",8
                        loop
                    endif
                    iguala()
                    set color to gr/w
                    @ 14,26 say wcodigo_for     pict "@!"
                    @ 14,32 say wfornecedor     pict "@s10"
                    exit
                enddo
                set color to w,n/bg,,,n/w
                @ 18,26 get walicota    pict "!!"
                read
                limpa()
                if Confirma ("Confirma Dados?")=1    && gravacao dos dados
                    select 1
                    set index to ind01.est,ind02.est
                    seek wcodigo_est
                    if found()
                        if rec_lock()
                            revar()
                            unlock
                        endif
                    endif
                endif
                loop
            case xop = 3
                limpa()
                if Confirma ("Confirma Exclusao?")=1    && gravacao dos dados
                    limpa()
                    select 1
                    set index to ind01.est,ind02.est
                    seek wcodigo_est
                    if found()
                        if rec_lock()
                            delete
                        endif
                    endif
                endif
                loop
            case xop = 4
                select 1
                set index to ind01.est,ind02.est
                seek wcodigo_est
                if found()
                    if rec_lock()
                        replace especifica with memoed(especifica,"",11,47,16,72)
                        unlock
                    endif
                endif
            other
                loop
        endcase
    else
        limpa()
        @ 23,05 say "<ESC> Retorna."
        set color to w,n/bg,,,n/w
        @ 08,18 get wdescricao  pict "@!" valid !empty(wdescricao)
        @ 09,18 get wunidade    pict "@!"
        @ 10,18 get wfabrica    pict "@!"
        @ 11,18 get wminino     pict "99999"
        @ 12,19 get wc_saldo    pict "!"
        @ 13,27 get wpreco_ven  pict "@e 999,999.999"
        @ 18,26 get walicota    pict "!!"
        read
        if lastkey() = 27 .or. readkey() = 12
            @ 23,03 say space(75)
            close all
            return
        endif
        limpa()
        if Confirma ("Confirma Dados?")=1    && gravacao dos dados
            limpa()
            select 1
            set index to ind01.est,ind02.est
            if add_rec(10)
                revar()
                replace especifica with memoed(especifica,"",11,47,16,72)
            endif
            loop
        endif
    endif
enddo
close all
return

    ****************************************************************************

procedure pes_est
select 1
set index to ind02.est
go top
sai=.f.
set color to n/bg
lp=06
tk=00
ul=09
save screen to tela93
set color to gr+/bg
quadro(04,01,22,76)
quadro(05,55,16,75)
@ 05,03 say "  Codigo                    Descricao              "
@ 06,62 say "RESUMO"
for q = 06 to 16
    set color to w+/bg
    @ q,03 say "                                                    "
next
set color to n/bg
@ 18,04 say "ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ"
@ 19,04 say "Preco de Venda  R$                  Fabrica:"
@ 20,04 say "Saldo Loja:"
set color to r/bg
@ 21,04 say "                   <F9> Consulta por Descricao"
do while .not. sai
    set color to gr+/b
    l=6   && primeira linha da tela l = 9
    do while l<=16 .and. .not. eof()
        set color to /bg
        do dadosc2 with l
        skip
        l=l+1
    enddo
    nl=l
    do while nl<=16
        @ nl,03 say space(52)
        nl=nl+1
    enddo
    l=l-1
    ul=l
    if tk<0 .or. tk=22
        goto rec
        lp=6
    else
        skip lp-(l+1)
    endif
    pl=.f.
    if l<16
        if lp>6
            lp = 6
        endif
        pl=.t.
    endif
    set color to w+/r
    do dadosc2 with lp
    set color to n/bg
    do while .t.
        tk=inkey(0)
        do case
            case tk=24
                skip
                if eof()
                    do mensagem with "Fim do Arquivo !",4
                    set color to n/bg
                    loop
                else
                    skip -1
                endif
                if lp<16 &&-----> 20
                    do dadosc2 with lp
                    lp=lp+1
                    skip
                    set color to w+/r
                    do dadosc2 with lp
                    set color to n/bg
                    loop
                else
                    set color to n/bg
                    do dadosc2 with lp
                    skip
                    if .not. eof()
                        scroll(6,3,16,54,1)
                    endif
                    set color to w+/r
                    do dadosc2 with lp
                    set color to n/bg
                    loop
                endif
            case tk=5
                if bof()
                    do mensagem with "Inicio do Arquivo !",4
                    set color to n/bg
                    loop
                endif
                if lp>6
                    if eof()
                        skip -1
                    endif
                    do dadosc2 with lp
                    lp=lp-1
                    skip -1
                    set color to w+/r
                    do dadosc2 with lp
                    set color to n/bg
                    loop
                else
                    set color to n/bg
                    if eof()
                        skip -1
                    endif
                    do dadosc2 with lp
                    skip -1
                    if .not. bof()
                        scroll(6,3,16,54,-1)
                    endif
                    set color to w+/r
                    do dadosc2 with lp
                    set color to n/bg
                    loop
                endif
            case tk=13
                save screen
                set color to n/bg
                sai = .t.
                t = 1
                iguala()
                return(wcodigo_est)
            case tk=18
                if bof()
                    do mensagem with "Inicio do Arquivo !",4
                    loop
                endif
                if eof()
                    if lastrec()>=14
                        skip -7
                    else
                        go top
                    endif
                else
                    if lastrec()>14
                        skip -(lp-9+7)
                    else
                        go top
                    endif
                endif
                exit
            case tk=3
                if ul<16
                    do mensagem with "Fim do Arquivo !",4
                    set color to n/bg
                    loop
                endif
                skip 9-lp
                exit
            case tk=1
                go top
                l  = 16
                lp = 6
                exit
            case tk=6
                go bottom
                l  = 16
                lp = 6
                exit
            case tk=-8
                do limpa
                set order to 2
                vnome=space(40)
                @ 23,10 say "Nome Produto : " get vnome picture "@!"
                read
                do limpa
                if empty(vnome)
                    loop
                endif
                rec=recno()
                vnome=trim(vnome)
                seek vnome
                if .not. found()
                    do mensagem with "Desculpe, Produto n„o Encontrado !",6,1
                    set color to n/bg
                    go rec
                    loop
                endif
                rec=recno()
                l  = 16
                lp = 6
                exit
            otherwise
                loop
        endcase
    enddo
    do limpa
enddo
close all
return

    ************************
procedure dadosc2
parameters lin
iguala()
@ lin, 06 say wcodigo_est pict "@!"
@ lin, 14 say wdescricao  pict "@s40"
set color to n/bg
@ 19,22 say wpreco_ven    pict "@e@z 999,999.999"
@ 19,48 say wfabrica      pict "@!"
@ 20,15 say wsaldo        pict "@e@z 999,999.99"
wespecifica = memoedit(especifica,07,56,15,74,.f.,.f.)
return
    **************************
