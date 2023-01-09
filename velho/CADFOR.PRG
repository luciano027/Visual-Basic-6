*****************************************************************************
* Programa...: CADCLI.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 02/12/98          Data da Ultima Atualizacao:                *
* Objetivo...: CADASTRO DE CLIENTES                                         *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************

select 1
if !net_use("fornece",.f.,"ind01.for","ind02.for")
    return
endif

do while .t.
    select 1
    inicia()
    set index to ind01.for
    go bott
    wcodigo_for = strzero(val(codigo_for)+1,3,0)
    set color to n/w
    dt = quadro(06,02,17,73)
    set color to w/r
    @ 06,25 say " CADASTRO FORNECEDOR "
    set color to /w
    @ 07,03 say "Codigo...:"
    @ 08,03 say "Empresa..:"
    @ 09,02 say "쳐컴컴컴컴컴컴컴컴컴컴컴컴컴   Cadastro Geral  컴컴컴컴컴컴컴컴컴컴컴컴�"
    @ 10,03 say "Endereco.:"
    @ 11,03 say "Bairro...:                                   CEP.:"
    @ 12,03 say "Cidade...:                                    UF.:"
    @ 13,03 say "Telefone.:                                    FAX:"
    @ 14,03 say "Contato..:                                Celular:"
    @ 15,03 say "R.G......:                          C.P.F./C.I.C.:"
    @ 16,03 say "C.G.C....:                          Ins.Estadual.:"
    set color to w/b
    @ 23,05 say "<ESC> Para sair. <000> para Consulta"
    set color to w,n/bg,,,n/w
    @ 07,13 get wcodigo_for pict "999"
    read
    if readkey() = 12  .or. lastkey() = 27
        close all
        return
    endif
    if empty(wcodigo_for) .or. wcodigo_for = "000"
        select 1
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
    @ 07,13 say wcodigo_for pict "999"
    select 1
    set index to ind01.for
    seek wcodigo_for
    if found()
        iguala()
        set color to r/w
        @ 08,13 say wfornecedor    pict "@!"
        @ 10,13 say wendereco   pict "@!"
        @ 11,13 say wbairro     pict "@!"
        @ 11,53 say wcep        pict "@!"
        @ 12,13 say wcidade     pict "@!"
        @ 12,53 say wuf         pict "!!"
        @ 13,13 say wtelefone   pict "@!"
        @ 13,53 say wfax        pict "@!"
        @ 14,13 say wcontato    pict "@!"
        @ 14,53 say wcelular    pict "@!"
        @ 15,13 say wrg         pict "@!"
        @ 15,53 say wcpf        pict "@!"
        @ 16,13 say wcgc        pict "@!"
        @ 16,53 say winscr      pict "@!"
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
                limpa()
                exit
            case xop = 2
                limpa()
                set color to w/b
                @ 23,05 say "<ESC> Para sair. "
                set color to w,n/bg,,,n/w
                @ 08,13 get wfornecedor    pict "@!"
                @ 10,13 get wendereco   pict "@!"
                @ 11,13 get wbairro     pict "@!"
                @ 11,53 get wcep        pict "@!"
                @ 12,13 get wcidade     pict "@!"
                @ 12,53 get wuf         pict "!!"
                @ 13,13 get wtelefone   pict "@!"
                @ 13,53 get wfax        pict "@!"
                @ 14,13 get wcontato    pict "@!"
                @ 14,53 get wcelular    pict "@!"
                @ 15,13 get wrg         pict "@!"
                @ 15,53 get wcpf        pict "@!"
                @ 16,13 get wcgc        pict "@!"
                @ 16,53 get winscr      pict "@!"
                read
                if readkey() = 12  .or. lastkey() = 27
                    close all
                    return
                endif
                limpa()
                if Confirma ("Confirma Dados?")=1    && gravacao dos dados
                    select 1
                    set index to ind01.for,ind02.for
                    seek wcodigo_for
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
                    set index to ind01.for,ind02.for
                    seek wcodigo_for
                    if found()
                        if rec_lock()
                            delete
                        endif
                    endif
                endif
                loop
            other
                loop
        endcase
    else
        set color to w/b
        @ 23,05 say "<ESC> Para sair. "
        set color to w,n/bg,,,n/w
        @ 08,13 get wfornecedor    pict "@!"
        @ 10,13 get wendereco   pict "@!"
        @ 11,13 get wbairro     pict "@!"
        @ 11,53 get wcep        pict "@!"
        @ 12,13 get wcidade     pict "@!"
        @ 12,53 get wuf         pict "!!"
        @ 13,13 get wtelefone   pict "@!"
        @ 13,53 get wfax        pict "@!"
        @ 14,13 get wcontato    pict "@!"
        @ 14,53 get wcelular    pict "@!"
        @ 15,13 get wrg         pict "@!"
        @ 15,53 get wcpf        pict "@!"
        @ 16,13 get wcgc        pict "@!"
        @ 16,53 get winscr      pict "@!"
        read
        if readkey() = 12  .or. lastkey() = 27
            close all
            return
        endif

        limpa()
        if Confirma ("Confirma Dados?")=1    && gravacao dos dados
            select 1
            set index to ind01.for,ind02.for
            if add_rec(10)
                revar()
            endif
        endif
    endif
    loop
enddo
close all
return

*****************************procedure pesquisa de fornecedor************************
procedure pes_for
set index to ind01.for
go bott
ycodigo_for = strzero(val(codigo_for)+1,3,0)
set index to ind02.for
go top
sai=.f.
set color to n/bg
lp=06
tk=00
ul=09
save screen to tela93
set color to gr+/bg
dt = quadro(04,01,19,56)
@ 05,03 say   "  Codigo                 Fornecedor                  "
for q = 06 to 16
    set color to w+/bg
    @ q,03 say "                                                    "
next
set color to n/bg
@ 17,04 say "컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴"
set color to r/bg
@ 18,04 say "         <F9> Consulta fornecedor               "
set color to n/bg

do while .not. sai
    l=6   && primeira linha da tela l = 9
    do while l<=16 .and. .not. eof()
        set color to /bg
        do dadofor with l
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
    do dadofor with lp
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
                    do dadofor with lp
                    lp=lp+1
                    skip
                    set color to w+/r
                    do dadofor with lp
                    set color to n/bg
                    loop
                else
                    set color to n/bg
                    do dadofor with lp
                    skip
                    if .not. eof()
                        scroll(6,3,16,54,1)
                    endif
                    set color to w+/r
                    do dadofor with lp
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
                    do dadofor with lp
                    lp=lp-1
                    skip -1
                    set color to w+/r
                    do dadofor with lp
                    set color to n/bg
                    loop
                else
                    set color to n/bg
                    if eof()
                        skip -1
                    endif
                    do dadofor with lp
                    skip -1
                    if .not. bof()
                        scroll(6,3,16,54,-1)
                    endif
                    set color to w+/r
                    do dadofor with lp
                    set color to n/bg
                    loop
                endif
            case tk=13
                save screen
                set color to n/bg
                sai = .t.
                t = 1
                return(CODIGO_for)
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
                set index to ind02.for
                vnome=space(40)
                @ 23,10 say "Nome do Fornecedor : " get vnome picture "@!"
                read
                do limpa
                if empty(vnome)
                    loop
                endif
                rec=recno()
                vnome=trim(vnome)
                seek vnome
                if .not. found()
                    do mensagem with "Desculpe, Fornecedor n꼘 Encontrado !",6,1
                    set color to n/bg
                    go rec
                    loop
                endif
                rec=recno()
                l  = 0
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

**********************************************************
procedure dadofor
* Apresentacao de uma linha de dados para consulta na tela
**********************************************************
parameters lin
@ lin, 06 say codigo_for pict "@!"
@ lin, 13 say fornecedor   pict "@s42"
return
**********************************************************
