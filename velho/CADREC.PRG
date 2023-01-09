*****************************************************************************
* Programa...: APAGAR.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao:                *
* Objetivo...: CADASTRO DE CONTAS A PAGAR                                   *
* Sistema....: SISPAN - Sistema de Panifica‡Æo                              *
*****************************************************************************

select 1
if !net_use("clientes",.f.,"ind01.cli","ind02.cli")
    return
endif

select 2
if !net_use("receber",.f.,"ind01.rec")
    return
endif

do while .t.
    select 1
    inicia()
    select 2
    inicia()
    Janela2("W",.F.,04,08,22,78,.F.)
    set color to r/bg
    @ 04,30 say "CADASTRO CONTAS A RECEBER"
    SetColor("N/W,W+/R")
    @ 05,09 say "Cliente.....:"
    @ 06,09 say "Documento...:"
    @ 07,09 say "Historico...:"
    @ 09,09 say "ÚÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿"
    @ 10,09 say "³Parcela³ Vencimento ³   Valor   ³ Data Pagto ³  Valor  ³ Documento³"
    @ 11,09 say "ÃÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄ´"
    @ 12,09 say "³       ³            ³           ³            ³         ³          ³"
    @ 13,09 say "³       ³            ³           ³            ³         ³          ³"
    @ 14,09 say "³       ³            ³           ³            ³         ³          ³"
    @ 15,09 say "³       ³            ³           ³            ³         ³          ³"
    @ 16,09 say "³       ³            ³           ³            ³         ³          ³"
    @ 17,09 say "³       ³            ³           ³            ³         ³          ³"
    @ 18,09 say "³       ³            ³           ³            ³         ³          ³"
    @ 19,09 say "ÃÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´"
    @ 20,09 say "³  Total a Pagar :                 Total Pago.:                    ³"
    @ 21,09 say "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ"
    limpa()
    save screen to telat
    do while .t.
        set color to w/b
        @ 23,05 say "<ESC> Para sair. <00000> para Consulta"
        set color to w,n/bg,,,n/w
        @ 05,23 get wcodigo_cli pict "99999"
        read
        if readkey() = 12  .or. lastkey() = 27
            exit
        endif
        if empty(wcodigo_cli) .or. wcodigo_cli = "00000"
            select 1
            save screen to tela_01
            do pes_cli with "clientes"
            restore screen from tela_01
            iguala()
        endif
        set color to /w
        @ 05,23 say wcodigo_cli pict "99999"
        select 1
        set index to ind01.cli
        seek wcodigo_cli
        if !found()
            do mensagem with "empresa nao cadastrado...",8
            loop
        endif
        iguala()
        set color to r/w
        @ 05,23 say wcodigo_cli    pict "@!"
        @ 05,30 say wnome      pict "@s39"
        exit
    enddo
    save screen to tela_21
    set color to w,n/bg,,,n/w
    @ 06,23 get wdocumento pict "@!"  when tab_docr(07,23,20,67,2,wcodigo_cli)
    read
    set key  5 to
    set key 24 to
    if readkey() = 12  .or. lastkey() = 27 .or. empty(wdocumento)
        close all
        return
    endif
    restore screen from tela_21
    set color to r/w
    @ 06,23 say wdocumento pict "@!"
    zcodigo_cli = wcodigo_cli
    zdocumento  = wdocumento
    do while .t.
        select 2
        set index to ind01.rec
        seek zcodigo_cli+zdocumento
        if found()
            iguala()
            znome = wnome
            set color to r/w
            @ 05,23 say wcodigo_cli    pict "@!"
            @ 05,30 say wnome          pict "@s39"
            @ 06,23 say wdocumento     pict "@!"
            @ 07,23 say whistorico     pict "@!"
            ydocumento  = wdocumento
            ycodigo_cli = wcodigo_cli
            yhistorico  = whistorico
            totaa = 0
            totap = 0
            li = 12
            select 2
            index on parcela to &zmicro for documento = wdocumento .and. codigo_cli = wcodigo_cli
            set index to &zmicro
            do while !eof()
                @ li,12 say wparcela    pict "@!"
                @ li,19 say wvenci      pict "@d"
                @ li,32 say wvalor_v    pict "@e@z 99,999.99"
                @ li,45 say wpagto      pict "@d"
                @ li,56 say wvalor_p    pict "@e@z 99,999.99"
                @ li,66 say wboleto     pict "@!"
                li = li + 1
                if li = 19
                    li = 12
                    lip()
                endif
                totaa = totaa + wvalor_v
                totap = totap + wvalor_p
                select 2
                skip
                iguala()
            enddo
            wdocumento  = ydocumento
            wcodigo_cli = ycodigo_cli
            whistorico  = yhistorico
            set color to r/w
            @ 20,28 say totaa    pict "@e@z 99,999.99"
            @ 20,56 say totap    pict "@e@z 99,999.99"
            limpa()
            save screen to tela_a
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            SetColor("W/W")
            Janela2("W",.F.,03,66,08,75,.F.)
            SetColor("N/W,W+/R")
            @ 04,67   prompt "RETORNO "
            @ 05,67   prompt "EXCLUIR "
            @ 06,67   prompt "BAIXA   "
            @ 07,67   prompt "PARCELA "
            menu to xop
            restore screen from tela_a
            do case
                case xop = 1
                    limpa()
                    exit
                case xop = 2
                    limpa()
                    if Confirma ("Confirma Exclusao ? Todo o Documento sera excluido? ")=1    && gravacao dos dados
                        select 2
                        index on parcela to indcon.rec for documento = ydocumento .and. codigo_cli = ycodigo_cli
                        set index to indcon.rec
                        do while !eof()
                            iguala()
                            if wdocumento = ydocumento
                                if rec_lock()
                                    delete
                                endif
                            endif
                            skip
                        enddo
                        exit
                    endif
                    loop
                case xop = 3
                    save screen to tela09
                    select 2
                    index on parcela to indcon.rec for documento = ydocumento .and. codigo_cli = ycodigo_cli
                    set index to indcon.rec
                    go top
                    do pescon_4 with 2
                    reg = recno()
                    locate for reg = recno()
                    iguala()
                    set color to w/b
                    dt = quadro(09,20,17,75)
                    @ 10,22 say "Parcela...:"
                    @ 11,22 say "Vencimento:"
                    @ 12,22 say "Valor.....:"
                    @ 13,22 say "Pagamento.:"
                    @ 14,22 say "Valor Pago:"
                    @ 15,22 say "Documento.:"
                    @ 10,33 say wparcela    pict "@!"
                    @ 11,33 say wvenci      pict "@d"
                    @ 12,33 say wvalor_v    pict "@e 99,999.99"
                    @ 13,33 say wpagto      pict "@d"
                    @ 14,33 say wvalor_p    pict "@e 99,999.99"
                    @ 15,33 say wboleto     pict "@!"
                    yvalor_v = wvalor_v
                    limpa()
                    set color to  w/b
                    @ 16,22      prompt "RETORNO "
                    @ 16,col()+2 prompt "ALTERAR "
                    @ 16,col()+2 prompt "EXCLUIR "
                    @ 16,col()+2 prompt "BAIXA "
                    menu to qop
                    do case
                        case qop = 1
                            set color to w/b
                            @ 23,01 say space(79)
                            restore screen from tela09
                            loop
                        case qop = 2
                            yvalor_v = wvalor_v
                            set color to w,n/bg,,,n/w
                            @ 10,33 get wparcela    pict "@!"
                            @ 11,33 get wvenci      pict "@d"
                            @ 12,33 get wvalor_v    pict "@e 99,999.99"
                            read
                            limpa()
                            if Confirma (" Dados Corretos? ")=1    && gravacao dos dados     select 1
                                select 2
                                set index to ind01.rec
                                go reg
                                if rec_lock()
                                    replace parcela    with wparcela
                                    replace venci      with wvenci
                                    replace valor_v    with wvalor_v
                                    unlock
                                endif
                            endif
                            restore screen from tela09
                            loop
                        case qop = 3
                            limpa()
                            if Confirma ("Confirma Exclusao? ")=1    && gravacao dos dados     select 1
                                select 2
                                set index to ind01.rec
                                go reg
                                if rec_lock()
                                    delete
                                endif
                            endif
                            restore screen from tela09
                            loop
                        case qop = 4
                            set color to w,n/bg,,,n/w
                            @ 13,33 get wpagto      pict "@d"
                            @ 14,33 get wvalor_p    pict "@e 99,999.99"
                            @ 15,33 get wboleto     pict "@!"
                            read
                            limpa()
                            if Confirma (" Dados Corretos? ")=1    && gravacao dos dados     select 1
                                select 2
                                set index to ind01.rec
                                go reg
                                if rec_lock()
                                    replace pagto      with wpagto
                                    replace valor_p    with wvalor_p
                                    replace boleto     with wboleto
                                    unlock
                                endif
                            endif
                            restore screen from tela09
                            loop
                        other
                            restore screen from tela09
                            loop
                    endcase
                    restore screen from tela09
                    loop
                case xop = 4
                    set color to w/b
                    @ 23,05 say "<ESC> Para sair."
                    set color to w,n/bg,,,n/w
                    @ li,12 get wparcela    pict "@!"
                    @ li,19 get wvenci      pict "@d"
                    @ li,32 get wvalor_v    pict "@e@z 99,999.99"
                    read
                    if readkey() = 12  .or. lastkey() = 27
                        exit
                    endif
                    limpa()
                    if Confirma (" Dados Corretos? ")=1    && gravacao dos dados     select 1
                        select 2
                        set index to ind01.rec
                        if add_rec(10)
                            revar()
                        endif

                        li = li + 1
                        if li = 19
                            li = 12
                            lip()
                        endif
                    endif
                    loop
                other
                    loop
            endcase
            loop
        else
            save screen to telat
            set color to w,n/bg,,,n/w
            @ 07,23 get whistorico pict "@!"
            read
            if readkey() = 12  .or. lastkey() = 27
                close all
                return
            endif
            li = 12
            do while .t.
                set color to w/b
                @ 23,05 say "<ESC> Para sair."
                set color to w,n/bg,,,n/w
                @ li,12 get wparcela    pict "@!"
                @ li,19 get wvenci      pict "@d"
                @ li,32 get wvalor_v    pict "@e@z 99,999.99"
                read
                if readkey() = 12  .or. lastkey() = 27
                    lip()
                    exit
                endif
                limpa()
                if Confirma (" Dados Corretos? ")=1    && gravacao dos dados     select 1
                    select 2
                    set index to ind01.rec
                    if add_rec(10)
                        revar()
                    endif
                    li = li + 1
                    if li = 19
                        li = 12
                        lip()
                    endif
                endif
            enddo
        endif
    enddo
enddo
close all
return

**********************pesquisa da banco********************************
function tab_docr(x1,y1,x2,y2,sel,docu)
sele &sel
index on dtos(venci) to indcon.tip for codigo_cli = docu .and. pagto = ctod("  /  /  ")
set index to indcon.tip
go top
set color to w/b
@ 23,01 say space(79)
@ 23,10 say "<"+chr(24)+"> Passa Tela. <"+chr(25)+"> Retorna Tela. <ESC> Retorna."
public telav:=savescreen(x1,y1,x2,y2)
@ x1,y1 clear to x2,y2
dt = quadro(x1,y1,x2,y2+6)
public xx1,xx2,yy1,yy2
xx1 = x1
xx2 = x2
yy1 = y1
yy2 = y2
linha = x1+1
@ linha,y1+1     say "Documento"
@ linha,col()+5  say "Parcela  "
@ linha,col()+5  say "Valor"
@ linha,col()+5  say "Vencimento"
linha++
@ linha,y1+1 to linha,(y2+6)-1
linha++
for x = 1 to 10
    @ linha,yy1+1    say documento pict "@!"
    @ linha,col()+5  say parcela   pict "@!"
    @ linha,col()+05 say valor_v   pict "@e 99,999.99"
    @ linha,col()+05 say venci     pict "@d"
    linha++
    skip
next
set key 5  to sobe_docr
set key 24 to desce_docr
return .t.

function sobe_docr
linha = 0
@ xx1,yy1 clear to xx2,yy2
dt = quadro(xx1,yy1,xx2,yy2+6)
linha = xx1+1
@ linha,yy1+1     say "Documento"
@ linha,col()+5  say "Parcela  "
@ linha,col()+5  say "Valor"
@ linha,col()+5  say "Vencimento"
linha++
@ linha,yy1+1 to linha,(yy2+6)-1
linha++
goto top && skip -10
for x = 1 to 10
    @ linha,yy1+1    say documento pict "@!"
    @ linha,col()+5  say parcela   pict "@!"
    @ linha,col()+05 say valor_v   pict "@e 99,999.99"
    @ linha,col()+05 say venci     pict "@d"
    linha++
    skip
next
return .t.

function desce_docr
@ xx1,yy1 clear to xx2,yy2
dt = quadro(xx1,yy1,xx2,yy2+6)
linha = xx1+1
@ linha,yy1+1     say "Documento"
@ linha,col()+5  say "Parcela  "
@ linha,col()+5  say "Valor"
@ linha,col()+5  say "Vencimento"
linha++
@ linha,yy1+1 to linha,(yy2+6)-1
linha++
if .not. eof()
    for x = 1 to 10
        @ linha,yy1+1    say documento pict "@!"
        @ linha,col()+5  say parcela   pict "@!"
        @ linha,col()+05 say valor_v   pict "@e 99,999.99"
        @ linha,col()+05 say venci     pict "@d"
        linha++
        skip
    next
endif
return .t.

procedure pescon_4
para sele
select &sele
go top
sai=.f.
lp=12
tk=0
ul=18
set color to w/b
@ 23,03 say "Sobe :"+chr(24)+" Desce :"+chr(25)+" <ENTER> Confirma  <ESC> Sai "
set color to /w
do while .not. sai
    l=12
    do while l<=18 .and. .not. eof()
        do dados64c with l
        skip
        l=l+1
    enddo
    nl=l
    do while nl<=18
        nl=nl+1
    enddo
    l=l-1
    ul=l
    if tk<0 .or. tk=18
        goto rec
        lp=12
    else
        skip lp-(l+1)
    endif
    pl=.f.
    if l<18
        if lp>12
            lp=l
            go bottom
        endif
        pl=.t.
    endif
    set color to w+/r
    do dados64c with lp
    set color to /w
    do while .t.
        tk=inkey(0)
        do case
            case tk=24
                skip
                if eof()
                    do mensagem with "Fim do Arquivo !",4
                    set color to /w
                    skip -1
                    loop
                else
                    skip -1
                endif
                if lp<18
                    do dados64c with lp
                    lp=lp+1
                    skip
                    set color to w+/r
                    do dados64c with lp
                    set color to /w
                    loop
                else
                    set color to /w
                    do dados64c with lp
                    skip
                    if .not. eof()
                        scroll(12,09,18,76,1)
                        @ lp,09 say "³       ³            ³           ³            ³         ³          ³"
                        @ 18,09 say "³       ³            ³           ³            ³         ³          ³"
                    endif
                    set color to w+/r
                    do dados64c with lp
                    set color to /w
                    loop
                endif
            case tk=5
                if bof()
                    do mensagem with "Inicio do Arquivo !",4
                    set color to /w
                    loop
                endif
                if lp>12
                    if eof()
                        skip -1
                    endif
                    do dados64c with lp
                    lp=lp-1
                    skip -1
                    set color to w+/r
                    do dados64c with lp
                    set color to /w
                    loop
                else
                    set color to /w
                    if eof()
                        skip -1
                    endif
                    do dados64c with lp
                    skip -1
                    if .not. bof()
                        scroll(12,09,18,76,-1)
                        @ lp,09 say "³       ³            ³           ³            ³         ³          ³"
                        @ 18,09 say "³       ³            ³           ³            ³         ³          ³"
                    endif
                    set color to w+/r
                    do dados64c with lp
                    set color to /w
                    loop
                endif
            case tk=13
                save screen
                set color to w+/r
                t = recno()
                return(t)
            case tk=27
                sai=.t.
                t = 0
                return(t)
            other
                loop
        endcase
    enddo
    set color to /w
enddo

**********************************************************
procedure dados64c
* Apresentacao de uma linha de dados para consulta na tela
**********************************************************
parameters linha,valor
iguala()
@ linha,12 say wparcela    pict "@!"
@ linha,19 say wvenci      pict "@d"
@ linha,32 say wvalor_v    pict "@e@z 99,999.99"
@ linha,45 say wpagto      pict "@d"
@ linha,56 say wvalor_p    pict "@e@z 99,999.99"
select 3
return

*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************


*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
