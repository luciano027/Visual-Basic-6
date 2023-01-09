*****************************************************************************
* Programa...: CADENT.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 02/12/98          Data da Ultima Atualizacao:                *
* Objetivo...: ENTRADA DE PRODUTOS NO ESTOQUE                               *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************
select 1
if !net_use("estoque",.f.,"ind01.est","ind02.est")
    return
endif
select 2
if !net_use("fornece",.f.,"ind01.for","ind02.for")
    return
endif
select 3
if !net_use("entrada",.f.,"ind01.ent")
    return
endif
select 4
if !net_use("nf",.f.,"ind01.nf")
    return
endif

do while .t.
    select 1
    inicia()
    select 2
    inicia()
    select 3
    inicia()
    select 4
    inicia()
    do telaent
    do while .t.
        set color to w/b
        @ 23,05 say "<ESC> Para sair.<000> para consulta"
        set color to w,n/bg,,,n/w
        save screen to telafor
        @ 04,13 get wcodigo_for pict "999"
        read
        if readkey() = 12 .or. lastkey() = 27
            select 2
            libera()
            close all
            return
        endif
        if wcodigo_for = "000"
            select 2
            save screen to tela_0
            do pes_for
            restore screen from tela_0
            iguala()
        endif
        set color to /w
        @ 04,13 say wcodigo_for pict "999"
        select 2
        set index to ind01.for
        seek wcodigo_for
        if !found()
            do mensagem with "Codigo do Fornecedor nao cadastrado...",8
            loop
        endif
        iguala()
        set color to /w
        @ 04,13 say wcodigo_for pict "999"
        @ 04,19 say wfornecedor pict "@s20"
        ycodigo_for = wcodigo_for
        yfornecedor = wfornecedor
        exit
    enddo
    save screen to tela_21
    set color to w,n/bg,,,n/w
    @ 05,13 get wnf pict "@!"  when tab_ent(05,23,19,43,4,ycodigo_for)
    read
    set key  5 to
    set key 24 to
    restore screen from tela_21
    set color to r/w
    @ 05,13 say wnf pict "@!"
    ynf = wnf
    if readkey() = 12 .or. lastkey() = 27 .or. empty(wnf)
        libera()
        close all
        return
    endif
    wcs       = "S"
    wtipo_con = "N"
    do while .t.
        select 4
        set index to ind01.nf
        seek ycodigo_for+ynf+"N"
        if found()
            iguala()
            set color to r/w
            @ 04,66 say ynf       pict "@!"
            @ 05,66 say wdata     pict "@d"
            @ 06,13 say wdata     pict "@d"
            @ 06,67 say wvalor    pict "@e@z 99,999.99"
            @ 07,71 say wdesconto pict "@e@z 999.99"
            select 3
            set index to ind01.ent
            seek ycodigo_for+ynf
            iguala()
            ydata       = wdata
            lir         = 10
            wtotal      = 0
            do while wnf = ynf
                if wcs = "S"
                    set color to /w
                    @ lir,03 say wcodigo_est   pict "@!"
                    select 1
                    set index to ind01.est
                    seek wcodigo_est
                    iguala()
                    @ lir,10 say wdescricao       pict "@s20"
                    @ lir,39 say wquantidade   pict "@e 99999.99"
                    @ lir,48 say wpreco_com    pict "@e 99,999.99"
                    lir = lir + 1
                    if lir > 16
                        inkey(.5)
                        do liment
                        lir = 10
                    endif
                    wtotal = wtotal + (wquantidade*wpreco_com)
                endif
                select 3
                skip
                iguala()
            enddo
            wnf         = ynf
            wdata       = ydata
            wcodigo_for = ycodigo_for
            @ 19,48 say wtotal     pict "@e 99,999.99"
            limpa()
            save screen to tela_an
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            Janela2("W",.F.,10,61,16,77,.F.)
            SetColor("N/W,W+/R")
            @ 11,62   prompt "RETORNO        "
            @ 12,62   prompt "INCLUIR        "
            @ 13,62   prompt "CONSULTAR      "
            @ 14,62   prompt "EXTRATO        "
            @ 15,62   prompt "ALTERAR NF     "
            menu to xop
            restore screen from tela_an
            do case
                case xop = 1
                    limpa()
                    exit
                case xop = 2
                    limpa()
                    do while .t.
                        wcodigoest  = 0
                        wquant      = 0
                        wnf         = ynf
                        wdata       = ydata
                        set color to w/b
                        @ 23,05 say "<ENTER> Para Consulta. <ESC> para Sair."
                        set color to n/bg
                        @ lir,03 get wcodigoest pict "999999"
                        read
                        if readkey() = 12 .or. lastkey() = 27
                            exit
                        endif
                        wcodigo_est = strzero(wcodigoest,6,0)
                        if empty(wcodigo_est) .or. wcodigo_est = "000000"
                            save screen to tela_a
                            select 1
                            set index to ind01.est
                            do pes_est
                            iguala()
                            restore screen from tela_a
                        endif
                        set color to /w
                        @ lir,03 say wcodigo_est pict "999999"
                        select 1
                        set index to ind01.est
                        seek wcodigo_est
                        if !found()
                            do mensagem with "Codigo nao Cadastrado...",8
                            loop
                        endif
                        iguala()
                        limpa()
                        set color to  /w
                        @ lir,10 say wdescricao   pict "@s20"
                        set color to w,n/bg,,,n/w
                        @ lir,39 get wquantidade   pict "@e 99999.99"
                        @ lir,48 get wpreco_com    pict "@e 99,999.99"
                        read
                        if wpreco_com = 0
                            do mensagem with " Valor nao pode ficar zerado...",8
                            loop
                        endif
                        wtotal = wtotal + (wquantidade*wpreco_com)
                        set color to /w
                        @ 19,48  say wtotal  pict "@e 99,999.99"
                        if Confirma ("Confirma Dados?")=1    && gravacao dos dados
                            ypreco_com  = wpreco_com
                            rcodigo_est = wcodigo_est
                            rcodigo_for = wcodigo_for
                            wcodigo_est = rcodigo_est
                            select 1
                            set index to ind01.est,ind02.est
                            seek wcodigo_est
                            iguala()
                            if rec_lock()
                                replace saldo      with wsaldo+wquantidade
                                replace preco_com  with ypreco_com
                                replace codigo_for with ycodigo_for
                                unlock
                            endif
                            select 3
                            set index to ind01.ent
                            if add_rec(10)
                                revar()
                            endif
                            set color to /w
                            @ 19,48 say wtotal     pict "@e 99,999.99"
                            lir = lir + 1
                            if lir > 16
                                liment()
                                lir = 10
                            endif
                            loop
                        else
                            wtotal = wtotal - (wquantidade*wpreco_com)
                            set color to /w
                            @ 19,48 say wtotal     pict "@e 99,999.99"
                            loop
                        endif
                    enddo
                    limpa()
                    loop
                case xop = 3
                    save screen to tela09
                    select 3
                    index on nf to &zmicro for nf = wnf .and. codigo_for = wcodigo_for
                    set index to &zmicro
                    go top
                    liment()
                    do pescon_ent with 3
                    reg = recno()
                    locate for reg = recno()
                    set color to w/b
                    dt = quadro(09,20,15,75)
                    @ 10,22 say "Codigo....:"
                    @ 11,22 say "descricao.:"
                    @ 12,22 say "Quantidade:"
                    @ 13,22 say "Pr.Compra.:"
                    wcodigo_est = codigo_est
                    yquantidade = quantidade
                    wquantidade = quantidade
                    select 1
                    set index to ind01.est
                    seek wcodigo_est
                    iguala()
                    wvalor  = wpreco_com * wquantidade
                    @ 10,33  say wcodigo_est  pict "@!"
                    @ 11,33  say wdescricao      pict "!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
                    @ 12,33  say wquantidade  pict "@e 99,999.99"
                    @ 13,33  say wvalor       pict "@e 99,999.99"
                    limpa()
                    set color to  w/b
                    @ 14,22      prompt "RETORNO "
                    @ 14,col()+2 prompt "ALTERAR "
                    @ 14,col()+2 prompt "EXCLUIR "
                    menu to qop
                    do case
                        case qop = 1
                            set color to w/b
                            @ 23,01 say space(79)
                            restore screen from tela09
                            loop
                        case qop = 2
                            ycodigo_est = wcodigo_est
                            yquantidade = wquantidade
                            ydata       = wdata
                            set color to w/b
                            @ 10,33 say wcodigo_est  pict "999999"
                            @ 11,33 say wdescricao   pict "!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
                            set color to w,n/bg,,,n/w
                            @ 12,33 get wquantidade pict "@e 99,999.99"
                            read
                            p = "S"
                            wtotal = wtotal + (wpreco_com  * wquantidade)
                            limpa()
                            @ 23,03 say "Confirma Dados <S/N> ?" get p pict "@!" valid p $ "SN"
                            read
                            limpa()
                            if upper(p) = "S"
                                zquant = wquantidade
                                select 1
                                set index to ind01.est,ind02.est
                                seek wcodigo_est
                                iguala()
                                if rec_lock()
                                    replace saldo    with (wsaldo-yquantidade)+wquantidade
                                    unlock
                                endif
                                select 3
                                set index to ind01.ent
                                go reg
                                iguala()
                                if rec_lock()
                                    replace quantidade with zquant
                                    unlock
                                endif
                                lir = lir + 1
                                if lir > 16
                                    liment()
                                    lir = 10
                                endif
                                restore screen from tela09
                                loop
                            endif
                            restore screen from tela09
                            loop
                        case qop = 3
                            yquant      = wquantidade
                            if Confirma ("Confirma Exclusao?")=1    && gravacao dos dados
                                select 1
                                set index to ind01.est,ind02.est
                                seek wcodigo_est
                                iguala()
                                if rec_lock()
                                    replace saldo    with wsaldo-yquant
                                    unlock
                                endif
                                select 3
                                set index to ind01.ent
                                go reg
                                if rec_lock()
                                    delete
                                endif
                                restore screen from tela09
                                liment()
                                loop
                            endif
                        other
                            restore screen from tela09
                            loop
                    endcase
                    restore screen from tela09
                    liment()
                    loop
                case xop = 4
                    save screen to tela12
                    do while .t.
                        select 3
                        index on nf to &zmicro for nf = ynf .and. codigo_for = wcodigo_for
                        set index to &zmicro
                        limpa()
                        set color to w+/r
                        dt = quadro(15,18,17,35)
                        set color to *+w/r
                        @ 16,20 say "Aguarde..."
                        set color to w/b
                        set print to arq.txt
                        set devi to print
                        set cons off
                        wpag = 1
                        geral = 0
                        do while inkey()#27 .and. !eof()
                            @ prow(),02    say zempresa  pict "@!"
                            @ prow()+1,02  say zendereco pict "@!"
                            @ prow()+2,02  say "Extrato Entrada de Mercadorias"
                            @ prow(),72    say "Pag."
                            @ prow(),76    say wpag pict "999"
                            @ prow()+2,02  say "Fornecedor..:"
                            @ prow(),pcol()+1 say ycodigo_for pict "@!"
                            @ prow(),pcol()+1 say yfornecedor pict "@!"
                            @ prow()+1,02     say "Data........:"
                            @ prow(),pcol()+1 say wdata pict "@d"
                            @ prow()+2,02  say"Codigo    Especificacao                  Quantidade Preco      Valor     "
                            @ prow()+1,02  say"~~~~~~~~~ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~ ~~~~~~~~~~ ~~~~~~~~~~"
                            pagina = 0
                            do while prow()<50 .and. !eof()
                                iguala()
                                wquant = wquantidade
                                select 1
                                seek wcodigo_est
                                iguala()
                                @ prow()+1,02 say wcodigo_est  pict "@!"
                                @ prow(),12   say wdescricao    pict "!!!!!!!!!!!!!!!!!!!!!!!!"
                                @ prow(),44   say wquant        pict "@e@z 99999.99"
                                @ prow(),55   say wpreco_com    pict "@e@z 99,999.99"
                                compras = (wquant * wpreco_com)
                                @ prow(),65   say compras     pict "999,999.99"
                                wlin = 1
                                pagina = pagina + compras
                                select 3
                                skip
                                esc = inkey()
                                if esc = 12
                                    @ prow()+2,40 say "***********  Impressao Cancelada..."
                                    close all
                                    return
                                endif
                            enddo
                            geral = geral + pagina
                            @ prow()+1,01  say "-------------------------------------------------------------------------"
                            @ prow()+1,03  say "Total na Pagina........."
                            @ prow(),63    say pagina pict "@e@z 999,999.99"
                            @ prow()+1,03  say "Total a Pagar .........."
                            @ prow(),63    say geral  pict "@e@z 999,999.99"
                            @ prow()+1,01  say "-------------------------------------------------------------------------"
                            @ prow()+2,01  say "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
                            @ prow()+1,01  say "              Sistema de Administracao Estoque - Versao 1.0"
                            @ prow(),02    say date() pict "@d"
                            @ prow(),69    say time() pict "99:99"
                            eject
                            wpag = wpag + 1
                        enddo
                        set cons on
                        set devi to screen
                        set print off
                        set print to
                        lertexto(zdireto)
                        exit
                    enddo
                    restore screen from tela12
                    loop
                case xop = 5
                    set color to w,n/bg,,,n/w
                    @ 05,66 get wdata     pict "@d"
                    @ 06,67 get wvalor    pict "@e@z 99,999.99"
                    @ 07,71 get wdesconto pict "@e@z 999.99"
                    read
                    select 4
                    set index to ind01.nf
                    seek ycodigo_for+ynf+"N"
                    if found()
                        if rec_lock()
                            revar()
                            unlock
                        endif
                    endif
                    loop
                other
                    loop
            endcase
            loop
        else
            limpa()
            lir = 10
            t = "S"
            total = 0
            wdata = date()
            set color to n/bg
            @ 06,13 get wdata pict "@d"
            read
            if readkey() = 12 .or. lastkey() = 27
                saida = .t.
                return
            endif
            do while .t.
                wcodigoest  = 0
                wquantidade = 0
                set color to w/b
                @ 23,05 say "<ENTER> Para Consulta. <ESC> para Sair."
                set color to n/bg
                @ lir,03 get wcodigoest pict "999999"
                read
                if readkey() = 12 .or. lastkey() = 27
                    exit
                endif
                wcodigo_est = strzero(wcodigoest,6,0)
                if empty(wcodigo_est) .or. wcodigo_est = "000000"
                    save screen to tela_a
                    select 1
                    set index to ind01.est
                    do pes_est
                    iguala()
                    restore screen from tela_a
                endif
                set color to /w
                @ lir,03 say wcodigo_est pict "@!"
                select 1
                set index to ind01.est
                seek wcodigo_est
                if !found()
                    do mensagem with "Codigo nao Cadastrado...",8
                    loop
                endif
                iguala()
                limpa()
                set color to  /w
                @ lir,10 say wdescricao    pict "@s20"
                set color to w,n/bg,,,n/w
                @ lir,39 get wquantidade   pict "@e@z 99999.99"
                @ lir,48 get wpreco_com    pict "@e@z 99,999.99"
                read
                if wpreco_com  = 0
                    do mensagem with " Valor nao pode ficar zerado...",8
                    loop
                endif
                total = total + (wquantidade*wpreco_com)
                set color to /w
                @ 19,48  say total         pict "@e 99,999.99"
                if Confirma ("Confirma Dados?")=1    && gravacao dos dados
                    ypreco_com    = wpreco_com
                    rcodigo_est = wcodigo_est
                    select 1
                    set index to ind01.est,ind02.est
                    seek wcodigo_est
                    if found()
                        iguala()
                        if rec_lock()
                            replace saldo      with wsaldo + wquantidade
                            replace preco_com  with ypreco_com
                            replace codigo_for with ycodigo_for
                            unlock
                        endif
                    endif
                    wcodigo_for = ycodigo_for
                    select 3
                    set index to ind01.ent
                    if add_rec(10)
                        revar()
                    endif
                    select 4
                    set index to ind01.nf
                    seek ycodigo_for+ynf+"N"
                    if !found()
                        if add_rec(10)
                            revar()
                        endif
                    endif
                    set color to /w
                    @ 19,48 say total      pict "@e 99,999.99"
                    lir = lir + 1
                    if lir > 16
                        liment()
                        lir = 10
                    endif
                else
                    total = total - (wquantidade*wpreco_com)
                    set color to /w
                    @ 19,48 say total      pict "@e 99,999.99"
                endif
            enddo
        endif
    enddo
    loop
enddo
close all
return

procedure pescon_ent
para sele
select &sele
go top
sai=.f.
lp=10
tk=0
ul=22
set color to w/b
@ 23,03 say "Sobe :"+chr(24)+" Desce :"+chr(25)+" <ENTER> Confirma  <ESC> Sai "
set color to /w
do while .not. sai
    l=10   && primeira linha da tela l = 8
    do while l<=17 .and. .not. eof()
        do dadosent with l
        skip
        l=l+1
    enddo
    nl=l
    do while nl<=17
        nl=nl+1
    enddo
    l=l-1
    ul=l
    if tk<0 .or. tk=22
        goto rec
        lp=10
    else
        skip lp-(l+1)
    endif
    pl=.f.
    if l<22
        if lp>10
            lp=l
            go bottom
        endif
        pl=.t.
    endif
    set color to w+/r
    do dadosent with lp
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
                if lp<15
                    do dadosent with lp
                    lp=lp+1
                    skip
                    set color to w+/r
                    do dadosent with lp
                    set color to /w
                    loop
                else
                    set color to /w
                    do dadosent with lp
                    skip
                    if .not. eof()
                        scroll(10,02,17,57,1)
                        @ lp,02 say "³      ³                            ³        ³         ³"
                        @ 17,02 say "³      ³                            ³        ³         ³"
                    endif
                    set color to w+/r
                    do dadosent with lp
                    set color to /w
                    loop
                endif
            case tk=5
                if bof()
                    do mensagem with "Inicio do Arquivo !",4
                    set color to /w
                    loop
                endif
                if lp>10
                    if eof()
                        skip -1
                    endif
                    do dadosent with lp
                    lp=lp-1
                    skip -1
                    set color to w+/r
                    do dadosent with lp
                    set color to /w
                    loop
                else
                    set color to /w
                    if eof()
                        skip -1
                    endif
                    do dadosent with lp
                    skip -1
                    if .not. bof()
                        scroll(10,02,17,57,-1)
                        @ lp,02 say "³      ³                            ³        ³         ³"
                        @ 17,02 say "³      ³                            ³        ³         ³"
                    endif
                    set color to w+/r
                    do dadosent with lp
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
procedure dadosent
    * Apresentacao de uma linha de dados para consulta na tela
    **********************************************************
parameters linha,valor
iguala()
@ linha,03  say wcodigo_est   pict "@!"
select 1
set index to ind01.est
seek wcodigo_est
iguala()
@ linha,10 say wdescricao    pict "@s20"
@ linha,39 say wquantidade   pict "@e 99999.99"
@ linha,48 say wpreco_com    pict "@e 99,999.99"
select 3
return

procedure telaent
set color to w/b
@ 23,05 say "<ESC> Para sair."
Janela2("W",.F.,03,01,21,58,.F.)
SetColor("N/W,W+/R")
set color to w/r
@ 03,19 say "  ENTRADA NO ESTOQUE  "
SetColor("N/W,W+/R")
@ 04,02 say "Fornecedor:"
@ 05,02 say "Documento.:"
@ 06,02 say "Data......:"
@ 07,02 say "ÚÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄ¿"
@ 08,02 say "³Codigo³         Descricao          ³Quantid ³  Valor  ³"
@ 09,02 say "ÃÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄ´"
@ 10,02 say "³      ³                            ³        ³         ³"
@ 11,02 say "³      ³                            ³        ³         ³"
@ 12,02 say "³      ³                            ³        ³         ³"
@ 13,02 say "³      ³                            ³        ³         ³"
@ 14,02 say "³      ³                            ³        ³         ³"
@ 15,02 say "³      ³                            ³        ³         ³"
@ 16,02 say "³      ³                            ³        ³         ³"
@ 17,02 say "³      ³                            ³        ³         ³"
@ 18,02 say "ÃÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄ´"
@ 19,02 say "³                            T O T A L                 ³"
@ 20,02 say "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ"
set color to n/w
dt = quadro(03,60,08,77)
set color to w/r
@ 03,65 say "  NF   "
set color to  /w
@ 04,61 say "NF..:"
@ 05,61 say "Data:"
@ 06,61 say "Valor:"
@ 07,61 say "Desconto:"
return

return

procedure liment
set color to /w
@ 10,02 say "³      ³                            ³        ³         ³"
@ 11,02 say "³      ³                            ³        ³         ³"
@ 12,02 say "³      ³                            ³        ³         ³"
@ 13,02 say "³      ³                            ³        ³         ³"
@ 14,02 say "³      ³                            ³        ³         ³"
@ 15,02 say "³      ³                            ³        ³         ³"
@ 16,02 say "³      ³                            ³        ³         ³"
@ 17,02 say "³      ³                            ³        ³         ³"
return

**********************pesquisa da banco********************************
function tab_ent(x1,y1,x2,y2,sel,docu)
sele &sel
index on dtos(data) to &zmicro for codigo_for = docu .and. tipo_con = "N"
set index to &zmicro
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
@ linha,y1+1     say "NF"
@ linha,col()+11  say "Data"
linha++
@ linha,y1+1 to linha,(y2+6)-1
linha++
for x = 1 to 10
    @ linha,yy1+1    say nf        pict "@!"
    @ linha,col()+5  say data      pict "@d"
    linha++
    skip
next
set key 5  to sobe_ent
set key 24 to desce_ent
return .t.

function sobe_ent
linha = 0
@ xx1,yy1 clear to xx2,yy2
dt = quadro(xx1,yy1,xx2,yy2+6)
linha = xx1+1
@ linha,yy1+1    say "NF"
@ linha,col()+11  say "Data"
linha++
@ linha,yy1+1 to linha,(yy2+6)-1
linha++
skip -10
for x = 1 to 10
    @ linha,yy1+1    say nf        pict "@!"
    @ linha,col()+5  say data      pict "@d"
    linha++
    skip
next
return .t.

function desce_ent
@ xx1,yy1 clear to xx2,yy2
dt = quadro(xx1,yy1,xx2,yy2+6)
linha = xx1+1
@ linha,yy1+1    say "NF"
@ linha,col()+11  say "Data"
linha++
@ linha,yy1+1 to linha,(yy2+6)-1
linha++
if .not. eof()
    for x = 1 to 10
        @ linha,yy1+1    say nf        pict "@!"
        @ linha,col()+5  say data      pict "@d"
        linha++
        skip
    next
endif
return .t.




    *****************************************************************************
    *                          F   I   M                                        *
    *****************************************************************************

