*****************************************************************************
* Programa...: CADCLI.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 02/12/98          Data da Ultima Atualizacao:                *
* Objetivo...: CADASTRO DE CLIENTES                                         *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************
select 1
if !net_use("apagar",.f.,"ind01.apa")
    return
endif
select 2
if !net_use("fornece",.f.,"ind01.for","ind02.for")
    return
endif

do while .t.
    select 1
    inicia()
    select 2
    inicia()
    Janela2("W",.F.,04,08,20,70,.F.)
    set color to w/r
    @ 04,30 say " BAIXA CONTAS A PAGAR  "
    SetColor("N/W,W+/R")
    @ 05,09 say "Fornecedor..:"
    @ 06,09 say "Documento...:"
    @ 07,09 say "Empresa.....:"
    @ 08,09 say "ÚÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄ¿"
    @ 09,09 say "³Parcela³ Vencimento ³   Valor   ³ Data Pagto ³   Valor   ³"
    @ 10,09 say "ÃÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄ´"
    @ 11,09 say "³ 001   ³            ³           ³            ³           ³"
    @ 12,09 say "³ 002   ³            ³           ³            ³           ³"
    @ 13,09 say "³ 003   ³            ³           ³            ³           ³"
    @ 14,09 say "³ 004   ³            ³           ³            ³           ³"
    @ 15,09 say "³ 005   ³            ³           ³            ³           ³"
    @ 16,09 say "³ 006   ³            ³           ³            ³           ³"
    @ 17,09 say "ÃÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄ´"
    @ 18,09 say "³  Total a Pagar :                 Total Pago.:           ³"
    @ 19,09 say "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ"
    limpa()
    save screen to telat
    do while .t.
        set color to w/b
        @ 23,05 say "<ESC> Para sair. <0000> para Consulta"
        set color to w,n/bg,,,n/w
        @ 05,23 get wcodigo_for pict "9999"
        read
        if readkey() = 12  .or. lastkey() = 27
            close all
            return
        endif
        if empty(wcodigo_for) .or. wcodigo_for = "0000"
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
        @ 05,23 say wcodigo_for pict "9999"
        select 2
        set index to ind01.for
        seek wcodigo_for
        if !found()
            do mensagem with "Codigo do Fornecedor nao cadastrado...",8
            loop
        endif
        iguala()
        set color to r/w
        @ 05,23 say wcodigo_for  pict "@!"
        @ 07,23 say wempresa     pict "@s20"
        exit
    enddo
    save screen to tela_21
    set color to w,n/bg,,,n/w
    @ 06,23 get wnf pict "@!"  when tab_doc(07,23,20,35,1,wcodigo_for)
    read
    set key  5 to
    set key 24 to
    restore screen from tela_21
    set color to r/w
    @ 06,23 say wnf pict "@!"
    ztipo = wcodigo_for
    znf   = wnf
    do while .t.
        select 1
        set index to ind01.apa
        seek ztipo+znf
        if found()
            iguala()
            set color to r/w
            @ 05,23 say wcodigo_for pict "9999"
            @ 07,23 say wempresa    pict "@s20"
            @ 06,23 say wnf         pict "@!"
            @ 11,19 say wdata_v1    pict "@d"
            @ 11,32 say wvalor_v1   pict "@e@z 99,999.99"
            @ 11,45 say wdata_p1    pict "@d"
            @ 11,57 say wvalor_p1   pict "@e@z 99,999.99"
            @ 12,19 say wdata_v2    pict "@d"
            @ 12,32 say wvalor_v2   pict "@e@z 99,999.99"
            @ 12,45 say wdata_p2    pict "@d"
            @ 12,57 say wvalor_p2   pict "@e@z 99,999.99"
            @ 13,19 say wdata_v3    pict "@d"
            @ 13,32 say wvalor_v3   pict "@e@z 99,999.99"
            @ 13,45 say wdata_p3    pict "@d"
            @ 13,57 say wvalor_p3   pict "@e@z 99,999.99"
            @ 14,19 say wdata_v4    pict "@d"
            @ 14,32 say wvalor_v4   pict "@e@z 99,999.99"
            @ 14,45 say wdata_p4    pict "@d"
            @ 14,57 say wvalor_p4   pict "@e@z 99,999.99"
            @ 15,19 say wdata_v5    pict "@d"
            @ 15,32 say wvalor_v5   pict "@e@z 99,999.99"
            @ 15,45 say wdata_p5    pict "@d"
            @ 15,57 say wvalor_p5   pict "@e@z 99,999.99"
            @ 16,19 say wdata_v6    pict "@d"
            @ 16,32 say wvalor_v6   pict "@e@z 99,999.99"
            @ 16,45 say wdata_p6    pict "@d"
            @ 16,57 say wvalor_p6   pict "@e@z 99,999.99"
            wtotaa = wvalor_v1+wvalor_v2+wvalor_v3+wvalor_v4+wvalor_v5+wvalor_v6
            wtotap = wvalor_p1+wvalor_p2+wvalor_p3+wvalor_p4+wvalor_p5+wvalor_p6
            set color to r/w
            @ 18,28 say wtotaa    pict "@e@z 99,999.99"
            @ 18,57 say wtotap    pict "@e@z 99,999.99"
            limpa()
            save screen to tela_a
            set color to w/b
            @ 23,05 say "<ESC> Para sair."
            SetColor("W/W")
            Janela2("W",.F.,03,66,08,75,.F.)
            SetColor("N/W,W+/R")
            @ 04,67   prompt "RETORNO "
            @ 05,67   prompt "ALTERAR "
            @ 06,67   prompt "EXCLUIR "
            @ 07,67   prompt "BAIXA   "
            menu to xop
            restore screen from tela_a
            do case
                case xop = 1
                    limpa()
                    exit
                case xop = 2
                    save screen to telat
                    set color to w,n/bg,,,n/w
                    @ 11,19 get wdata_v1    pict "@d"
                    @ 11,32 get wvalor_v1   pict "@e@z 99,999.99"
                    @ 11,45 get wdata_p1    pict "@d"
                    @ 11,57 get wvalor_p1   pict "@e@z 99,999.99"
                    read
                    @ 12,19 get wdata_v2    pict "@d"
                    @ 12,32 get wvalor_v2   pict "@e@z 99,999.99"
                    @ 12,45 get wdata_p2    pict "@d"
                    @ 12,57 get wvalor_p2   pict "@e@z 99,999.99"
                    read
                    if !empty(wdata_v2)
                        @ 13,19 get wdata_v3    pict "@d"
                        @ 13,32 get wvalor_v3   pict "@e@z 99,999.99"
                        @ 13,45 get wdata_p3    pict "@d"
                        @ 13,57 get wvalor_p3   pict "@e@z 99,999.99"
                        read
                        if !empty(wdata_v3)
                            @ 14,19 get wdata_v4    pict "@d"
                            @ 14,32 get wvalor_v4   pict "@e@z 99,999.99"
                            @ 14,45 get wdata_p4    pict "@d"
                            @ 14,57 get wvalor_p4   pict "@e@z 99,999.99"
                            read
                            if !empty(wdata_v4)
                                @ 15,19 get wdata_v5    pict "@d"
                                @ 15,32 get wvalor_v5   pict "@e@z 99,999.99"
                                @ 15,45 get wdata_p5    pict "@d"
                                @ 15,57 get wvalor_p5   pict "@e@z 99,999.99"
                                read
                                if !empty(wdata_v5)
                                    @ 16,19 get wdata_v6    pict "@d"
                                    @ 16,32 get wvalor_v6   pict "@e@z 99,999.99"
                                    @ 16,45 get wdata_p6    pict "@d"
                                    @ 16,57 get wvalor_p6   pict "@e@z 99,999.99"
                                    read
                                endif
                            endif
                        endif
                    endif
                    limpa()
                    if Confirma (" Dados Corretos? ")=1    && gravacao dos dados     select 1
                        select 1
                        set index to ind01.apa
                        seek ztipo+znf
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
                    if Confirma ("Confirma Exclusao ? Todo o Documento sera excluido? ")=1    && gravacao dos dados
                        select 1
                        set index to ind01.apa
                        seek ztipo+znf
                        if found()
                            if rec_lock()
                                delete
                            endif
                        endif
                        exit
                    endif
                    loop
                case xop = 4
                    save screen to telat
                    set color to w,n/bg,,,n/w
                    @ 11,45 get wdata_p1    pict "@d"
                    @ 11,57 get wvalor_p1   pict "@e@z 99,999.99"
                    read
                    @ 12,45 get wdata_p2    pict "@d"
                    @ 12,57 get wvalor_p2   pict "@e@z 99,999.99"
                    read
                    if !empty(wdata_p2)
                        @ 13,45 get wdata_p3    pict "@d"
                        @ 13,57 get wvalor_p3   pict "@e@z 99,999.99"
                        read
                        if !empty(wdata_p3)
                            @ 14,45 get wdata_p4    pict "@d"
                            @ 14,57 get wvalor_p4   pict "@e@z 99,999.99"
                            read
                            if !empty(wdata_p4)
                                @ 15,45 get wdata_p5    pict "@d"
                                @ 15,57 get wvalor_p5   pict "@e@z 99,999.99"
                                read
                                if !empty(wdata_p5)
                                    @ 16,45 get wdata_p6    pict "@d"
                                    @ 16,57 get wvalor_p6   pict "@e@z 99,999.99"
                                    read
                                endif
                            endif
                        endif
                    endif
                    limpa()
                    if Confirma (" Dados Corretos? ")=1    && gravacao dos dados     select 1
                        select 1
                        set index to ind01.apa
                        seek ztipo+znf
                        if found()
                            if rec_lock()
                                revar()
                                unlock
                            endif
                        endif
                    endif
                    loop
            endcase
            loop
        else
            do mensagem with "NF nao cadastrada...",8
            exit
        endif
    enddo
enddo
close all
return


**********************pesquisa da banco********************************
function tab_doc(x1,y1,x2,y2,sel,docu)
sele &sel
index on dtos(data_v1) to indcon.tip for;
codigo_for = docu .and. data_p1 = ctod("  /  /  ") .or.;
codigo_for = docu .and. data_p2 = ctod("  /  /  ") .or.;
codigo_for = docu .and. data_p3 = ctod("  /  /  ") .or.;
codigo_for = docu .and. data_p4 = ctod("  /  /  ") .or.;
codigo_for = docu .and. data_p5 = ctod("  /  /  ")
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
linha++
@ linha,y1+1 to linha,(y2+6)-1
linha++
for x = 1 to 10
    @ linha,yy1+1    say nf        pict "@!"
    linha++
    skip
next
set key 5  to sobe_doc
set key 24 to desce_doc
return .t.

function sobe_doc
linha = 0
@ xx1,yy1 clear to xx2,yy2
dt = quadro(xx1,yy1,xx2,yy2+6)
linha = xx1+1
@ linha,yy1+1     say "Documento"
linha++
@ linha,yy1+1 to linha,(yy2+6)-1
linha++
skip -10
for x = 1 to 10
    @ linha,yy1+1    say nf        pict "@!"
    linha++
    skip
next
return .t.

function desce_doc
@ xx1,yy1 clear to xx2,yy2
dt = quadro(xx1,yy1,xx2,yy2+6)
linha = xx1+1
@ linha,yy1+1     say "Documento"
linha++
@ linha,yy1+1 to linha,(yy2+6)-1
linha++
if .not. eof()
    for x = 1 to 10
        @ linha,yy1+1    say nf        pict "@!"
        linha++
        skip
    next
endif
return .t.

*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
