*****************************************************************************
* Programa...: VISTAEXC.PRG                                                 *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......:                   Data da Ultima Atualizacao:                *
* Objetivo...: CONSULTA VENDA A VISTA                                       *
* Sistema....: SISPAN - Sistema de Panifica‡Æo                              *
*****************************************************************************
qop = 1
do while .t.
    limpa()
    save screen to tela001
    set color to w/bg,/w+,,,n
    @ 13,22      prompt " RETORNO "
    @ 13,col()+2 prompt " Cancelar Item "
    menu to qop
    do case
        case qop = 1
            set color to w/b
            @ 23,01 say space(79)
            return
        case qop = 2
            restore screen from tela001
            if Confirma ("Confirma Delecao <S/N> ?")=1    && gravacao dos dados
                qual= witem
                if val(qual)>0
                    par_fiscal= strzero(val(qual),3)
                    enviacodigo(4)  &&* Cancelamento de Item do Cupom Sendo Confeccionado
                    total=total-totite[val(qual)]
                endif
                select 3
                set index to ind01.vis,ind02.vis,ind03.vis
                go reg
                iguala()
                delete
                wnf = "V"+wticket
                select 4
                set index to ind01.sai,ind02.sai,ind03.sai
                seek wcodigo_est+wnf
                delete
                select 1
                set index to ind01.est,ind02.est
                seek wcodigo_est
                iguala()
                wsaldo  = wsaldo  + wquantidade
                if rec_lock()
                    replace saldo with wsaldo
                    unlock
                endif
                do mensagem with "Registro Deletado...",8
                limpas()
                return
            endif
        other
            return
    endcase
enddo
return

procedure pescon_31
para sele
select &sele
go top
sai=.f.
lp=7
tk=0
ul=14
set color to w/B
@ 23,03 say "Sobe :"+chr(24)+" Desce :"+chr(25)+" <ENTER> Confirma  <ESC> Sai "
set color to /w
do while .not. sai
    l=7   && primeira linha da tela l = 8
    do while l<=14 .and. .not. eof()
        do dadoS61C with l
        skip
        l=l+1
    enddo
    nl=l
    do while nl<=14
        nl=nl+1
    enddo
    l=l-1
    ul=l
    if tk<0 .or. tk=14
        goto rec
        lp=7
    else
        skip lp-(l+1)
    endif
    pl=.f.
    if l<14
        if lp>7
            lp=l
            go bottom
        endif
        pl=.t.
    endif
    set color to w+/r
    do dadoS61C with lp
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
                if lp<14
                    do dadoS61C with lp
                    lp=lp+1
                    skip
                    set color to w+/r
                    do dadoS61C with lp
                    set color to /w
                    loop
                else
                    set color to n/w
                    do dados61c with lp
                    skip
                    if .not. eof()
                        scroll(07,02,14,53,1)
                        @ lp,02 say "³      ³                   ³       ³       ³       ³"
                        @ 14,02 say "³      ³                   ³       ³       ³       ³"
                    endif
                    set color to w+/r
                    do dados61c with lp
                    set color to n/w
                    loop
                endif
            case tk=5
                if bof()
                    do mensagem with "Inicio do Arquivo !",4
                    set color to /w
                    loop
                endif
                if lp>7
                    if eof()
                        skip -1
                    endif
                    do dados61c with lp
                    lp=lp-1
                    skip -1
                    set color to w+/r
                    do dados61c with lp
                    set color to n/w
                    loop
                else
                    set color to n/w
                    if eof()
                        skip -1
                    endif
                    do dados61c with lp
                    skip -1
                    if .not. bof()
                        scroll(07,02,14,53,-1)
                        @ lp,02 say "³      ³                   ³       ³       ³       ³"
                        @ 14,02 say "³      ³                   ³       ³       ³       ³"
                    endif
                    set color to w+/r
                    do dados61c with lp
                    set color to n/w
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
procedure dadoS61C
* Apresentacao de uma linha de dados para consulta na tela
**********************************************************
parameters linha,valor
iguala()
@ linha,03 say wcodigo_est pict "999999"
select 1
set index to ind01.est
seek wcodigo_est
iguala()
@ linha,10 say wdescricao pict "@s19"
select 3
@ linha,30 say wquantidade pict "@e@z 9999.999"
@ linha,39 say wpreco_ven  pict "@e@z 999.99"
wcompra = (wpreco_ven * wquantidade)
@ linha,46 say wcompra    pict "@e@z  999.99"
select 3
return

procedure tela3
set color to n/w
dt = quadro(04,54,11,77)
set color to w/r
@ 04,56 say "      CAIXA      "
set color to  /w
@ 04,02 say "ÚÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄ¿"
@ 05,02 say "³Codigo³     Descricao     ³ Quant.³  Unit.³ Total ³"
@ 06,02 say "ÃÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄ´"
@ 07,02 say "³      ³                   ³       ³       ³       ³"
@ 08,02 say "³      ³                   ³       ³       ³       ³"
@ 09,02 say "³      ³                   ³       ³       ³       ³"
@ 10,02 say "³      ³                   ³       ³       ³       ³"
@ 11,02 say "³      ³                   ³       ³       ³       ³"
@ 12,02 say "³      ³                   ³       ³       ³       ³"
@ 13,02 say "³      ³                   ³       ³       ³       ³"
@ 14,02 say "³      ³                   ³       ³       ³       ³"
@ 15,02 say "³      ³                   ³       ³       ³       ³"
@ 16,02 say "ÃÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄ´"
@ 17,02 say "³                                                  ³"
@ 18,02 say "³                                                  ³"
@ 19,02 say "³                                                  ³"
@ 20,02 say "³                                                  ³"
@ 21,02 say "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ"
set color to n/w
dt = quadro(04,54,11,77)
set color to w/r
@ 04,58 say "      CAIXA      "
set color to  /w
@ 05,55 say "Ticket..:"
@ 06,55 say "Data....:"
@ 07,55 say "Compras.:"
@ 08,55 say "Desconto:"
@ 09,55 say "Dinheiro:"
@ 10,55 say "Troco...:"
dt = quadro(14,54,17,77)
set color to w/r
@ 14,58 say "    VENDEDOR     "
set color to  /w
@ 15,55 say "Codigo..:"
@ 16,55 say "Nome....:"
return

procedure limpas
set color to /w
@ 07,02 say "³      ³                   ³       ³       ³       ³"
@ 08,02 say "³      ³                   ³       ³       ³       ³"
@ 09,02 say "³      ³                   ³       ³       ³       ³"
@ 10,02 say "³      ³                   ³       ³       ³       ³"
@ 11,02 say "³      ³                   ³       ³       ³       ³"
@ 12,02 say "³      ³                   ³       ³       ³       ³"
@ 13,02 say "³      ³                   ³       ³       ³       ³"
@ 14,02 say "³      ³                   ³       ³       ³       ³"
@ 15,02 say "³      ³                   ³       ³       ³       ³"
lir = 7
return

*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
