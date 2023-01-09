save screen to tela09
do while .t.
    save screen to telaat
    Janela2("W",.F.,14,60,19,72,.F.)
    SetColor("N/W,W+/R")
    @ 15,61 prompt "Retorno    "
    @ 16,61 prompt "Dinheiro   "
    @ 17,61 prompt "Cheque     "
    @ 18,61 prompt "Cartao     "
    menu to wop
    restore screen from telaat
    do case
        case wop = 1
            limpa()
            exit
        case wop = 2
            limpa()
            save screen to telawop2
            @ 23,05 say "<ESC> Para sair."
            set color to  w/b
            quadro(16,05,20,75)
            @ 17,06 say "Data.....:                  Documento:"
            @ 18,06 say "Descricao:"
            @ 19,06 say "Valor....:"
            set color to w/b
            whistorico = "Venda a Vista - Dinheiro "
            @ 17,16  say wdata_ti    pict "@d"
            @ 17,44  say wticket     pict "@!"
            @ 18,16  say whistorico  pict "@!"
            @ 19,16  say wtotal      pict "@e@z 99,999.99"
            p = "S"
            limpa()
            @ 23,03 say "Confirma Dados <S/N> ?" get p pict "@!" valid p $ "SN"
            read
            limpa()
            if upper(p) = "S"
                select 5
                set index to ind01.tic
                seek wticket
                if found()
                    replace pagamento with "S"
                    replace total_ti  with wtotal
                    replace historico with whistorico
                    replace tipo_ti   with "D"
                endif
            endif
            restore screen from telawop2
            exit
        case wop = 3
            save scree to telawop3
            wbanco   = space(03)
            wagencia = space(04)
            wconta   = space(10)
            wcheque  = space(10)
            limpa()
            @ 23,05 say "<ESC> Para sair."
            set color to  w/b
            quadro(16,05,20,77)
            @ 17,06 say "Data.....:                  Documento:"
            @ 18,06 say "Banco....:       Agencia..:      Conta:             Cheque:"
            @ 19,06 say "Valor....:"
            set color to w/b
            @ 17,16  say wdata_ti    pict "@d"
            @ 17,44  say wticket     pict "@!"
            @ 19,16  say wtotal      pict "@e@z 99,999.99"
            @ 18,16  get wbanco      pict "@!"
            @ 18,33  get wagencia    pict "@!"
            @ 18,45  get wconta      pict "@!"
            @ 18,65  get wcheque     pict "@!"
            read
            p = "S"
            limpa()
            @ 23,03 say "Confirma Dados <S/N> ?" get p pict "@!" valid p $ "SN"
            read
            limpa()
            if upper(p) = "S"
                select 5
                set index to ind01.tic
                seek wticket
                if found()
                    replace pagamento with "S"
                    replace total_ti  with wtotal
                    replace historico with "Pagamento a Vista - Cheque"
                    replace banco     with wbanco
                    replace agencia   with wagencia
                    replace conta     with wconta
                    replace cheque    with wcheque
                    replace tipo_ti   with "Q"
                endif
            endif
            restore screen from telawop3
            exit
        case wop = 4
            save screen to telawop4
            wcartao  = space(10)
            wnumero  = space(18)
            limpa()
            @ 23,05 say "<ESC> Para sair."
            set color to  w/b
            quadro(16,05,20,77)
            @ 17,06 say "Data.....:                  Documento:"
            @ 18,06 say "Cartao...:                     Numero:"
            @ 19,06 say "Valor....:"
            set color to w/b
            @ 17,16  say wdata_ti    pict "@d"
            @ 17,44  say wticket     pict "@!"
            @ 19,16  say wtotal      pict "@e@z 99,999.99"
            @ 18,16  get wcartao     pict "@!"
            @ 18,45  get wnumero     pict "@!"
            read
            p = "S"
            limpa()
            @ 23,03 say "Confirma Dados <S/N> ?" get p pict "@!" valid p $ "SN"
            read
            limpa()
            if upper(p) = "S"
                select 5
                set index to ind01.tic
                seek wticket
                if found()
                    replace pagamento with "S"
                    replace total_ti  with wtotal
                    replace historico with "Pagamento a Vista - Cartao"
                    replace cartao_t  with wcartao
                    replace cartao_n  with wnumero
                    replace tipo_ti   with "C"
                endif
            endif
            restore screen from telawop4
            exit
        other
            loop
    endcase
    exit
enddo
return
