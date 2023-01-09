*****************************************************************************
* Programa...: PRAZOEXC.PRG                                                 *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 07/01/97          Data da Ultima Atualizacao:                *
* Objetivo...: ALTERA/EXCUI ITENS VENDAS REVENDAS                           *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************

qop = 1
limpa()
do while .t.
    set color to n/bg, /w+,,,n
    @ 13,22      prompt " RETORNO "
    @ 13,col()+2 prompt " ALTERAR "
    @ 13,col()+2 prompt " EXCLUIR "
    menu to qop
    do case
        case qop = 1
            limpa()
            return
        case qop = 2
            do while .t.
                ycodigo_est = wcodigo_est
                yquantidade = wquantidade
                set color to w,n/bg,,,n/w
                @ 10,31 get wcodigo_est pict "999999"
                read
                select 1
                set index to ind01.est
                seek wcodigo_est
                if !found()
                    do mensagem with "Codigo nao Cadastrado..."
                    loop
                endif
                iguala()
                set color to gr+/bg
                @ 10,41 say wdescricao pict "!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
                if alltrim(wdescricao,1,5) = "CREDI"
                    do mensagem with "Credito nao pode ser alterado ..."
                    loop
                endif
                wprec = wpreco_apr
                @ 12,38 say wprec       pict "@e 99,999.99"
                set color to w,n/bg,,,n/w
                @ 11,34 get wquantidade pict "@e 999.99"
                read
                if yquantidade < wquantidade
                    wsaldo = (wsaldo+yquantidade)-wquantidade
                    if wsaldo < 0
                        do mensagem with "Saldo menor que quantidade solicitada...",8
                        loop
                    endif
                endif
                exit
            enddo
            p = "S"
            limpa()
            @ 23,03 say "Confirma Dados <S/N> ?" get p pict "@!" valid p $ "SN"
            read
            limpa()
            if upper(p) = "S"
                select 1
                set index to ind01.est,ind02.est
                seek wcodigo_est
                iguala()
                wvalor = wpreco_ven
                if wc_saldo = "S"
                    wsaldo  = (wsaldo+yquantidade) - wquantidade
                    if rec_lock()
                        replace saldo with wsaldo
                        unlock
                    endif
                endif
                select 2
                set index to ind01.cli,ind02.cli
                seek wcodigo_cli
                if rec_lock()
                    replace debito   with wtotal
                    unlock
                endif
                select 3
                set index to ind01.pra
                go reg
                if rec_lock()
                    revar()
                    unlock
                endif
                wnf       = "CLIEN"+wcodigo_cli
                wpreco    = wprec
                wtipo_ven = "P"
                ydata     = wdata
                select 4
                index on codigo_cli to indcon.sai for codigo_cli = ycodigo_cli
                set index to indcon.sai,ind01.sai
                do whil !eof()
                    if wcodigo_est = ycodigo_est .and. wdata = ydata
                        if rec_lock()
                            replace quantidade  with wquantidade
                            unlock
                        endif
                    endif
                    skip
                    iguala()
                enddo
            endif
            return
        case qop = 3
            ydata       = wdata
            ycodigo_est = wcodigo_est
            ydescricao  = wdescricao
            yquantidade = wquantidade
            ycodigo_est = wcodigo_est
            p = "S"
            limpa()
            @ 23,03 say "Confirma Delecao <S/N> ?" get p pict "@!" valid p $ "SN"
            read
            limpa()
            if upper(p) = "S"
                select 3
                set index to ind01.pra
                go reg
                iguala()
                if rec_lock()
                    delete
                endif
                select 1
                set index to ind01.est,ind02.est
                seek wcodigo_est
                iguala()
                if wc_saldo = "S"
                    wsaldo  = wsaldo  + wquantidade
                    gtotal  = wpreco_apr * wquantidade
                    if rec_lock()
                        replace saldo with wsaldo
                        unlock
                    endif
                endif
                wnf = "CLIEN"+wcodigo_cli
                select 4
                index on codigo_cli to indcon.sai for codigo_cli = ycodigo_cli
                set index to indcon.sai,ind01.sai
                do whil !eof()
                    if wcodigo_est = ycodigo_est .and. wdata = ydata
                        if rec_lock()
                            delete
                            unlock
                        endif
                    endif
                    skip
                    iguala()
                enddo
                if alltrim(ydescricao,1,5) = "CREDI"
                    select 2
                    set index to ind01.cli,ind02.cli
                    seek ycodigo_cli
                    if rec_lock()
                        iguala()
                        replace credito  with wcredito-yquantidade
                        unlock
                    endif
                endif
                limpam()
                return
            endif
        other
            return
    endcase
enddo
return
        *****************************************************************************
        *                          F   I   M                                        *
        *****************************************************************************
