*****************************************************************************
* Programa...: MENU.PRG                                                     *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 05/06/95          Data da Ultima Atualizacao: 05/06/95       *
* Objetivo...: Memu Principal                                               *
* Sistema....: CONTROLE DE LOCACAO VIDEO/CONTAS A PAGAR/FORNECEDOR          *
*****************************************************************************

** ---------------- preparando o anbiente de trabalho  -----------------**
set talk off
set dele on
set echo off
set date brit
set century on
set conf on
set excl off
setcancel(.f.)

** ---------------------------------------------------------------------**
set proc to arqproc
set proc to demo2
set message to 21 center
*** set key 28 to help
set key -1 to popcalc
set key -2 to vendaif
set key -3 to consul
set key -4 to venda
set key -5 to nfsaida
set key -6 to cadgenda

abertura()

clear screen
set color to  w/b
@ 03,00 to 20,79

AtivaScrSave()     // ativa saver
TypeScrSave(3)    // determina tipo de saver
TimeScrSave(360)  // aguarda n segundos

select 1
if !net_use("empresa",.t.,"","")
    return
else
    locate for !empty(empresa)
    if found()
        zempresa  = empresa
        zendereco = endereco
        ztelefone = telefone
        zfax      = fax
        zmicro    = alltrim(micro)
        zdiretorio= alltrim(diretorio)
        zdireto   = alltrim(direto)
        zsalto = salto
    endif
endif
use

set defa to &zdiretorio

public zmicro,zendereco,zempresa

inicio()
* abertura()
if substr(zdiretorio,1,1) = "C" .or. substr(zdiretorio,1,1) = "c"
    save screen to tela013
    waiton("Organizando Arquivos...")
    atualiza()
    testa_arq()
    restore screen from tela013
endif

do while .t.
    set color to n/w
    SHOWTIME(01,58,.F.,"R/w",.F.,.F.)             // CHAMA O RELOGIO NA TELA
    do sinal with "Menu Principal"
    set color to w/b
    @ 02,02       prompt "  Tabelas   "
    @ 02,col()+01 prompt " Movimento  "
    @ 02,col()+01 prompt "   Vendas   "
    @ 02,col()+01 prompt "Utilitarios " &&&53
    @ 02,col()+23 prompt "Fim"
    save screen to telamenu
    menu to opcao
    do case
        case opcao=1
            save screen to telaopcao1
            do sinal with "  Tabelas           "
            limpa()
            do menutab
            close all
            limpa()
            restore screen from telaopcao1
        case opcao = 2
            save screen to telaopcao2
            do sinal with " Movimento          "
            do menumov
            close all
            limpa()
            restore screen from telaopcao2
        case opcao = 3
            save screen to telaopcao2
            do sinal with "Venda a Vista       "
            do menuvend
            libera()
            close all
            limpa()
            restore screen from telaopcao2
        case opcao = 4
            save screen to telaopcao6
            do sinal with "Utilitarios         "
            do menuutil
            close all
            limpa()
            restore screen from telaopcao6
            loop
        case opcao = 5
            save screen to telafim
            if Confirma ("Confirma Saida do Sistema? ")=1    && gravacao dos dados
                set color to
                rolatela(00,01,24,78)
                clear
                close all
                quit
            endif
            restore screen from telafim
        other
            tone(261.7,2)
            tone(392,7)
            loop
    endcase
enddo

** ---------------------------- Consulta de Dados no Estoque ---------------
procedure consul
if used()
    do mensagem with "Pesquisa nao funciona com uma ARQUIVO ABERTO"
    return
endif
close all
select 1
if !net_use("estoque",.f.,"ind01.est","ind02.est")
    return
endif
save screen to tela_0
set curso on
do pes_est
set curso on
restore screen from tela_0
set color to w/b
close all
return

** -----------------------------Saida no Estoque com NF ---------------------
procedure venda
if used()
    do mensagem with "Venda nao funciona com uma ARQUIVO ABERTO"
    return
endif
close all
set curso on
save screen to telao22
do sinal with "Venda a Vista  "
do cadsai1
set curso on
close all
restore screen from telao22
set color to w/b
return

**----------------------- Saida no Estoque sem NF --------------------------
procedure nfsaida
if used()
    do mensagem with "Venda nao funciona com uma ARQUIVO ABERTO"
    return
endif
close all
set curso on
save screen to telao22
do sinal with "Ficha do Cliente    "
do cadprazo
close all
restore screen from telao22
set color to w/b
set curso on
return


** --------------------------- Agenda de Telefone -------------------------
procedure telefone
set color to
popphone()
set color to w/b
return


** -----------------------------Saida no Estoque com NF ---------------------
procedure vendaif
if used()
    do mensagem with "Venda nao funciona com uma ARQUIVO ABERTO"
    return
endif
close all
set curso on
save screen to telao22
do sinal with "Orcamento      "
do cadsai2
set curso on
close all
restore screen from telao22
set color to w/b
return

** -----------------------------Saida no Estoque com NF ---------------------
procedure cadgenda
if used()
    do mensagem with "Venda nao funciona com uma ARQUIVO ABERTO"
    return
endif
close all
select 1
if !net_use("agenda",.f.,"ind01.age","ind02.age")
    return
endif
set curso on
save screen to telao22
do sinal with "Agenda         "
do pes_age
set curso on
close all
restore screen from telao22
set color to w/b
return


                                        **************************************************************************                                                                           *
                                        *                             F       I       M                           *
                                        ***************************************************************************
