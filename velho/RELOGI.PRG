*****************************************************************************
* Programa...: RELOGI.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......:                   Data da Ultima Atualizacao:                *
* Objetivo...: RELOGIO ON LINE                                              *
* Sistema....: SISVEN - SISTEMA VENDAS E CONTROLE DE CAIXA IF               *
*****************************************************************************
save screen to relarel
esc = inkey()
set color to w/r
dt = quadro(9,9,15,66)
do while esc # 27
    dt= logo1(time(),10,10)
    esc = inkey()
enddo
set color to  w/b
restore screen from relarel
return

function  logo1(texto,lin,col)
local l1,l2,l3,l4,l5,g[6,14]
if len(texto) >13
    @ 23,10 say "ATE 11 CARACTERES SAO PERMITIDOS NO TEXTO"
    inkey(3)
    return(.f.)
endif
if lin=nil
    lin=10
endif
if col=nil
    col=1
endif
g[1,01]="1"
g[2,01]="  лл   "
g[3,01]="   л   "
g[4,01]="   л   "
g[5,01]="   л   "
g[6,01]="  ллл  "
g[1,02]="2"
g[2,02]="лллллл "
g[3,02]="     л "
g[4,02]="лллллл "
g[5,02]="л      "
g[6,02]="лллллл "
g[1,03]="3"
g[2,03]="лллллм "
g[3,03]="     л "
g[4,03]="ллллл  "
g[5,03]="     л "
g[6,03]="лллллп "
g[1,04]="4"
g[2,04]="л    л "
g[3,04]="л    л "
g[4,04]="лллллл "
g[5,04]="     л "
g[6,04]="     л "
g[1,05]="5"
g[2,05]="лллллл "
g[3,05]="л      "
g[4,05]="лллллл "
g[5,05]="     л "
g[6,05]="лллллл "
g[1,06]="6"
g[2,06]="лл     "
g[3,06]="л      "
g[4,06]="лллллл "
g[5,06]="л    л "
g[6,06]="лллллл "
g[1,07]="7"
g[2,07]="лллллл "
g[3,07]="     л "
g[4,07]="     л "
g[5,07]="     л "
g[6,07]="     л "
g[1,08]="8"
g[2,08]="лллллл "
g[3,08]="л    л "
g[4,08]="лллллл "
g[5,08]="л    л "
g[6,08]="лллллл "
g[1,09]="9"
g[2,09]="лллллл "
g[3,09]="л    л "
g[4,09]="лллллл "
g[5,09]="     л "
g[6,09]="    лл "
g[1,10]="0"
g[2,10]="лллллл "
g[3,10]="л  л л "
g[4,10]="л л  л "
g[5,10]="лл   л "
g[6,10]="лллллл "
g[1,11]=":"
g[2,11]="       "
g[3,11]="  ВВ   "
g[4,11]="       "
g[5,11]="  ВВ   "
g[6,11]="       "
g[1,12]="."
g[2,12]="       "
g[3,12]="       "
g[4,12]="       "
g[5,12]="  ВВ   "
g[6,12]="       "
g[1,13]=","
g[2,13]="       "
g[3,13]="       "
g[4,13]="       "
g[5,13]="  ВВ   "
g[6,13]="       "
g[1,13]="-"
g[2,13]="       "
g[3,13]="       "
g[4,13]="  ВВВ  "
g[5,13]="       "
g[6,13]="       "

l1:= l2:= l3:= l4:= l5:=""
z=1
do while z<=len(texto)
    letra=upper(subs(texto,z,1))
    if letra=" "
        l1=l1+space(5)
        l2=l2+space(5)
        l3=l3+space(5)
        l4=l4+space(5)
        l5=l5+space(5)
    else
        indice=ascan(g[1],letra)
        if indice<>0
            l1+=g[2,indice]
            l2+=g[3,indice]
            l3+=g[4,indice]
            l4+=g[5,indice]
            l5+=g[6,indice]
        endif
    endif
    z=z+1
enddo
@ lin++,col say l1
@ lin++,col say l2
@ lin++,col say l3
@ lin++,col say l4
@ lin++,col say l5
return nil
