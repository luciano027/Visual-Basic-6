*****************************************************************************
* Programa...: CALEND.PRG                                                   *
* Programador: LUCIANO MOREIRA ARAUJO                                       *
* Data.......: 02/12/98          Data da Ultima Atualizacao:                *
* Objetivo...: CALENDARIO ON LINE                                           *
* Sistema....:                                                              *
*****************************************************************************
*******************
*function Calendario
*****************************************************************************
*                                                                           *
*  Author    : T_A_P                                                        *
*  Date      : 04/25/88                                                     *
*  Notes     : (C) Copyright 1987-88.  All rights reserved.                 *
*  Comment   : This procedure will display a calendar and allow user to     *
*              change month and year by pressing the arrow key.             *
*              The reason I am not using UDF because I want this program    *
*              to work on any version of Clipper.                           *
*              If you have any comment, please let me know by send it to:   *
*                                                                           *
*              T_A_P                                                        *
*              12336 Inletridge  Apt. H                                     *
*              St. Louis, Mo. 63043                                         *
*  Modificado: Alexandre Soares Campos   Agora, a passagem dos Meses e      *
*              Anos e' feita telas teclas Home, End, PgUp e PgDn, e as      *
*              teclas de setas MOVIMENTAM o calendarioa pelo visor, de      *
*              modo similar `a calculadora. Foram incluido os parametros    *
*              Cale_Lin e Cale_Col, que sao as coordenadas do calendario    *
*                                                                           *
*****************************************************************************
*
parameters cale_lin, cale_col, m_date
publ st_col, no_days
priv telaant,x
if type("Cale_Lin") != 'N'
    cale_lin = 0
endif
if type("Cale_Col") != 'N'
    cale_col = 0
endif
m_date = date()
save screen to telaant
m_day=day(m_date)
m_month=month(m_date)
m_year=year(m_date)
*
do disp_cal                                                   && Display calendar
do disp_hdate with m_date                                     && Display heading date
do find_col with ctod('01/'+str(m_month,2)+'/'+str(m_year,4)) && Find starting column
* Atencao: caso o set date nao esteja setado para british,
* o comando acima dever'a ser:
* do find_col with ctod(+str(m_month,2)+'/01/'+str(m_year,4))
do disp_days with date()                                      && Display days
do while .t.
    key=inkey(0)
    if key=27
        rest screen from telaant
        set color to w/b
        return .t.
    endif
    do case
        case key=3         && PgDn
            m_year=m_year-1
        case key=18        && PgUp
            m_year=m_year+1
        case key=6         && End
            m_month=m_month-1
            if m_month<1
                m_month=12
                m_year=m_year-1
            endif
        case key=1         && Home
            m_month=m_month+1
            if m_month>12
                m_month=1
                m_year=m_year+1
            endif
        case key=4         && ->
            rest screen from telaant
            cale_col = cale_col + 5
            cale_col = if(cale_col>50,50,cale_col)
        case key=19        && <-
            rest screen from telaant
            cale_col = cale_col - 5
            cale_col = if(cale_col<-1,-1,cale_col)
        case key=5         && ^
            rest screen from telaant
            cale_lin = cale_lin - 1
            cale_lin = if(cale_lin<-1,-1,cale_lin)
        case key=24        && v
            rest screen from telaant
            cale_lin = cale_lin + 1
            cale_lin = if(cale_lin>6,6,cale_lin)
    endcase

    if m_day>=28 .and. m_day<=31
        do last_day with m_month, m_year
        m_day=no_days
    endif
    m_date=ctod(str(m_day,2)+'/'+str(m_month,2)+'/'+str(m_year,4))


    if key = 4 .or. key=19 .or. key=5 .or. key=24
        do disp_cal
    endif
    do find_col with ctod('01/'+str(m_month,2)+'/'+str(m_year,4))



    do disp_hdate with m_date
    do clear_date with cale_col+2
    do disp_days with m_date
    loop
enddo
set color to  /bg+
* set color to n/w
*
*
*
proc disp_cal
*
set color to w/n
set color to +bg/n
@ cale_lin+1,cale_col+1,cale_lin+19,cale_col+29 box "ÚÄ¿³ÙÄÀ³ "
@ cale_lin+3,cale_col+1   say "Ã"
@ cale_lin+3,cale_col+29  say "´"
@ cale_lin+5,cale_col+1   say "Ã"
@ cale_lin+5,cale_col+29  say "´"
@ cale_lin+7,cale_col+1   say "Ã"
@ cale_lin+7,cale_col+29  say "´"
@ cale_lin+9,cale_col+1   say "Ã"
@ cale_lin+9,cale_col+29  say "´"
@ cale_lin+11,cale_col+1  say "Ã"
@ cale_lin+11,cale_col+29 say "´"
@ cale_lin+13,cale_col+1  say "Ã"
@ cale_lin+13,cale_col+29 say "´"
@ cale_lin+15,cale_col+1  say "Ã"
@ cale_lin+15,cale_col+29 say "´"
@ cale_lin+17,cale_col+1  say "Ã"
@ cale_lin+17,cale_col+29 say "´"
@ cale_lin+3,cale_col+2   say "ÄÄÄÂÄÄÄÂÄÄÄÂÄÄÄÂÄÄÄÂÄÄÄÂÄÄÄ"
@ cale_lin+4,cale_col+5   say "³   ³   ³   ³   ³   ³"
@ cale_lin+5,cale_col+2   say "ÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄ"
@ cale_lin+6,cale_col+5   say "³   ³   ³   ³   ³   ³"
@ cale_lin+7,cale_col+2   say "ÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄ"
@ cale_lin+8,cale_col+5   say "³   ³   ³   ³   ³   ³"
@ cale_lin+9,cale_col+2   say "ÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄ"
@ cale_lin+10,cale_col+5  say "³   ³   ³   ³   ³   ³"
@ cale_lin+11,cale_col+2  say "ÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄ"
@ cale_lin+12,cale_col+5  say "³   ³   ³   ³   ³   ³"
@ cale_lin+13,cale_col+2  say "ÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄ"
@ cale_lin+14,cale_col+5  say "³   ³   ³   ³   ³   ³"
@ cale_lin+15,cale_col+2  say "ÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄÅÄÄÄ"
@ cale_lin+16,cale_col+5  say "³   ³   ³   ³   ³   ³"
@ cale_lin+17,cale_col+2  say "ÄÄÄÁÄÄÄÁÄÄÄÁÄÄÄÁÄÄÄÁÄÄÄÁÄÄÄ"
set color to +r/b
@ cale_lin+2,cale_col+2   say ""+space(27)+""
set color to +w/b
@ cale_lin+4,cale_col+2  say "Dom"
@ cale_lin+4,cale_col+6  say "Seg"
@ cale_lin+4,cale_col+10 say "Ter"
@ cale_lin+4,cale_col+14 say "Qua"
@ cale_lin+4,cale_col+18 say "Qui"
@ cale_lin+4,cale_col+22 say "Sex"
@ cale_lin+4,cale_col+26 say "Sab"
set color to +gr/b
@ cale_lin+18,cale_col+2 say "Hom End MES - PgUp PgDn ANO"
return
*
*
proc disp_hdate
para m_date
*
* Called from: TAPCAL
* Date       : 04/23/1987
* Comment    : This procedure will display the new heading date for the calendar.
*
set colo to rg+/b
@ cale_lin+2,cale_col+3 say space(25)
@ cale_lin+2,cale_col+5 say substr("Janeiro  FevereiroMar‡o    Abril    "+;
"Maio     Junho    Julho    Agosto   Setembro Outubro  "+;
"Novembro Dezembro",month(m_date)*9-8,9)
@ cale_lin+2,col()+2 say day(m_date)
@ cale_lin+2,col()+2 say year(m_date)
return
*
*
proc disp_days
para m_date
*
* Called from: TAPCAL
* Date       : 04/20/88
* Revised    : 06/15/88
* Comment    : This procedure will display the date on the calendar.
*
st_row=cale_lin+6
do last_day with month(m_date), year(m_date)
for x = 1 to no_days
    set colo to w

    if x=day(m_date)
        set colo to n/w
    endif
    @ st_row,st_col say str(x,2)
    st_col=st_col+4
    st_col=if(st_col>cale_col+27,cale_col+2,st_col)
    st_row=if(st_col=cale_col+2,st_row+2,st_row)
next
return
*
*
proc clear_date
para st_col
*
* Called from : TAPCAL
* Date        : 04/20/87
*
st_row=cale_lin+6
no_days=42
set colo to w
for x = 1 to no_days
    @st_row,st_col say '  '
    st_col=st_col+4
    st_col=if(st_col>cale_col+27,cale_col+2,st_col)
    st_row=if(st_col=cale_col+2,st_row+2,st_row)
next
return
*
*
proc find_col
para m_date
*
* Called from: TAPCAL
* Date       : 4/20/1987
* Comment    : This procedure will find the starting column's depend which day is the first
*              day of the month.
*
fst_day=dow(m_date)
c=cale_col+2
for x=1 to 7
    if fst_day=x
        st_col=c
        exit
    endif
    c=c+4
next
return
*
*
proc last_day
para m_month,m_year
*
*  Called from: DISP_DAYS(TAPCAL)
*  Date       : 04/20/88
*  Revised    : 06/15/88
*  Comment    : This procedure will find the maximum no. of days of a
*               particular month.
*
if m_month=2
    no_days=if(int(m_year/4)=m_year/4,29,28)
    return
endif
no_days=if(str(m_month,2)$' 1 3 5 7 8 10 12',31,30)
return

*****************************************************************************
*                          F   I   M                                        *
*****************************************************************************
