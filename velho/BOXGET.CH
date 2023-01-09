/*
   File:   BOXGET.CH
   System: GRUMPFISH LIBRARY
   Author: Greg Lief
   Copyright (c) 1993 Greg Lief
   User-defined command for the Grumpfish Library BoxGet() funciton
*/

#ifndef _BOXGET_DEFINED

#xcommand BOXGET <var>             ;
          PROMPT <prom>            ;
          [ PICTURE <pict>  ]      ;
          [ VALID <valid>   ]      ;
          [ BOXCOLOR <boxcolor> ]  ;
          [ COLOR <color>   ]      ;
          [ ROW <row>       ]      ;
          [ COLUMN <column> ]      ;
          [ TITLE <title> ]        ;
          [ <norest:NORESTORE> ]   ;
          [ <restall:RESTOREALL> ] ;
          [ <double:DOUBLE> ]      ;
                               =>  ;
         BoxGet(<prom>, <row>, <column>,                                   ;
            getnew( maxrow() + 1, maxcol() + 1,                            ;
            { | _grumpy | if(pcount() = 0, <var>, <var> := _grumpy ) },    ;
            <(var)>, <pict>, <color> ), <{valid}>, <boxcolor>, <.norest.>, ;
            <.restall.>, <.double.>, <title> )

#define _BOXGET_DEFINED

#endif

//ÄÄÄÄÄ end of file BOXGET.CH
