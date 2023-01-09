/*
   File:   APICK.CH
   System: GRUMPFISH LIBRARY
   Author: Greg Lief
   Copyright (c) 1993 Greg Lief
   User-defined command for the Grumpfish Library APick() funciton
*/

#ifndef _APICK_DEFINED

#xcommand APICK [ <var> ]           ; // no need to specify this with tagging
          ARRAY <array>             ;
          [ TOP <top> ]             ;
          [ LEFT <left> ]           ;
          [ BOTTOM <bottom> ]       ;
          [ RIGHT <right> ]         ;
          [ BOXCOLOR <boxcolor> ]   ;
          [ BARCOLOR <barcolor> ]   ;
          [ STATCOLOR <statcolor> ] ;
          [ LENGTH <length> ]       ;
          [ TAGGING <tags> ]        ;
          [ UNSELECTED <unsel>]     ;
          [ HIGHLIGHT <highlight>]  ;
          [ INITIAL <initial>]      ;
          [ SELECTED <selected>]    ;
          [ UNAVAILABLE <unava>]    ;
          [ TITLE <title>]          ;
          [ BOXTYPE <box>]          ;
          [ <nowrap:NOWRAP> ]       ;
          [ <nomove:NOMOVE> ]       ;
          [ <nobar:NOBAR> ]         ;
          [ <norestore:NORESTORE> ] ;
                              =>    ;
         [<var> :=] APick( <top>, <left>, <bottom>, <right>, <array>, ;
                           <boxcolor>, <barcolor>, <statcolor>,       ;
                           <length>, [@<tags>], <unsel>, <highlight>, ;
                           <initial>, <selected>, <unava>, <title>,   ;
                           ! <.nowrap.>, <box>, ! <.nomove.>,         ;
                           ! <.nobar.>, ! <.norestore.> )

#define _APICK_DEFINED

#endif

//ÄÄÄÄÄ end of file APICK.CH
