/*
   File:   SHADOWB.CH
   System: GRUMPFISH LIBRARY
   Author: Greg Lief
   Copyright (c) 1993 Greg Lief
   User-defined command for the Grumpfish Library ShadowBox() funciton
*/

#ifndef _SHADOWBOX_DEFINED

#xcommand SHADOWBOX               ;
         [ BUFFER <buffer> ]      ;
         [ TOP <top> ]            ;
         [ LEFT <left> ]          ;
         [ BOTTOM <bottom> ]      ;
         [ RIGHT <right> ]        ;
         [ TYPE <type> ]          ;
         [ TITLE <title> ]        ;
         [ FOOTER <footer> ]      ;
         [ CROSSBAR <crossbar> ]  ;
         [ COLOR <color> ]        ;
                              =>  ;
         [<buffer> := ] ShadowBox(<top>, <left>, <bottom>, <right>, ;
                        <type>, <title>, <crossbar>, <color>, <footer>)

#define _SHADOWBOX_DEFINED

#endif

//ÄÄÄÄÄ end of file SHADOWB.CH
