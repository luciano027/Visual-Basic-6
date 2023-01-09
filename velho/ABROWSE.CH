/*
   File:   ABROWSE.CH
   System: GRUMPFISH LIBRARY
   Author: Greg Lief
   Copyright (c) 1993 Greg Lief
   User-defined command for the Grumpfish Library ABrowse() funciton
*/

#ifndef _ABROWSE_DEFINED

#xcommand ABROWSE [ <nVar> ]                    ;
          ARRAY <aData>                         ;
          [ TOP <nTop> ]                        ;
          [ LEFT <nLeft> ]                      ;
          [ BOTTOM <nBottom> ]                  ;
          [ RIGHT <nRight> ]                    ;
          [ HEADINGS <aHeads> ]                 ;
          [ INITIAL <nInitial>]                 ;
          [ WIDTHS <aWidths> ]                  ;
          [ TITLE <cTitle> ]                    ;
          [ <celledit:EDIT,CELLEDIT,EDITABLE> ] ;
          [ <maintain:MAINTAIN,MAINTENANCE> ]   ;
                              =>    ;
         [<nVar> :=] ABrowse( <nTop>, <nLeft>, <nBottom>, <nRight>,     ;
                              <aData>, <aHeads>, <nInitial>, <aWidths>, ;
                              <cTitle>, <.celledit.>, <.maintain.> )

#define _ABROWSE_DEFINED

#endif

//ÄÄÄÄÄ end of file ABROWSE.CH
