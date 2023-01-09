/*
   File:   GRUMPBRO.CH
   System: GRUMPFISH LIBRARY
   Author: Greg Lief
   Copyright (c) 1993 Greg Lief
   User-defined command for the Grumpfish Library GrumpBrow() funciton
*/

#ifndef _GRUMPBROW_DEFINED

#xcommand GRUMPBROW                     ;
         [ TO <var> ]                   ;
         [ SECURITY <sec> ]             ;
         [ TOP <top> ]                  ;
         [ LEFT <left> ]                ;
         [ BOTTOM <bottom> ]            ;
         [ RIGHT <right> ]              ;
         [ TITLE <title> ]              ;
         [ FIELDS <fields> ]            ;
         [ HEADINGS <headings> ]        ;
         [ PICTURES <pictures> ]        ;
         [ ALTERNATES <alts> ]          ;
         [ LOW <start> ]                ;
         [ HIGH <end> ]                 ;
         [ VBLOCK <vblock> ]            ;
         [ HBLOCK <hblock> ]            ;
         [ EXTRA  <extra> ]             ;
         [ EXTRAKEY <extrakey> ]        ;
         [ EXTRATITLE <extratitle> ]    ;
         [ EXITKEY <exitkey> ]          ;
         [ INDEX_DESCRIPTION <aDescs> ] ;
         [ VALIDS <aValids> ]           ;
         [ WHENS <aWhens> ]             ;
         [ AUTOREFRESH <nseconds> ]     ;
         [ LOCK <nlock> ]               ;
         [ <carry:CARRY> ]              ;
         [ MEMOWIDTH <memowidth> ]      ;
         [ MEMOHEIGHT <memoheight> ]    ;
         [ COLORSPEC <colorspec>   ]    ;
         [ COLORBLOCKS <cblocks>   ]    ;
         [ INITIAL <initial>  ]         ;
         [ <gotop:GOTOP> ]              ;
         [ ALTERNATE_KEYS <aKeys> ]     ;
                                     => ;
         [ <var> := ] GrumpBrow(<sec>, <top>, <left>, <bottom>, <right>,  ;
                      .t., <fields>, <headings>, <pictures>,              ;
                      <alts>, <start>, <end>, <{vblock}>, <{hblock}>,     ;
                      <extra>, <extrakey>, <extratitle>, <exitkey>,       ;
                      <aDescs>, <aValids>, <aWhens>, <nseconds>, <nlock>, ;
                      <.carry.>, <memowidth>, <title>, <memoheight>,      ;
                      <colorspec>, <cblocks>, <initial>, <.gotop.>, <aKeys> )

/*
  The following directive is for backwards compatibility with version
  3.x.  The NOBOX clause had to be removed in light of the preprocessor
  limitation of 32 optional clauses per directive (and because I felt
  there were other clauses far more important than this one).  If you
  truly desire not to have a box drawn around the Grumpbrow() window,
  pass a logical False as the 6th parameter to the function.
*/
#xcommand GRUMPBROW                     ;
         [<clauses,...> ]               ;
         NOBOX                          ;
         [<moreClauses,...> ]           ;
                                     => ;
         GRUMPBROW [ <clauses> ] [ <moreClauses> ]

#define _GRUMPBROW_DEFINED

#endif

//ÄÄÄÄÄ end of file GRUMPBRO.CH
