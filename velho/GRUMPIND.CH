/*
   File:   GRUMPIND.CH
   System: GRUMPFISH LIBRARY
   Author: Greg Lief
   Copyright (c) 1993 Greg Lief
   User-defined commands for use with IndexBar() and IndexBar2()
   NOTE: If you wish to make use of specialized CA-Clipper 5.2
         indexing (e.g., order bags, conditional, etcetera), you
         should put the following statement somewhere in your
         program prior to #including this header file:

           #define CLIPPER52
*/

/*
   INDEXBAR() directives
*/

#ifdef CLIPPER52

#command INDEX ON <key> [TAG <(cOrderName)> ] TO <(file)>  ;
         GRAPH [ROW <row>] [MESSAGE <msg>]                 ;
         [FOR <for>]                                       ;
         [<all:ALL>]                                       ;
         [WHILE <while>]                                   ;
         [NEXT <next>]                                     ;
         [RECORD <rec>]                                    ;
         [<rest:REST>]                                     ;
         [EVAL <eval>]                                     ;
         [EVERY <every>]                                   ;
         [<unique: UNIQUE>]                                ;
         [<ascend: ASCENDING>]                             ;
         [<descend: DESCENDING>]                           ;
                                     =>                    ;
         IndexBar(<(file)>, <"key">, <row>, .f., <.unique.>, ;
                  <msg>, <"for">, <{for}>, <.all.>, <{while}>, ;
                  <next>, <rec>, <every>, <.rest.>, <.descend.>, ;
                  <(cOrderName)> )
#endif

#xcommand INDEX ON <key> TO <(file)> GRAPH                 ;
          [<u: UNIQUE>] [ROW <row>] [MESSAGE <msg>]        ;
                                                        => ;
          IndexBar(<(file)>, <"key">, <row>, .f., <.u.>, <msg>)


/*
   INDEXBAR2() directives
*/

#xcommand INITIALIZE INDEXES => indexlist := {}

#ifdef CLIPPER52

#command INDEX ON <key> [TAG <(cOrderName)> ] TO <(file)>    ;
         ALIAS <alias>                                       ;
         [FOR <for>]                                         ;
         [<all:ALL>]                                         ;
         [WHILE <while>]                                     ;
         [NEXT <next>]                                       ;
         [RECORD <rec>]                                      ;
         [<rest:REST>]                                       ;
         [EVAL <eval>]                                       ;
         [EVERY <every>]                                     ;
         [<unique: UNIQUE>]                                  ;
         [<ascend: ASCENDING>]                               ;
         [<descend: DESCENDING>]                             ;
                                     =>                      ;
         aadd(indexlist,                                     ;
                 { <"alias">, <(file)>, <"key">, <.unique.>, ;
                   <"for">, <{for}>, <.all.>, <{while}>,     ;
                   <next>, <rec>, <.rest.>, <.descend.>,     ;
                   <(cOrderName)>                            ;
                 } )
#endif

#xcommand INDEX ON <key> TO <(file)> ALIAS <alias> [<u: UNIQUE>] => ;
          aadd(indexlist, { <"alias">, <(file)>, <"key">, <.u.> } )

#xcommand PROCESS INDEXES => IndexBar2(indexlist)


* end of file GRUMPIND.CH
