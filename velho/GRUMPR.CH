/*
    File:    GRUMPR.CH
    System:  GRUMPFISH LIBRARY
    Author:  Greg Lief
    Dialect: Clipper 5.01
    Copyright (c) 1988-91, Greg Lief
    Header file to be used in conjunction with GrumpLib @..GGET and GREAD

    IMPORTANT: If you are using Clipper 5.2, you should put the
               following statement somewhere in your program prior to
               #including this header file:

               #define CLIPPER52

               This contains both Clipper 5.2 and Clipper 5.x versions.
               The differences are:

               (a) Clipper 5.2 requires that you explicitly display the
               GET object with the display() method, whereas Clipper
               5.01 does not; and

               (b) The Clipper 5.2 ReadModal() allows you to pass a
               second parameter indicating which GET should initially
               be made active.  I have therefore added an optional
               INITIAL ## clause that can be used in conjunction with
               the READ command.  Why this is not part of the stock
               language definition (STD.CH) is completely beyond me.

*/

#ifdef CLIPPER52

//ÄÄÄÄÄ ARRAY (and optional NOEDIT) clause
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [ARRAY <array>]                                 ;
                        [ATITLE <ctitle>]                               ;
                        [ATOP <top>]                                    ;
                        [ALEFT <left>]                                  ;
                        [ABOTTOM <bottom>]                              ;
                        [ARIGHT <right>]                                ;
                        [<noedit:NOEDIT>]                               ;
                        [MESSAGE <message>]                             ;
                        [COLOR <color>]                                 ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, ):display()     ;
             )                                                     ;    ;
       Atail(GetList):preblock :=                                       ;
       { | g, x | x := apick(<top>, <left>, <bottom>, <right>,          ;
              <array>,,,,,,,,max(ascan(<array>, trim(g:varGet())), 1),,,;
              <ctitle>),                                                ;
              if(x > 0, g:varPut(padr(<array>\[x], len(g:varGet()))), ),;
              ! <.noedit.> }                                         ;  ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | getapplykey(get, key) } ) }                       ;
       [ ; Atail(GetList):colorDisp( <color> )]                         ;
       [ ; Atail(GetList):cargo := { <message> } ]


//ÄÄÄÄÄ PROPER clause
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [PROPER]                                        ;
                        [MESSAGE <message>]                             ;
                        [COLOR <color>]                                 ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
            GetList,                                                    ;
            _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}>):display();
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | gkeyproper(get, key) } ) }                        ;
       [ ; Atail(GetList):colorDisp( <color> )]                         ;
       [ ; Atail(GetList):cargo := { <message> } ]


//ÄÄÄÄÄ LIST clause
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [LIST <list>]                                   ;
                        [WHEN <when>]                                   ;
                        [MESSAGE <message>]                             ;
                        [COLOR <color>]                                 ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}>):display();
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | gkeylist(get, key) } ) }                          ;
       [ ; Atail(GetList):colorDisp( <color> )]                         ;
       [ ; Atail(GetList):cargo := { <message>, <list> } ]


//ÄÄÄÄÄ MATH clause
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [MATH]                                          ;
                        [MESSAGE <message>]                             ;
                        [COLOR <color>]                                 ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}>):display();
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | getapplykey(get, key) } ) }                  ;    ;
       ATail(Getlist):picture := "@Q"                                   ;
       [ ; Atail(GetList):colorDisp( <color> )]                         ;
       [ ; Atail(GetList):cargo := { <message> } ]


//ÄÄÄÄÄ STEP clause
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [STEP [INCREMENT <step>]]                       ;
                        [<enh:ENHANCED>]                                ;
                        [MESSAGE <message>]                             ;
                        [COLOR <color>]                                 ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}>):display();
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | gkeystep(get, key, <step>, <.enh.>) } ) }         ;
       [ ; Atail(GetList):colorDisp( <color> )]                         ;
       [ ; Atail(GetList):cargo := { <message> } ]


//ÄÄÄÄÄ PASSWORD clause, using default character ("*")
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [PASSWORD ]                                     ;
                        [MESSAGE <message>]                             ;
                        [COLOR <color>]                                 ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}>):display();
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | gkeypass(get, key) } ) }                     ;    ;
       ATail(Getlist):picture := "@P*"                                  ;
       [ ; Atail(GetList):colorDisp( <color> )]                         ;
       [ ; Atail(GetList):cargo := { <message> } ]


//ÄÄÄÄÄ PASSWORD clause with user-defined override character
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [PASSWORD <char>]                               ;
                        [MESSAGE <message>]                             ;
                        [COLOR <color>]                                 ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}>):display();
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | gkeypass(get, key) } ) }                     ;    ;
       ATail(Getlist):picture := "@P" + <char>                          ;
       [ ; Atail(GetList):colorDisp( <color> )]                         ;
       [ ; Atail(GetList):cargo := { <message> } ]


//ÄÄÄÄÄ CALCULATOR clause (with message)
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [<calc:CALCULATOR>]                             ;
                        [<curs:CURSOR,SHOWCURSOR>]                      ;
                        [COLOR <color>]                                 ;
                        [MESSAGE <message>]                             ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}>):display();
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | gkeycalc(get, key) }, <.curs.> ) }                ;
       [ ; Atail(GetList):cargo := { <message>, #<calc> } ]             ;
       [ ; Atail(GetList):colorDisp( <color> )]


//ÄÄÄÄÄ CALCULATOR clause (without message)
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [<calc:CALCULATOR>]                             ;
                        [COLOR <color>]                                 ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}>):display();
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | gkeycalc(get, key) } ) }                          ;
       [ ; Atail(GetList):cargo := { , #<calc> } ]                      ;
       [ ; Atail(GetList):colorDisp( <color> )]


//ÄÄÄÄÄ to allow SAY and GET in same statement
#xcommand @ <row>, <col> SAY <sayxpr>                                   ;
                        [<sayClauses,...>]                              ;
                        GGET <var>                                      ;
                        [<getClauses,...>]                              ;
                                                                        ;
      => @ <row>, <col> SAY <sayxpr> [<sayClauses>]                     ;
       ; @ Row(), Col()+1 GGET <var> [<getClauses>]

//ÄÄÄÄÄ MESSAGE clause only
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [COLOR <color>]                                 ;
                        [MESSAGE <message>]                             ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}>):display();
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | getapplykey(get, key) } ) }                       ;
       [ ; Atail(GetList):colorDisp( <color> )]                         ;
       [ ; Atail(GetList):cargo := { <message> } ]

#xcommand READ                       ;
          TIMEOUT <seconds>          ;
          [INITIAL <init>]           ;
          [EXITEVENT <exitevent>]    ;
          [WARNING <warnsecs>]       ;
          [WARNEVENT <warnevent>]    ;
          [<validation:VALIDATION>]  ;
          [<esc:ALLOW_ESCAPE>]       ;
                                  => ;
          gftimeout(1, <seconds> )                          ;
          [ ; gftimeout(2, <{exitevent}>) ]                 ;
          [ ; gftimeout(3, <warnsecs>) ]                    ;
          [ ; gftimeout(4, <{warnevent}>) ]                 ;
          ; ReadModal( getlist [, <init> ] )                ;
          [; gfvalid(getlist, <.esc.>, <.validation.>)   ]  ;
          ; getlist := {}                                   ;
          ; gftimeout()

#xcommand READ VALIDATION [INITIAL <init>]           ;
          [<esc:ALLOW_ESCAPE>]                       ;
                                                  => ;
          ReadModal( getlist [, <init> ] )           ;
          ; gfvalid(getlist, <.esc.>)                ;
          ; getlist := {}

#xcommand READ INITIAL <init>                        ;
                                                  => ;
          ReadModal( getlist, <init> )               ;
          ; getlist := {}

#xcommand READ SAVE INITIAL <init>                   ;
                                                  => ;
          ReadModal( getlist, <init> )

#xcommand READ INITIAL <init> SAVE                   ;
                                                  => ;
          ReadModal( getlist, <init> )

#else   // Clipper 5.0x syntax follows...

//ÄÄÄÄÄ ARRAY (and optional NOEDIT) clause
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [ARRAY <array>]                                 ;
                        [ATITLE <ctitle>]                               ;
                        [ATOP <top>]                                    ;
                        [ALEFT <left>]                                  ;
                        [ABOTTOM <bottom>]                              ;
                        [ARIGHT <right>]                                ;
                        [<noedit:NOEDIT>]                               ;
                        [MESSAGE <message>]                             ;
                        [COLOR <color>]                                 ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, )               ;
             )                                                     ;    ;
       Atail(GetList):preblock :=                                       ;
       { | g, x | x := apick(<top>, <left>, <bottom>, <right>,          ;
              <array>,,,,,,,,max(ascan(<array>, trim(g:varGet())), 1),,,;
              <ctitle>),                                                ;
              if(x > 0, g:varPut(padr(<array>\[x], len(g:varGet()))), ),;
              ! <.noedit.> }                                         ;  ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | getapplykey(get, key) } ) }                       ;
       [ ; Atail(GetList):colorDisp( <color> )]                         ;
       [ ; Atail(GetList):cargo := { <message> } ]


//ÄÄÄÄÄ PROPER clause
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [PROPER]                                        ;
                        [MESSAGE <message>]                             ;
                        [COLOR <color>]                                 ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}>)       ;
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | gkeyproper(get, key) } ) }                        ;
       [ ; Atail(GetList):colorDisp( <color> )]                         ;
       [ ; Atail(GetList):cargo := { <message> } ]


//ÄÄÄÄÄ LIST clause
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [LIST <list>]                                   ;
                        [WHEN <when>]                                   ;
                        [MESSAGE <message>]                             ;
                        [COLOR <color>]                                 ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}>)       ;
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | gkeylist(get, key) } ) }                          ;
       [ ; Atail(GetList):colorDisp( <color> )]                         ;
       [ ; Atail(GetList):cargo := { <message>, <list> } ]


//ÄÄÄÄÄ MATH clause
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [MATH]                                          ;
                        [MESSAGE <message>]                             ;
                        [COLOR <color>]                                 ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}>)       ;
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | getapplykey(get, key) } ) }                  ;    ;
       ATail(Getlist):picture := "@Q"                                   ;
       [ ; Atail(GetList):colorDisp( <color> )]                         ;
       [ ; Atail(GetList):cargo := { <message> } ]


//ÄÄÄÄÄ STEP clause
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [STEP [INCREMENT <step>]]                       ;
                        [<enh:ENHANCED>]                                ;
                        [MESSAGE <message>]                             ;
                        [COLOR <color>]                                 ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}>)       ;
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | gkeystep(get, key, <step>, <.enh.>) } ) }         ;
       [ ; Atail(GetList):colorDisp( <color> )]                         ;
       [ ; Atail(GetList):cargo := { <message> } ]


//ÄÄÄÄÄ PASSWORD clause, using default character ("*")
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [PASSWORD ]                                     ;
                        [MESSAGE <message>]                             ;
                        [COLOR <color>]                                 ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}>)       ;
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | gkeypass(get, key) } ) }                     ;    ;
       ATail(Getlist):picture := "@P*"                                  ;
       [ ; Atail(GetList):colorDisp( <color> )]                         ;
       [ ; Atail(GetList):cargo := { <message> } ]


//ÄÄÄÄÄ PASSWORD clause with user-defined override character
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [PASSWORD <char>]                               ;
                        [MESSAGE <message>]                             ;
                        [COLOR <color>]                                 ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}>)       ;
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | gkeypass(get, key) } ) }                     ;    ;
       ATail(Getlist):picture := "@P" + <char>                          ;
       [ ; Atail(GetList):colorDisp( <color> )]                         ;
       [ ; Atail(GetList):cargo := { <message> } ]


//ÄÄÄÄÄ CALCULATOR clause (with message)
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [<calc:CALCULATOR>]                             ;
                        [<curs:CURSOR,SHOWCURSOR>]                      ;
                        [COLOR <color>]                                 ;
                        [MESSAGE <message>]                             ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}> )      ;
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | gkeycalc(get, key) }, <.curs.> ) }                ;
       [ ; Atail(GetList):cargo := { <message>, #<calc> } ]             ;
       [ ; Atail(GetList):colorDisp( <color> )]


//ÄÄÄÄÄ CALCULATOR clause (without message)
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [<calc:CALCULATOR>]                             ;
                        [COLOR <color>]                                 ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}> )      ;
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | gkeycalc(get, key) } ) }                          ;
       [ ; Atail(GetList):cargo := { , #<calc> } ]                      ;
       [ ; Atail(GetList):colorDisp( <color> )]


//ÄÄÄÄÄ to allow SAY and GET in same statement
#xcommand @ <row>, <col> SAY <sayxpr>                                   ;
                        [<sayClauses,...>]                              ;
                        GGET <var>                                      ;
                        [<getClauses,...>]                              ;
                                                                        ;
      => @ <row>, <col> SAY <sayxpr> [<sayClauses>]                     ;
       ; @ Row(), Col()+1 GGET <var> [<getClauses>]

//ÄÄÄÄÄ MESSAGE clause only
#xcommand @ <row>, <col> GGET <var>                                     ;
                        [PICTURE <pic>]                                 ;
                        [VALID <valid>]                                 ;
                        [WHEN <when>]                                   ;
                        [COLOR <color>]                                 ;
                        [MESSAGE <message>]                             ;
                                                                        ;
      => SetPos( <row>, <col> )                                         ;
       ; AAdd(                                                          ;
               GetList,                                                 ;
               _GET_( <var>, <(var)>, <pic>, <{valid}>, <{when}> )      ;
             )                                                     ;    ;
       ATail(GetList):reader := { | g | GFReader(g,                     ;
       { | get, key | getapplykey(get, key) } ) }                       ;
       [ ; Atail(GetList):colorDisp( <color> )]                         ;
       [ ; Atail(GetList):cargo := { <message> } ]

#xcommand READ                       ;
          TIMEOUT <seconds>          ;
          [EXITEVENT <exitevent>]    ;
          [WARNING <warnsecs>]       ;
          [WARNEVENT <warnevent>]    ;
          [<validation:VALIDATION>]  ;
          [<esc:ALLOW_ESCAPE>]       ;
                                  => ;
          gftimeout(1, <seconds> )                          ;
          [ ; gftimeout(2, <{exitevent}>) ]                 ;
          [ ; gftimeout(3, <warnsecs>) ]                    ;
          [ ; gftimeout(4, <{warnevent}>) ]                 ;
          ; ReadModal( getlist )                            ;
          [; gfvalid(getlist, <.esc.>, <.validation.>)   ]  ;
          ; getlist := {}                                   ;
          ; gftimeout()

#xcommand READ VALIDATION [<esc:ALLOW_ESCAPE>]       ;
                                                  => ;
          ReadModal( getlist )                       ;
          ; gfvalid(getlist, <.esc.>)                ;
          ; getlist := {}


#endif // CLIPPER52


//ÄÄÄÄÄ @..GET..RANGE (preprocessed to @..GET..VALID)

#command @ <row>, <col> GGET <var>                                      ;
                        [<clauses,...>]                                 ;
                        RANGE <lo>, <hi>                                ;
                        [<moreClauses,...>]                             ;
                                                                        ;
      => @ <row>, <col> GGET <var>                                      ;
                        [<clauses>]                                     ;
                        VALID {|_1| RangeCheck(_1,, <lo>, <hi>)}        ;
                        [<moreClauses>]

* eof GRUMPR.CH
