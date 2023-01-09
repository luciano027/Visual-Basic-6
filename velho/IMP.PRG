 linha = 1
 set device to print
 @ prow(),1 say chr(27)+chr(120)+chr(0)
 @ prow(),1 say chr(15)
 @ prow()+1,84 say date() pict "@d"
 @ prow()+1,01 say ""
 @ prow()+1,01 say ""
 @ prow()+1,01 say ""
 @ prow()+1,03 say date() pict "@d"
         @ prow(),20   say "XXXXXXXXXX"
         @ prow(),53   say "XXX"
         @ prow(),63   say date() pict "@d"
* @ prow()+1,01 say ""
 @ prow()+1,93 say "xxxxxxxxx"
 @ prow()+1,01 say ""
 @ prow()+1,01 say ""
 @ prow()+1,01 say "xxxxxxxxxXxxx10xxxxXxxx20xxxxXxxx30xxxxXxxx40xxxxXxxx50xxxxXxxx60xx"
 @ prow()+1,01 say "xxxxxxxxxXxxx10xxxxXxxx20xxxxXxxx30xxxxXxxx40xxxxXxxx50xxxxXxxx60xx"
 @ prow()+1,01 say "xxxxxxxxxXxxx10xxxxXxxx20xxxxXxxx30xxxxXxxx40xxxxXxxx50xxxxXxxx60xx"
 @ prow()+1,01 say ""
 @ prow()+1,01 say ""
 @ prow()+1,01 say ""
 @ prow()+1,01 say "xxxxxxxxxXxxx10xxxxXxxx20xxxxXxxx30xxxxXxxx40xxxxXxxx50xxxxXxxx60xx"
 @ prow()+1,01 say "xxxxxxxxxXxxx10xxxxXxxx20xxxxXxxx30xxxxXxxx40xxxxXxxx50xxxxXxxx60xx"
 @ prow()+1,01 say ""
 @ prow()+1,2 say chr(18)
                    @ prow()+1,3 say ""
                    @ prow()+1,4 say ""
                    set printer off
                    set device to screen
