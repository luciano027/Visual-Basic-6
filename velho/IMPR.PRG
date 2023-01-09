 linha = 1
 set device to print
 @ prow(),1 say chr(27)+chr(120)+chr(0)
 @ prow(),1 say chr(15)
 do while linha < 15
    @ prow()+1,2 say "xxxxxxxxxXxxx10xxxxXxxx20xxxxXxxx30xxxxXxxx40xxxxXxxx50xxxxXxxx60xxxxXxxx70xxxxXxxx80xxxxXxxx90xxxxXxxxx100xxXxxxx110xxX"
    linha = linha + 1
 enddo
 @ prow()+1,2 say chr(18)
                    @ prow()+1,3 say ""
                    @ prow()+1,4 say ""
                    set printer off
                    set device to screen
