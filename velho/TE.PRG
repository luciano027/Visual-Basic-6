select 1
use estoque

do while !eof()
    wcodigo_cli = codigo_cli
    replace codigo_cli with strzero(val(wcodigo_cli),5,0)
    skip
enddo
select 2
use saida
do while !eof()
    wcodigo_cli = codigo_cli
    replace codigo_cli with strzero(val(wcodigo_cli),5,0)
    skip
enddo
select 3
use prazo
do while !eof()
    wcodigo_cli = codigo_cli
    replace codigo_cli with strzero(val(wcodigo_cli),5,0)
    skip
enddo


