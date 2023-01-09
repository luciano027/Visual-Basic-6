/***

Programador: Robson Passos Barbosa
E-Mail     : robsonpassos@base.com.br
Home Page  : www.geocities.com/SiliconValley/Sector/3489

#xCommand    SET SCRNSAVE TIME To <n>  =>  SSaveSys(2,<n>)
#xCommand    SET SCRNSAVE TYPE To <n>  =>  SSaveSys(3,<n>)
#xCommand    SET SCRNSAVE KEY  To <n>  =>  SSaveSys(4,<n>)
#xCommand    SET SCRNSAVE ON           =>  SSaveSys(5)
#xCommand    SET SCRNSAVE OFF          =>  SSaveSys(6)
   */

//---------------------------------------------------------------------------
function AtivaScrSave()
//---------------------------------------------------------------------------
SSaveSys( 2, 5 )
SSaveSys( 3, 1 )
SSaveSys( 4, 1 )
SSaveSys( 5 )

//---------------------------------------------------------------------------
function TimeScrSave( nTime )
//---------------------------------------------------------------------------
SSaveSys( 6 )
SSaveSys( 2, nTime )
SSaveSys( 5 )

//---------------------------------------------------------------------------
function TypeScrSave( nType )
//---------------------------------------------------------------------------
nType := IIF( nType == NIL, 1, IIF( nType > 6, 6, nType ) )
SSaveSys( 3, nType )

//---------------------------------------------------------------------------
function DeAtivaScrSave()
//---------------------------------------------------------------------------
SSaveSys( 5 )

//---------------------------------------------------------------------------
function Invert()
//---------------------------------------------------------------------------
Return( "" )
