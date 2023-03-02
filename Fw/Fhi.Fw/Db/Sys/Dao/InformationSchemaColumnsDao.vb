Imports Fhi.Fw.Db.Sys.Vo

Namespace Db.Sys.Dao
    ''' <summary>
    ''' INFORMATION_SCHEMA.COLUMNSビューの簡単なCRUDを集めたDAO
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface InformationSchemaColumnsDao : Inherits DaoEachTable(Of InformationSchemaColumnsVo)
    End Interface
End Namespace