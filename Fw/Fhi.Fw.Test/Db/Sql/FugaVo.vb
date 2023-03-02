Imports Fhi.Fw.Domain

Namespace Db.Sql
    Friend Class FugaVo
        Public Class StrCollectionObject : Inherits CollectionObject(Of String)

            Public Sub New()
            End Sub

            Public Sub New(ByVal src As CollectionObject(Of String))
                MyBase.New(src)
            End Sub

            Public Sub New(ByVal initialList As IEnumerable(Of String))
                MyBase.New(initialList)
            End Sub
        End Class
        Private _name As String
        Private _intArray As Integer()
        Private _strArray As String()
        Private _strList As List(Of String)
        Private _hogeArray As HogeVo()
        Private _hogeList As List(Of HogeVo)
        Private _fugaArray As FugaVo()
        Public Property StrCollection As StrCollectionObject

        Public Property Name() As String
            Get
                Return _name
            End Get
            Set(ByVal value As String)
                _name = value
            End Set
        End Property

        Public Property IntArray() As Integer()
            Get
                Return _intArray
            End Get
            Set(ByVal value As Integer())
                _intArray = value
            End Set
        End Property

        Public Property StrArray() As String()
            Get
                Return _strArray
            End Get
            Set(ByVal value As String())
                _strArray = value
            End Set
        End Property

        Public Property StrList() As List(Of String)
            Get
                Return _strList
            End Get
            Set(ByVal value As List(Of String))
                _strList = value
            End Set
        End Property

        Public Property HogeArray() As HogeVo()
            Get
                Return _hogeArray
            End Get
            Set(ByVal value As HogeVo())
                _hogeArray = value
            End Set
        End Property

        Public Property HogeList() As List(Of HogeVo)
            Get
                Return _hogeList
            End Get
            Set(ByVal value As List(Of HogeVo))
                _hogeList = value
            End Set
        End Property

        Public Property FugaArray() As FugaVo()
            Get
                Return _fugaArray
            End Get
            Set(ByVal value As FugaVo())
                _fugaArray = value
            End Set
        End Property
    End Class
End Namespace