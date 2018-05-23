Imports AutoCADElectricalSpecifications.com.vasilchenko.Modules

Namespace com.vasilchenko.Classes
    Public Class DrawingSpecificationItem
        Private _objTAG As New List(Of String)
        Private _strFamily As String
        Private _strCatalogName As String
        Private _strLocation As String
        Private _shtInstance As Short
        Private _strDescription As String
        Private _strManufacture As String
        Private _shtCount As Short
        Private _strFullDescription As String
        Private _strNote As String = ""
        Private _strArticle As String = ""
        Private _strUnit As String
        Public ReadOnly Property TAG As String
            Get
                _objTAG.Sort()
                If _objTAG.Count < 3 Then
                    Dim strTag As String = _objTAG(0)
                    For s = 1 To _objTAG.Count - 1
                        strTag += ", " & _objTAG(s)
                    Next
                    Return strTag
                ElseIf _objTAG.Count > 2 Then
                    Dim strTag As String = _objTAG(0) & "-" & _objTAG(_objTAG.Count - 1)
                    Return strTag
                End If
                Return ""
            End Get
        End Property
        Public Sub AddTag(value As String)
            If Not Me._objTAG.Exists(Function(x) x.Equals(value)) Then
                Me._objTAG.Add(value)
            End If
            Me._shtCount += 1
        End Sub
        Public Property Family As String
            Get
                Return _strFamily
            End Get
            Set(value As String)
                Me._strFamily = value
            End Set
        End Property

        Public Property CatalogName As String
            Get
                Return _strCatalogName
            End Get
            Set(value As String)
                Me._strCatalogName = value
            End Set
        End Property

        Public Property Location As String
            Get
                Return _strLocation
            End Get
            Set(value As String)
                Me._strLocation = value
            End Set
        End Property

        Public Property Instance As Short
            Get
                Return _shtInstance
            End Get
            Set(value As Short)
                Me._shtInstance = value
            End Set
        End Property
        Public Property Description As String
            Get
                Return _strDescription
            End Get
            Set(value As String)
                Me._strDescription = value
            End Set
        End Property

        Public Property Manufacture As String
            Get
                Return _strManufacture
            End Get
            Set(value As String)
                Me._strManufacture = value
            End Set
        End Property

        Public Property Count As Short
            Get
                Return _shtCount
            End Get
            Set(value As Short)
                Me._shtCount = value
            End Set
        End Property

        Public Property Note As String
            Get
                Return _strNote
            End Get
            Set(value As String)
                Me._strNote = value
            End Set
        End Property

        Public Property Article As String
            Get
                Return _strArticle
            End Get
            Set(value As String)
                Me._strArticle = value
            End Set
        End Property

        Public Property Unit As String
            Get
                Return _strUnit
            End Get
            Set(value As String)
                Me._strUnit = value
            End Set
        End Property

        Public ReadOnly Property FullDescription As String
            Get
                Me._strFullDescription = Me.Description & ", " & vbCr & Me.Manufacture & " " & Me.CatalogName & IIf(Me.Article <> "", " (" & Me.Article & ")", "")
                Return _strFullDescription
            End Get
        End Property
    End Class

End Namespace
