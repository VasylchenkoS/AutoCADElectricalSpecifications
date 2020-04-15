Imports System.Linq
Imports AutoCADElectricalSpecifications.com.vasilchenko.Modules

Namespace com.vasilchenko.Classes
    Public Class DrawingSpecificationItem
        Private _objTag As New List(Of String)
        Private _strFullDescription As String
        Private _countList As List(Of Integer)

        Sub New()
            _countList = New List(Of Integer)
            _objTag = New List(Of String)
        End Sub
        Public ReadOnly Property Tag As String
            Get
                If _objTag.Count = 0 Then
                    Return "-"
                Else
                    _objTag = _objTag.OrderBy(Function(x) AdditionalFunctions.GetLastNumericFromString(x)).ToList
                    Dim strTag As String = _objTag(0)
                    For s = 1 To _objTag.Count - 1
                        strTag += ", " & _objTag(s)
                    Next
                    Return strTag
                End If
            End Get
        End Property
        Public Sub AddTag(value As String)
            Dim str() As String = Split(value, ",")
            For Each txt In str
                txt = txt.Trim
                If Not _objTag.Exists(Function(x) x.Equals(txt)) Then
                    _objTag.Add(txt)
                End If
            Next
        End Sub
        Public Property Family As String

        Public Property CatalogName As String

        Public Property Location As String

        Public Property Instance As Short

        Public Property Description As String

        Public Property Manufacture As String

        Public Property Count As Double

        Public Property Note As String = ""

        Public Property Article As String = ""

        Public Property Unit As String

        Public ReadOnly Property CountList As List(Of Integer)
            Get
                _countList.Sort()
                _countList.Reverse()
                Return _countList
            End Get
        End Property
        Public Sub AddLength(value As Integer)
            _countList.Add(value)
        End Sub

        Public Property BlockName As String

        Public ReadOnly Property FullDescription As String
            Get
                'Me._strFullDescription = Me.Description & ", " & vbCr & Me.Manufacture & " " & s & IIf(Me.Article <> "", " (" & Me.Article & ")", "")
                _strFullDescription = Description & IIf(CatalogName.ToLower.Contains("nomark"), "", ", " & CatalogName) &
                                         vbCr &
                                         IIf(Article.Equals(""), "", " (" & Article & ")")
                Return _strFullDescription
            End Get
        End Property

        Protected Overrides Sub Finalize()
            _countList = Nothing
            _objTag = Nothing
        End Sub

    End Class

End Namespace
