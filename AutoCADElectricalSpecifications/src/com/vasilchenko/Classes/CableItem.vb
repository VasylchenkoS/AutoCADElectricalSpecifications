Imports System.Linq
Imports AutoCADElectricalSpecifications.com.vasilchenko.Modules

Namespace com.vasilchenko.Classes
    Public Class CableItemClass
        Private _objWires As New List(Of String)
        Private _strFullDescription As String

        Public ReadOnly Property Wires As String
            Get
                if _objWires.Count = 0 Then
                    Return "-"
                else
                    _objWires = _objWires.OrderBy(Function(x) AdditionalFunctions.GetLastNumericFromString(x)).ToList
                    Dim strTag As String = _objWires(0)
                    For s = 1 To _objWires.Count - 1
                        strTag += ", " & _objWires(s)
                    Next
                    Return strTag
                end if
            End Get
        End Property

        Public Sub AddWire(value As String)
            If Not _objWires.Exists(Function(x) x.Equals(value)) Then
                _objWires.Add(value)
            End If
        End Sub

        Public Property TAG As String

        Public Property Length As Integer

        Public Property Source As String

        Public Property Destination As String
        
        Public Property TableDescription As String = ""

        Public Property CatalogName As String

        Public Property Location As String

        Public Property Instance As Short

        Public Property Description As String

        Public Property Manufacture As String

        Public Property Note As String = ""
        
        Public Property Type As String

        Public Property Section As String

        Public Property Cores As String

        Public Property Article As String = ""

        Public Property Unit As String

        Public ReadOnly Property FullDescription As String
            Get
                'Me._strFullDescription = Me.Description & ", " & vbCr & Me.Manufacture & " " & s & IIf(Me.Article <> "", " (" & Me.Article & ")", "")
                _strFullDescription = Description & IIf(CatalogName.ToLower.Contains("nomark"), "", ", " & CatalogName) &
                                         vbCr &
                                         IIf(Article.Equals(""), "", " (" & Article & ")")
                Return _strFullDescription
            End Get
        End Property
        
    End Class

End Namespace
