Imports System.Linq
Imports AutoCADElectricalSpecifications.com.vasilchenko.Classes
Imports AutoCADElectricalSpecifications.com.vasilchenko.DatabaseConnection
Imports Autodesk.AutoCAD.DatabaseServices
Namespace com.vasilchenko.Modules
    Module ItemListMaker
        Public Function MakeSheetItemList(acDatabase As Database, acTransaction As Transaction) As SortedList(Of String, List(Of DrawingSpecificationItem))
            Dim specItemList As New List(Of DrawingSpecificationItem)
            DatabaseDataAccessObject.FillItemData(specItemList)
            DatabaseDataAccessObject.FillMountItemData(specItemList)
            DatabaseDataAccessObject.FillSheetDinAndDuctItemData(specItemList)
            DatabaseDataAccessObject.FillMountTerminalData(specItemList)
            Return MakeListForSheet(specItemList)
        End Function

        Private Function MakeListForSheet(inputList As List(Of DrawingSpecificationItem)) As SortedList(Of String, List(Of DrawingSpecificationItem))
            Dim resultList As New SortedList(Of String, List(Of DrawingSpecificationItem))

            For Each location In inputList.GroupBy(Function(x) x.Location)
                Dim instKey As Short = location(0).Instance
                Dim locKey As string

                Dim locList as List(Of DrawingSpecificationItem) = location.ToList

                If instKey >= 800 AndAlso instKey >= 899 Then
                    locKey = "Кабельно-провідникова продукція"
                ElseIf instKey >= 700 AndAlso instKey >= 799 Then
                    locKey = "Обладнання на ПС"
                Else 
                    locKey = location(0).Location
                End If

                If resultList.ContainsKey(locKey) Then
                    'Dim curList = resultList.Item(instKey)
                    'For Each element In location
                    '    If curList.Exists(Function(x) x.CatalogName.Equals(element.CatalogName)) Then
                    '        With curList.Find(Function(x) x.CatalogName.Equals(element.CatalogName))
                    '            .AddTag(element.TAG)
                    '            '.Count += element.Count - 1
                    '        End With
                    '    Else
                    '        curList.Add(element)
                    '    End If
                    'Next
                    MsgBox("что-то пошло не так")
                Else
                    resultList.Add(locKey, locList)
                End If
            Next

            Return resultList
        End Function
    End Module
End Namespace

