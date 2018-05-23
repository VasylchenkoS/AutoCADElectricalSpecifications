Imports System.Linq
Imports AutoCADElectricalSpecifications.com.vasilchenko.Classes
Imports AutoCADElectricalSpecifications.com.vasilchenko.DatabaseConnection
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Namespace com.vasilchenko.Modules
    Module ItemListMaker
        Public Function MakeSheetItemList(acDocument As Document, acDatabase As Database, acTransaction As Transaction) As SortedList(Of Short, List(Of DrawingSpecificationItem))
            Dim specItemList As New List(Of DrawingSpecificationItem)
            specItemList = DBDataAccessObject.FillSheetItemData(acDatabase, specItemList)
            specItemList = DBDataAccessObject.FillSheetTerminalData(acDatabase, specItemList)
            Dim a = MakeListForSheet(specItemList)
            Return MakeListForSheet(specItemList)
        End Function

        Public Function MakeListForSheet(inputList As List(Of DrawingSpecificationItem)) As SortedList(Of Short, List(Of DrawingSpecificationItem))
            Dim resultList As New SortedList(Of Short, List(Of DrawingSpecificationItem))

            For Each locations In inputList.GroupBy(Function(x) x.Location)
                Dim key As Short = locations(0).Instance \ 10

                Dim locList = locations.ToList
                If key = 8 Then
                    For Each item In locList
                        item.Location = "Кабельно-провідникова продукція"
                    Next
                ElseIf key = 5 Then
                    For Each item In locList
                        item.Location = "Обладнання на ПС"
                    Next
                End If

                If resultList.ContainsKey(key) Then
                    Dim curList = resultList.Item(key)
                    For Each element In locations
                        If curList.Exists(Function(x) x.CatalogName.Equals(element.CatalogName)) Then
                            With curList.Find(Function(x) x.CatalogName.Equals(element.CatalogName))
                                .AddTag(element.TAG)
                                .Count += element.Count - 1
                            End With
                        Else
                            curList.Add(element)
                        End If
                    Next
                Else
                    resultList.Add(key, locList)
                End If
            Next

            Return resultList
        End Function
    End Module
End Namespace

