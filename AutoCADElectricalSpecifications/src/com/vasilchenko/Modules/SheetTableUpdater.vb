Imports System.Linq
Imports AutoCADElectricalSpecifications.com.vasilchenko.Classes
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput

Namespace com.vasilchenko.Modules
    Module SheetTableUpdater
        Public Sub UpdateSheet(acDocument As Document, acDatabase As Database, acTransaction As Transaction, acEditor As Editor)

            Dim itemSortedList As SortedList(Of Short, List(Of DrawingSpecificationItem)) = ItemListMaker.MakeSheetItemList(acDocument, acDatabase, acTransaction)
            Dim acPromptEntOpt As New PromptEntityOptions(vbLf & "Select Table:")
            acPromptEntOpt.SetRejectMessage("You must select a Table." & vbCrLf)
            acPromptEntOpt.AddAllowedClass(GetType(Table), False)
            Dim acResult As PromptEntityResult = acEditor.GetEntity(acPromptEntOpt)
            Dim acStatus As PromptStatus = acResult.Status

            If acResult.Status <> PromptStatus.OK Then
                Exit Sub
            End If

            Dim acTable As Table = DirectCast(acTransaction.GetObject(acResult.ObjectId, OpenMode.ForWrite), Table)

            If IsNothing(acTable) Then
                Throw New ArgumentNullException("Таблица не найдена")
            ElseIf acTable.Columns.Count <> 5 Then
                Throw New ArgumentException("Таблица выбрана неверно")
            End If

            Dim locCount As Short = 1
            Dim rowsNum As Short = 1
            For Each key In itemSortedList.Keys
                Dim itemList = itemSortedList.Item(key)
                If acTable.Cells(rowsNum, 0).TextString <> locCount & ". " & itemList(0).Location Then
                    acTable.DeleteRows(1, acTable.Rows.Count - 1)
                    SheetTableDrawing.FillTable(itemSortedList, acTable)
                Else
                    rowsNum += 1
                    locCount += 1
                    For Each instance In itemList.GroupBy(Function(x) x.Instance).OrderBy(Function(x) x.Key)
                        For Each element In instance.OrderBy(Function(x) x.TAG)
                            If acTable.Cells(rowsNum, 1).TextString <> element.TAG Then
                                acTable.DeleteRows(1, acTable.Rows.Count - 1)
                                SheetTableDrawing.FillTable(itemSortedList, acTable)
                            Else
                                rowsNum += 1
                            End If
                        Next
                    Next
                End If
            Next
        End Sub
    End Module
End Namespace
