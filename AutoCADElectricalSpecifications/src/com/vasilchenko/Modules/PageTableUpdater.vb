Imports AutoCADElectricalSpecifications.com.vasilchenko.Classes
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput

Namespace com.vasilchenko.Modules
    Public Module PageTableUpdater
        Public Function UpdatePageTable(acDatabase As Database, acTransaction As Transaction, acEditor As Editor) As ObjectId

            Dim list As List(Of Integer) = PageListMaker.SelectPages()
            Dim itemSortedList As SortedList(Of String, List(Of DrawingSpecificationItem)) = PageListMaker.MakePagesItemList(acDatabase, acTransaction, list)

            Dim acPromptEntOpt As New PromptEntityOptions(vbLf & "Select Table:")
            acPromptEntOpt.SetRejectMessage("You must select a Table." & vbCrLf)
            acPromptEntOpt.AddAllowedClass(GetType(Table), False)
            Dim acResult As PromptEntityResult = acEditor.GetEntity(acPromptEntOpt)

            If acResult.Status <> PromptStatus.OK Then
                Return Nothing
            End If

            Dim acTable = DirectCast(acTransaction.GetObject(acResult.ObjectId, OpenMode.ForWrite), Table)

            If IsNothing(acTable) Then
                Throw New ArgumentNullException("Таблица не найдена")
            ElseIf acTable.Columns.Count <> 6 Then
                Throw New ArgumentException("Таблица выбрана неверно")
            End If

            acTable.DeleteRows(1, acTable.Rows.Count - 1)
            PageTableDrawing.FillTable(itemSortedList, acTable)

            Return acResult.ObjectId
        End Function

    End Module
End Namespace
