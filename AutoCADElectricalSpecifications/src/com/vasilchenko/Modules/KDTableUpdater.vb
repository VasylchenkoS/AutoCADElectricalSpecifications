﻿Imports AutoCADElectricalSpecifications.com.vasilchenko.Classes
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput

Namespace com.vasilchenko.Modules
    Public Module KDTableUpdater
        Public Function UpdateSheetTable(acDocument As Document, acDatabase As Database, acTransaction As Transaction, acEditor As Editor) As ObjectId

            Dim itemSortedList As SortedList(Of String, List(Of DrawingSpecificationItem)) = KDItemListMaker.MakeKDItemList(acDatabase, acTransaction)
            Dim acPromptEntOpt As New PromptEntityOptions(vbLf & "Select Table:")
            acPromptEntOpt.SetRejectMessage("You must select a Table." & vbCrLf)
            acPromptEntOpt.AddAllowedClass(GetType(Table), False)
            Dim acResult As PromptEntityResult = acEditor.GetEntity(acPromptEntOpt)

            If acResult.Status <> PromptStatus.OK Then
                Return Nothing
            End If

            Dim acTable = DirectCast(acTransaction.GetObject(acResult.ObjectId, OpenMode.ForWrite), Table)

            If IsNothing(acTable) Then
                Throw New ArgumentNullException($"Таблица не найдена")
            ElseIf acTable.Columns.Count <> 7 Then
                Throw New ArgumentException("Таблица выбрана неверно")
            End If
            acTable.DeleteRows(1, acTable.Rows.Count - 1)
            KDTableDrawing.FillTable(itemSortedList, acTable)

            AdditionalFunctions.FootnoteUpdater(acDatabase, acTransaction, acEditor, acTable)

            Return acResult.ObjectId
        End Function
    End Module
End Namespace
