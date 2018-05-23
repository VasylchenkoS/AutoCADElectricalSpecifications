Imports System.Linq
Imports AutoCADElectricalSpecifications.com.vasilchenko.Classes
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.GraphicsInterface

Namespace com.vasilchenko.Modules
    Module SheetTableDrawing
        Dim dblRowsHeight As Double = 8
        Public Sub SheetTable(acDocument As Document, acDatabase As Database, acTransaction As Transaction, acEditor As Editor)
            Dim itemList As SortedList(Of Short, List(Of DrawingSpecificationItem)) = ItemListMaker.MakeSheetItemList(acDocument, acDatabase, acTransaction)

            Dim acPoint As PromptPointResult = acEditor.GetPoint("\nEnter table insertion point: ")
            If acPoint.Status <> PromptStatus.OK Then
                Exit Sub
            End If

            Dim acBlockTable As BlockTable = DirectCast(acTransaction.GetObject(acDatabase.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim acObjectId As ObjectId = acBlockTable(BlockTableRecord.ModelSpace)
            Dim acBlkTblRcrd As BlockTableRecord = DirectCast(acTransaction.GetObject(acObjectId, OpenMode.ForWrite), BlockTableRecord)

            Dim intColumnsNum As Integer = 4

            Dim acTextStyleTable As TextStyleTable = acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead)

            TextStyleEdit(acDatabase, acTransaction)

            Dim acTable As New Table
            With acTable
                .LayerId = CType(acTransaction.GetObject(acDatabase.LayerTableId, OpenMode.ForRead, False), LayerTable)("0")
                .TableStyle = acDatabase.Tablestyle

#Disable Warning BC40000 ' Тип или член устарел
                .HorizontalCellMargin = 1.0
#Enable Warning BC40000 ' Тип или член устарел

#Disable Warning BC40000 ' Тип или член устарел
                .VerticalCellMargin = 0.0
#Enable Warning BC40000 ' Тип или член устарел

                .Position = acPoint.Value
                .InsertRows(0, 16, 1)
                .Rows(0).Alignment = CellAlignment.MiddleCenter
                .InsertColumns(0, 1, intColumnsNum)
                .Columns(0).Width = 10
                .Columns(1).Width = 25
                .Columns(2).Width = 105
                .Columns(3).Width = 10
                .Columns(4).Width = 35
                .Cells(0, 0).TextString = "№№"
                .Cells(0, 1).TextString = "Позиція"
                .Cells(0, 2).TextString = "Найменування"
                .Cells(0, 3).TextString = "Кіл."
                .Cells(0, 4).TextString = "Примітка"

                acTable.DeleteRows(1, 1)

                FillTable(itemList, acTable)
            End With

            acTable.RecomputeTableBlock(True)
            acTable.BreakEnabled = True
            acTable.RecomputeTableBlock(True)
            acTable.BreakOptions = TableBreakOptions.EnableBreaking Or TableBreakOptions.RepeatTopLabels Or TableBreakOptions.AllowManualHeights
            acTable.RecomputeTableBlock(True)
            acTable.BreakFlowDirection = TableBreakFlowDirection.Right
            acTable.SetBreakHeight(0, 287)
            acTable.SetBreakHeight(1, 232)
            acTable.SetBreakSpacing(25)
            acTable.RecomputeTableBlock(True)

            acBlkTblRcrd.AppendEntity(acTable)
            acTransaction.AddNewlyCreatedDBObject(acTable, True)

        End Sub

        Friend Sub FillTable(itemSortedList As SortedList(Of Short, List(Of DrawingSpecificationItem)), ByRef acTable As Table)
            Dim intRowsNum As Integer = 1
            Dim locCount As Short = 1
            For Each key In itemSortedList.Keys
                Dim itemList = itemSortedList.Item(key)
                InsertLocations(itemList, locCount, acTable, intRowsNum)
            Next
        End Sub

        Public Sub InsertLocations(itemList As List(Of DrawingSpecificationItem), ByRef locCount As Short, ByRef acTable As Table, ByRef intRowsNum As Integer)
            acTable.InsertRows(intRowsNum, dblRowsHeight, 1)
            acTable.Rows(intRowsNum).Style = "Данные"
            acTable.Rows(intRowsNum).IsMergeAllEnabled = True
            acTable.Cells(intRowsNum, 0).TextString = locCount & ". " & itemList(0).Location
            acTable.Cells(intRowsNum, 0).Alignment = CellAlignment.MiddleCenter
            intRowsNum += 1
            locCount += 1
            Dim orderNum As Short = 1
            For Each instance In itemList.GroupBy(Function(x) x.Instance).OrderBy(Function(x) x.Key)
                For Each element In instance.OrderBy(Function(x) x.TAG)
                    InsertElement(acTable, intRowsNum, orderNum, element)
                Next
            Next
        End Sub

        Private Sub InsertElement(acTable As Table, ByRef intRowsNum As Integer, ByRef orderNum As Short, element As DrawingSpecificationItem)
            Dim elementCount As Short
            acTable.InsertRows(intRowsNum, dblRowsHeight, 1)
            If element.Unit.Equals("к-т.") OrElse element.Unit.Equals("уп.") Then
                elementCount = element.Count / AdditionalFunctions.GetNumericFromString(element.Description)
            Else
                elementCount = element.Count
            End If
            With acTable
                .Cells(intRowsNum, 0).TextString = orderNum
                .Cells(intRowsNum, 1).TextString = element.TAG
                .Cells(intRowsNum, 2).TextString = element.FullDescription
                .Cells(intRowsNum, 3).TextString = elementCount
                .Cells(intRowsNum, 4).TextString = element.Note
                .Cells(intRowsNum, 0).Alignment = CellAlignment.MiddleCenter
                .Cells(intRowsNum, 1).Alignment = CellAlignment.MiddleLeft
                .Cells(intRowsNum, 2).Alignment = CellAlignment.MiddleLeft
                .Cells(intRowsNum, 3).Alignment = CellAlignment.MiddleCenter
                .Cells(intRowsNum, 4).Alignment = CellAlignment.MiddleLeft
                intRowsNum += 1
                orderNum += 1
            End With
        End Sub
        Private Sub TextStyleEdit(acDatabase As Database, acTransaction As Transaction)
            Dim acTextStyleTable As TextStyleTable = CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead, False), TextStyleTable)
            Dim acStyleId As ObjectId = acTextStyleTable("Standard")
            Dim acTextStyleTableRecord As TextStyleTableRecord = DirectCast(acTransaction.GetObject(acStyleId, OpenMode.ForWrite), TextStyleTableRecord)
            With acTextStyleTableRecord
                .TextSize = 2.5
                .Font = New FontDescriptor("ISOCPEUR", False, True, Nothing, Nothing)
            End With
        End Sub
    End Module
End Namespace

