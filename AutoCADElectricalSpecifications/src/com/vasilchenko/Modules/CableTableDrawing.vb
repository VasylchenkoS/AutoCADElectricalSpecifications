Imports System.Linq
Imports AutoCADElectricalSpecifications.com.vasilchenko.Classes
Imports AutoCADElectricalSpecifications.com.vasilchenko.DatabaseConnection
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.GraphicsInterface

Namespace com.vasilchenko.Modules
    Public Module CableTableDrawing
        ReadOnly DblRowsHeight As Double = 8
        Public Sub CableTable(acDocument As Document, acDatabase As Database, acTransaction As Transaction, acEditor As Editor)
            Dim itemList As SortedList(Of Short, List(Of CableItemClass)) = DatabaseDataAccessObject.FillCableItemData()
            
            Dim acPoint As PromptPointResult = acEditor.GetPoint("\nEnter table insertion point: ")
            If acPoint.Status <> PromptStatus.OK Then
                Exit Sub
            End If

            Dim acBlockTable As BlockTable = DirectCast(acTransaction.GetObject(acDatabase.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim acObjectId As ObjectId = acBlockTable(BlockTableRecord.ModelSpace)
            Dim acBlkTblRecord As BlockTableRecord = DirectCast(acTransaction.GetObject(acObjectId, OpenMode.ForWrite), BlockTableRecord)

            Const intColumnsNum As Short = 8

            TextStyleEdit(acDatabase, acTransaction)

            Dim acTable As New Table
            With acTable
                .LayerId = CType(acTransaction.GetObject(acDatabase.LayerTableId, OpenMode.ForRead, False), LayerTable)("0")
                .TableStyle = acDatabase.Tablestyle

#Disable Warning BC40000 ' Тип или член устарел
                .HorizontalCellMargin = 1.0
                .VerticalCellMargin = 0.0
#Enable Warning BC40000 ' Тип или член устарел

                .Position = acPoint.Value
                .InsertRows(0, 16, 1)
                .Rows(0).Alignment = CellAlignment.MiddleCenter
                .InsertColumns(0, 1, intColumnsNum - 1)
                .Columns(0).Width = 25
                .Columns(1).Width = 70
                .Columns(2).Width = 70
                .Columns(3).Width = 30
                .Columns(4).Width = 35
                .Columns(5).Width = 20
                .Columns(6).Width = 85
                .Columns(7).Width = 60
                .Cells(0, 0).TextString = "Маркування"
                .Cells(0, 1).TextString = "Звідки"
                .Cells(0, 2).TextString = "Куди"
                .Cells(0, 3).TextString = "Тип"
                .Cells(0, 4).TextString = "К-ть та переріз жил"
                .Cells(0, 5).TextString = "Довжина, м"
                .Cells(0, 6).TextString = "Марки кіл, які проходять в кабелі"
                .Cells(0, 7).TextString = "Примітка"
                acTable.DeleteRows(1, 1)

                FillTable(itemList, acTable)

                acTable.DeleteRows(1, 1)
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

            acBlkTblRecord.AppendEntity(acTable)
            acTransaction.AddNewlyCreatedDBObject(acTable, True)

            acTable.Dispose()
        End Sub

        Private Sub FillTable(itemSortedList As SortedList(Of Short, List(Of CableItemClass)), ByRef acTable As Table)
            Dim intRowsNum As Integer = 1
            acTable.InsertRows(intRowsNum, DblRowsHeight, 1)
            intRowsNum += 1
            For Each key In itemSortedList.Keys
                Dim itemList = itemSortedList.Item(key)
                InsertLocations(itemList, acTable, intRowsNum)
            Next
        End Sub

        Private Sub InsertLocations(itemList As List(Of CableItemClass), ByRef acTable As Table, ByRef intRowsNum As Integer)
            Dim orderNum As Short = 1
            For Each instance In itemList.GroupBy(Function(x) x.Instance).OrderBy(Function(x) x.Key)
                For Each element In instance.OrderBy(Function(x) x.Tag).ThenBy(Function(x) x.Manufacture).ThenBy(Function(x) x.CatalogName)
                    InsertElement(acTable, intRowsNum, orderNum, element)
                Next
                acTable.InsertRows(intRowsNum, DblRowsHeight, 1)
                intRowsNum += 1
            Next
        End Sub

        Private Sub InsertElement(acTable As Table, ByRef intRowsNum As Integer, ByRef orderNum As Short, element As CableItemClass)
            acTable.InsertRows(intRowsNum, DblRowsHeight, 1)
            With acTable
                .Cells(intRowsNum, 0).TextString = element.Tag
                .Cells(intRowsNum, 1).TextString = element.Source
                .Cells(intRowsNum, 2).TextString = element.Destination
                .Cells(intRowsNum, 3).TextString = element.Type
                .Cells(intRowsNum, 4).TextString = element.Cores & "x" & element.Section
                .Cells(intRowsNum, 5).TextString = element.Length
                .Cells(intRowsNum, 6).TextString = element.Wires
                .Cells(intRowsNum, 7).TextString = element.Note
                .Cells(intRowsNum, 0).Alignment = CellAlignment.MiddleLeft
                .Cells(intRowsNum, 1).Alignment = CellAlignment.MiddleLeft
                .Cells(intRowsNum, 2).Alignment = CellAlignment.MiddleLeft
                .Cells(intRowsNum, 3).Alignment = CellAlignment.MiddleLeft
                .Cells(intRowsNum, 4).Alignment = CellAlignment.MiddleLeft
                .Cells(intRowsNum, 5).Alignment = CellAlignment.MiddleCenter
                .Cells(intRowsNum, 6).Alignment = CellAlignment.MiddleLeft
                .Cells(intRowsNum, 7).Alignment = CellAlignment.MiddleLeft
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

