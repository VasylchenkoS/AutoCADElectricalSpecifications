Imports System.Text.RegularExpressions
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput

Imports Autodesk.AutoCAD.ApplicationServices

Namespace com.vasilchenko.Modules
    Module AdditionalFunctions
        Public Function GetLastNumericFromString(s As String) As Double
            If IsNothing(s) Then Return -1
            Dim result As Double = -1
            Dim rgx As New Regex("-?\d*\.?\d+", RegexOptions.IgnoreCase)
            Dim matches As MatchCollection = rgx.Matches(s)
            If matches.Count > 0 Then
                s = matches(matches.Count - 1).Value
                If IsNumeric(s) Then
                    result = Math.Abs(Double.Parse(s))
                Else
                    s = Replace(s, ".", ",")
                    result = Math.Abs(Double.Parse(s))
                End If
                Return result
            Else
                Return -1
            End If
        End Function

        Public Function GetFirstNumericFromString(s As String) As Double
            If IsNothing(s) Then Return -1
            Dim rgx As New Regex("-?\d*\.?\d+", RegexOptions.IgnoreCase)
            Dim matches As MatchCollection = rgx.Matches(s)
            If matches.Count > 0 Then
                s = Replace(matches(0).Value, ".", ",")
                Return Math.Abs(Double.Parse(s))
            Else
                Return -1
            End If
        End Function


        Public Sub FootnoteUpdater(acDatabase As Database, acTransaction As Transaction, acEditor As Editor, acTable As Table)

            Dim acFilterList(0) As TypedValue
            acFilterList.SetValue(New TypedValue(DxfCode.Start, "INSERT"), 0)

            Dim acSelFilter As New SelectionFilter(acFilterList)
            Dim acPrmtpSelResult As PromptSelectionResult = acEditor.SelectAll(acSelFilter)

            If acPrmtpSelResult.Status <> PromptStatus.OK Then
                Return
            End If

            Dim acSelSet As SelectionSet = acPrmtpSelResult.Value
            Dim idArray As ObjectId() = acSelSet.GetObjectIds()

            For Each acBlkId As ObjectId In idArray
                Dim acBlkRef As BlockReference = DirectCast(acTransaction.GetObject(acBlkId, OpenMode.ForWrite), BlockReference)
                Dim acBlkTblRcrd As BlockTableRecord
                If acBlkRef.IsDynamicBlock Then
                    acBlkTblRcrd = DirectCast(acTransaction.GetObject(acBlkRef.DynamicBlockTableRecord, OpenMode.ForWrite), BlockTableRecord)
                Else
                    acBlkTblRcrd = DirectCast(acTransaction.GetObject(acBlkRef.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)
                End If

                If acBlkTblRcrd.Name = "Footnote" Then
                    Dim acAttrCol As AttributeCollection = acBlkRef.AttributeCollection
                    Dim strRowNum As String = ""
                    For Each acAttrId As ObjectId In acAttrCol
                        Dim acAttrRef As AttributeReference = DirectCast(acTransaction.GetObject(acAttrId, OpenMode.ForRead), AttributeReference)
                        If acAttrRef.Tag = "DATA" Then
                            For i = 0 To acTable.Rows.Count
                                If acTable.Cells(i, 2).TextString.Contains(acAttrRef.TextString) Then strRowNum = acTable.Cells(i, 0).TextString
                            Next
                        End If
                    Next
                    For Each acAttrId As ObjectId In acAttrCol
                        Dim acAttrRef As AttributeReference = DirectCast(acTransaction.GetObject(acAttrId, OpenMode.ForWrite), AttributeReference)
                        If acAttrRef.Tag = "TABLENUM" Then acAttrRef.TextString = "Поз. #" & strRowNum
                    Next

                End If
            Next

            acEditor.Regen()

        End Sub

    End Module
End Namespace
