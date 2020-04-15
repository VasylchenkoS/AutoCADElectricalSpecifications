Imports AutoCADElectricalSpecifications.com.vasilchenko.Forms
Imports AutoCADElectricalSpecifications.com.vasilchenko.Modules
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime

Namespace com.vasilchenko
    Public Class Commands
        <CommandMethod("DEBUG_ASU_Specification", CommandFlags.Session)>
        Public Shared Sub SpecificationKd()

            Dim acDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDatabase As Database = acDocument.Database
            Dim acEditor As Editor = acDocument.Editor

            'Application.AcadApplication.ActiveDocument.SendCommand("(command ""_-Purge"")(command ""_ALL"")(command ""*"")(command ""_N"")" & vbCr)
            Application.AcadApplication.ActiveDocument.SendCommand("AEREBUILDDB" & vbCr)

            Dim acTableObjectID As ObjectId = Nothing

            Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                Using acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction()
                    Try
                        Dim uf As New SpecSelector
                        uf.ShowDialog()
                        If uf.rbProjUpdate.Checked = True Then
                        ElseIf uf.rbProjCreate.Checked = True Then
                            acTableObjectID = ProjectTableDrawing.ProjectTable(acDatabase, acTransaction, acEditor)
                            acEditor.WriteMessage("Таблица успешно создана")
                        ElseIf uf.rbSheetCreate.Checked = True Then
                            acTableObjectID = KDTableDrawing.DrawSheetTable(acDatabase, acTransaction, acEditor)
                            acEditor.WriteMessage("Таблица успешно создана")
                        ElseIf uf.rbSheetUpdate.Checked = True Then
                            acTableObjectID = KDTableUpdater.UpdateSheetTable(acDocument, acDatabase, acTransaction, acEditor)
                            acEditor.WriteMessage("Таблица успешно обновлена")
                        ElseIf uf.rbPageCreate.Checked = True Then
                            acTableObjectID = PageTableDrawing.DrawPagesTable(acDatabase, acTransaction, acEditor)
                            acEditor.WriteMessage("Таблица успешно создана")
                        ElseIf uf.rbPageUpdate.Checked = True Then
                            acTableObjectID = PageTableUpdater.UpdatePageTable(acDatabase, acTransaction, acEditor)
                            acEditor.WriteMessage("Таблица успешно создана")
                        Else
                            Exit Sub
                        End If
                        uf.Dispose()
                        acTransaction.Commit()
                    Catch ex As Exception
                        MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                        acTransaction.Abort()
                    End Try

                End Using
                If Not IsNothing(acTableObjectID) Then TableRowUpdater(acDatabase, acEditor, acTableObjectID)
            End Using
        End Sub

        Public Shared Sub TableRowUpdater(acDatabase As Database, acEditor As Editor, acTableObjectID As ObjectId)

            Using acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction()
                Dim acTable = DirectCast(acTransaction.GetObject(acTableObjectID, OpenMode.ForWrite), Table)
                For i = 1 To acTable.Rows.Count - 1
                    If acTable.Rows(i).Height Mod 8 <> 0 Then
                        For y As Integer = CInt(acTable.Rows(i).Height) To 1000 Step 1
                            If y Mod 8 = 0 Then
                                acTable.Rows(i).Height = y
                                Exit For
                            End If
                        Next
                    End If
                Next
                acTransaction.Commit()
            End Using
        End Sub

    End Class
End Namespace
