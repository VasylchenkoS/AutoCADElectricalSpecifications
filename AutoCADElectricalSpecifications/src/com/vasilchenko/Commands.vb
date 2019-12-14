Imports AutoCADElectricalSpecifications.com.vasilchenko.Forms
Imports AutoCADElectricalSpecifications.com.vasilchenko.Modules
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime

Namespace com.vasilchenko
    Public Class Commands
        <CommandMethod("ASU_Specification", CommandFlags.Session)>
        Public Shared Sub SpecificationKd()

            Dim acDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDatabase As Database = acDocument.Database
            Dim acEditor As Editor = acDocument.Editor

            'Application.AcadApplication.ActiveDocument.SendCommand("(command ""_-Purge"")(command ""_ALL"")(command ""*"")(command ""_N"")" & vbCr)
            Application.AcadApplication.ActiveDocument.SendCommand("AEREBUILDDB" & vbCr)

            Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                Dim acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction()
                Try
                    Dim uf As New SpecSelector
                    uf.ShowDialog()
                    If uf.rbProjUpdate.Checked = True Then
                    ElseIf uf.rbProjCreate.Checked = True Then
                        ProjectTableDrawing.ProjectTable(acDatabase, acTransaction, acEditor)
                        acEditor.WriteMessage("Таблица успешно создана")
                    ElseIf uf.rbSheetCreate.Checked = True Then
                        KDTableDrawing.DrawSheetTable(acDatabase, acTransaction, acEditor)
                        acEditor.WriteMessage("Таблица успешно создана")
                    ElseIf uf.rbSheetUpdate.Checked = True Then
                        KDTableUpdater.UpdateSheetTable(acDocument, acDatabase, acTransaction, acEditor)
                        acEditor.WriteMessage("Таблица успешно обновлена")
                    ElseIf uf.rbPageCreate.Checked = True Then
                        PageTableDrawing.DrawPagesTable(acDatabase, acTransaction, acEditor)
                        acEditor.WriteMessage("Таблица успешно создана")
                    ElseIf uf.rbPageUpdate.Checked = True Then
                        PageTableUpdater.UpdatePageTable(acDatabase, acTransaction, acEditor)
                        acEditor.WriteMessage("Таблица успешно создана")
                    Else
                        Exit Sub
                    End If
                    uf.Dispose()
                    acTransaction.Commit()
                Catch ex As Exception
                    MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                    acTransaction.Abort()
                Finally
                    acTransaction.Dispose()
                End Try
            End Using

        End Sub
    End Class
End Namespace
