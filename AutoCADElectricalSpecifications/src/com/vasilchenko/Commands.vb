Imports AutoCADElectricalSpecifications.com.vasilchenko.Modules
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime

Namespace com.vasilchenko.Main
    Public Class Commands
        <CommandMethod("SpecStart", CommandFlags.Session)>
        Public Shared Sub Main()

            Dim acDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acBatabase As Database = acDocument.Database
            Dim acEditor As Editor = acDocument.Editor

            Application.AcadApplication.ActiveDocument.SendCommand("(command ""_-Purge"")(command ""_ALL"")(command ""*"")(command ""_N"")" & vbCr)
            Application.AcadApplication.ActiveDocument.SendCommand("AEREBUILDDB" & vbCr)

            Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                Dim acTransaction As Transaction = acBatabase.TransactionManager.StartTransaction()
                Try
                    Dim uf As New ufSpecSelector
                    uf.ShowDialog()
                    If uf.rbProjUpdate.Checked = True Then
                    ElseIf uf.rbProjCreate.Checked = True Then
                    ElseIf uf.rbSheetCreate.Checked = True Then
                        SheetTableDrawing.SheetTable(acDocument, acBatabase, acTransaction, acEditor)
                    ElseIf uf.rbSheetUpdate.Checked = True Then
                        SheetTableUpdater.UpdateSheet(acDocument, acBatabase, acTransaction, acEditor)
                    Else Exit Sub
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
