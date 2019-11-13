Public Class Liquidar

    Private SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private SBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private oForm As SAPbouiCOM.Form           '//FORMA
    Private csFormUID As String
    Private stDocNum As String
    'Friend Monto As Double


    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        SBOApplication = oCatchingEvents.SBOApplication
        SBOCompany = oCatchingEvents.SBOCompany
        Me.stDocNum = stDocNum
    End Sub


    Public Function Liquidar()

        Dim oGrid As SAPbouiCOM.Grid
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim oOJDT As SAPbobsCOM.JournalEntries
        Dim llError As Long
        Dim lsError, ObjectCode As String
        Dim otekLiq As FrmtekLIQ

        Try
            oForm = SBOApplication.Forms.Item("tekReconciliation")
            oGrid = oForm.Items.Item("3").Specific
            oDataTable = oGrid.DataTable

            For i = 0 To oDataTable.Rows.Count - 1

                If oDataTable.GetValue("Liquidar", i) = "Y" Then

                    oOJDT = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

                    '///// Encabezado del Asiento
                    oOJDT.Series = 19
                    oOJDT.DueDate = DateTime.Now
                    oOJDT.ReferenceDate = DateTime.Now
                    oOJDT.Memo = "Asiento de corrección saldo a favor del cliente"
                    oOJDT.Reference = oDataTable.GetValue("Asiento", i)

                    '///// Renglones del Asiento
                    '//Debitos
                    oOJDT.Lines.Add()
                    oOJDT.Lines.SetCurrentLine(0)
                    oOJDT.Lines.AccountCode = "110601001"
                    oOJDT.Lines.ShortName = oDataTable.GetValue("RFC", i)
                    oOJDT.Lines.FederalTaxID = oDataTable.GetValue("RFC", i)
                    oOJDT.Lines.Debit = oDataTable.GetValue("Saldo_a_Favor", i)
                    oOJDT.Lines.LineMemo = "Debito por saldo a favor RFC: " + oDataTable.GetValue("RFC", i)
                    oOJDT.Lines.TaxDate = DateTime.Now
                    oOJDT.Lines.CostingCode = "001"
                    oOJDT.Lines.ProjectCode = "001"
                    oOJDT.DueDate = DateTime.Now

                    '//Creditos
                    oOJDT.Lines.Add()
                    oOJDT.Lines.SetCurrentLine(1)
                    oOJDT.Lines.AccountCode = "410101999"
                    oOJDT.Lines.Credit = oDataTable.GetValue("Saldo_a_Favor", i)
                    oOJDT.Lines.LineMemo = "Credito por saldo a favor RFC: " + oDataTable.GetValue("RFC", i)
                    oOJDT.Lines.TaxDate = DateTime.Now
                    oOJDT.DueDate = DateTime.Now
                    oOJDT.Lines.CostingCode = "001"
                    oOJDT.Lines.ProjectCode = "001"

                    If oOJDT.Add() <> 0 Then

                        SBOCompany.GetLastError(llError, lsError)
                        Err.Raise(-1, 1, lsError)

                    Else

                        Reconciliar(SBOCompany.GetNewObjectKey, oDataTable.GetValue("Asiento", i), oDataTable.GetValue("Saldo_a_Favor", i))

                    End If

                End If

            Next

            otekLiq = New FrmtekLIQ
            otekLiq.AgregarLineas()

        Catch ex As Exception
            SBOApplication.MessageBox("cerrarOrdenesVentas: " & ex.Message)
            Return -1
        End Try

    End Function


    Public Function Reconciliar(ByRef TransID1 As String, ByRef TransID2 As String, ByRef ReconcileAmount As String)

        Dim oReconService As SAPbobsCOM.InternalReconciliationsService
        Dim oParam As SAPbobsCOM.InternalReconciliationParams
        Dim oOposting As SAPbobsCOM.InternalReconciliationOpenTrans

        Try

            oReconService = SBOCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.InternalReconciliationsService)
            oParam = oReconService.GetDataInterface(SAPbobsCOM.InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
            oOposting = oReconService.GetDataInterface(SAPbobsCOM.InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)

            oOposting.InternalReconciliationOpenTransRows.Add()
            oOposting.InternalReconciliationOpenTransRows.Item(0).Selected = SAPbobsCOM.BoYesNoEnum.tYES
            oOposting.InternalReconciliationOpenTransRows.Item(0).TransId = TransID2
            oOposting.InternalReconciliationOpenTransRows.Item(0).ReconcileAmount = ReconcileAmount
            oOposting.InternalReconciliationOpenTransRows.Item(0).TransRowId = 0
            oOposting.InternalReconciliationOpenTransRows.Add()
            oOposting.InternalReconciliationOpenTransRows.Item(1).Selected = SAPbobsCOM.BoYesNoEnum.tYES
            oOposting.InternalReconciliationOpenTransRows.Item(1).TransId = TransID1
            oOposting.InternalReconciliationOpenTransRows.Item(1).ReconcileAmount = ReconcileAmount
            oOposting.InternalReconciliationOpenTransRows.Item(1).TransRowId = 1

            oParam = oReconService.Add(oOposting)

        Catch ex As Exception

            SBOApplication.MessageBox("Error al reconciliar: ", ex.Message)

        End Try

    End Function


End Class
