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
        Dim oOVPM As SAPbobsCOM.Payments
        Dim llError As Long
        Dim lsError, ObjectCode As String
        Dim otekLiq As FrmtekLIQ
        Dim stQuery As String
        Dim oRecSet As SAPbobsCOM.Recordset
        Dim TransId, Fecha As String

        Try
            oForm = SBOApplication.Forms.Item("tekReconciliation")
            oGrid = oForm.Items.Item("3").Specific
            oDataTable = oGrid.DataTable

            Fecha = ArreglarFechas(Now.Date)

            For i = 0 To oDataTable.Rows.Count - 1

                If oDataTable.GetValue("Liquidar", i) = "Y" Then

                    oOVPM = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
                    oRecSet = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    '///// Encabezado del PagoEfectuado
                    oOVPM.DocType = 0
                    oOVPM.CardCode = oDataTable.GetValue("RFC", i)
                    oOVPM.CashSum = 0
                    oOVPM.TransferAccount = "410101999"
                    oOVPM.TransferSum = oDataTable.GetValue("Saldo_a_Favor", i)
                    oOVPM.TransferDate = DateTime.Now
                    oOVPM.PayToCode = ""

                    'stQuery = "Select T0.""Line"" From (Select A.""TransId"",ROW_NUMBER() OVER (PARTITION BY C.""LicTradNum"" ORDER BY A.""TransId"" ASC) AS ""Line"" from OJDT A Inner Join JDT1 B on A.""TransId"" = B.""TransId"" INNER JOIN OCRD C on C.""CardCode"" = B.""ShortName"" where B.""BalDueCred""<>0 and C.""CardCode""='" & oDataTable.GetValue("RFC", i) & "' and C.""validFor""='Y' Order by C.""LicTradNum"",A.""TransId"") T0 Where T0.""TransId""=" & oDataTable.GetValue("Asiento", i)
                    'oRecSet.DoQuery(stQuery)

                    'If oRecSet.RecordCount > 0 Then

                    'Set.MoveFirst()

                    oOVPM.Invoices.DocEntry = oDataTable.GetValue("Asiento", i)
                    oOVPM.Invoices.DocLine = 1 '2
                    oOVPM.Invoices.InvoiceType = 24
                    oOVPM.Invoices.SumApplied = oDataTable.GetValue("Saldo_a_Favor", i)

                    'End If

                    If oOVPM.Add() <> 0 Then

                        SBOCompany.GetLastError(llError, lsError)
                        Err.Raise(-1, 1, lsError)

                    Else

                        stQuery = "Select T0.""TransId"" from OVPM T0 
                                   INNER JOIN VPM2 T1 ON T0.""DocEntry"" = T1.""DocNum"" 
                                   where T0.""CardCode""='" & oDataTable.GetValue("RFC", i) & "' and T0.""TrsfrAcct""=410101999 and T0.""TrsfrSum""=" & oDataTable.GetValue("Saldo_a_Favor", i) & " and T0.""TrsfrDate""='" & Fecha & "' and T1.""DocEntry""=" & oDataTable.GetValue("Asiento", i)
                        oRecSet.DoQuery(stQuery)

                        If oRecSet.RecordCount > 0 Then

                            oRecSet.MoveFirst()
                            TransId = oRecSet.Fields.Item("TransId").Value

                            UpdateOJDT(TransId)

                        End If

                    End If

                    End If

            Next

            otekLiq = New FrmtekLIQ
            otekLiq.AgregarLineas()

        Catch ex As Exception
            SBOApplication.MessageBox("Liquidar: " & ex.Message)
            Return -1
        End Try

    End Function


    Public Function UpdateOJDT(ByVal TransId As String)

        Dim oOJDT As SAPbobsCOM.JournalEntries
        Dim stQuery As String
        Dim oRecSet As SAPbobsCOM.Recordset
        Dim LineID As String
        Dim llError As Long
        Dim lsError, ObjectCode As String

        Try

            oOJDT = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
            oRecSet = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            '// Encabezado OJDT
            oOJDT.GetByKey(TransId)
            oOJDT.ProjectCode = "001"

            stQuery = "Select T1.""Line_ID"" from OJDT T0 INNER JOIN JDT1 T1 ON T0.""TransId"" = T1.""TransId"" WHERE T0.""TransId""=" & TransId
            oRecSet.DoQuery(stQuery)

            If oRecSet.RecordCount > 0 Then

                oRecSet.MoveFirst()

                For cont2 As Integer = 0 To oRecSet.RecordCount - 1

                    '//Lineas OJDT
                    LineID = oRecSet.Fields.Item("Line_ID").Value
                    oOJDT.Lines.SetCurrentLine(LineID)

                    oOJDT.Lines.ProjectCode = "001"
                    oOJDT.Lines.CostingCode = "001"

                    LineID = Nothing

                    oRecSet.MoveNext()

                Next

            End If

            If oOJDT.Update() <> 0 Then

                SBOCompany.GetLastError(llError, lsError)
                Err.Raise(-1, 1, lsError)

            End If

            TransId = Nothing

        Catch ex As Exception

            MsgBox("Error al Actualizar OJDT: " & ex.Message)

        End Try

    End Function

    Public Function ArreglarFechas(ByVal stFecha As String) As String

        Try
            Dim oRecSetF1 As SAPbobsCOM.Recordset
            Dim stQueryF1 As String
            Dim Fecha As String

            oRecSetF1 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            stQueryF1 = "select substring('" & stFecha & "',7,4)as ""año1"",substring('" & stFecha & "',4,2) as ""mes1"",substring('" & stFecha & "',1,2)as ""dia1"" from dummy;"
            oRecSetF1.DoQuery(stQueryF1)

            If oRecSetF1.RecordCount > 0 Then
                oRecSetF1.MoveFirst()
                Fecha = oRecSetF1.Fields.Item("año1").Value & "-" & oRecSetF1.Fields.Item("mes1").Value & "-" & oRecSetF1.Fields.Item("dia1").Value
            End If

            'MsgBox(Fecha1)
            Return Fecha

        Catch ex As Exception
            SBOApplication.MessageBox("ArreglasFecha. " & ex.Message)
        End Try

    End Function

    'Public Function Reconciliar(ByRef TransID1 As Integer, ByRef TransID2 As Integer, ByRef ReconcileAmount As Decimal, ByVal CardCode As String)

    '    Dim oReconService As SAPbobsCOM.InternalReconciliationsService
    '    Dim oParam As SAPbobsCOM.InternalReconciliationOpenTransParams
    '    Dim oOposting As SAPbobsCOM.InternalReconciliationOpenTrans
    '    Dim stQuery, stQuery2 As String
    '    Dim oRecSet, oRecSet2 As SAPbobsCOM.Recordset
    '    Dim TransRowId1, TransRowId2 As Integer

    '    Try

    '        oReconService = SBOCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.InternalReconciliationsService)
    '        oParam = oReconService.GetDataInterface(SAPbobsCOM.InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTransParams)
    '        oOposting = oReconService.GetDataInterface(SAPbobsCOM.InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)
    '        oRecSet = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRecSet2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        stQuery = "Select T1.""Line_ID"" As ""TransRowId1"" from OJDT T0 INNER JOIN JDT1 T1 On T0.""TransId"" = T1.""TransId"" WHERE T0.""TransId""=" & TransID1 & "  And T1.""ContraAct""='" & CardCode & "'"

    '        stQuery2 = "Select T1.""Line_ID""  as ""TransRowId2"" from OJDT T0 INNER JOIN JDT1 T1 ON T0.""TransId"" = T1.""TransId"" WHERE T0.""TransId""=" & TransID2 & "  and T1.""ShortName""='" & CardCode & "'"

    '        oRecSet.DoQuery(stQuery)
    '        oRecSet2.DoQuery(stQuery2)

    '        If oRecSet.RecordCount > 0 Then

    '            oRecSet.MoveFirst()
    '            TransRowId1 = oRecSet.Fields.Item("TransRowId1").Value

    '        End If

    '        If oRecSet2.RecordCount > 0 Then

    '            oRecSet2.MoveFirst()
    '            TransRowId2 = oRecSet2.Fields.Item("TransRowId2").Value

    '        End If

    '        '' Set date selection criteria
    '        'oParam.ReconDate = DateTime.Today
    '        'oParam.DateType = ReconSelectDateTypeEnum.rsdtPostDate
    '        'oParam.FromDate = New DateTime(2017, 10, 11)
    '        'oParam.ToDate = New DateTime(2017, 11, 11)
    '        ''' Set account Or card selection criterio                   
    '        oParam.CardOrAccount = 0
    '        oParam.InternalReconciliationBPs.Add()
    '        oParam.InternalReconciliationBPs.Item(0).BPCode = CardCode

    '        'oOposting.ReconDate = Now
    '        oOposting.InternalReconciliationOpenTransRows.Add()
    '        oOposting.InternalReconciliationOpenTransRows.Item(0).Selected = SAPbobsCOM.BoYesNoEnum.tYES
    '        oOposting.InternalReconciliationOpenTransRows.Item(0).TransId = TransID1
    '        oOposting.InternalReconciliationOpenTransRows.Item(0).ReconcileAmount = ReconcileAmount
    '        oOposting.InternalReconciliationOpenTransRows.Item(0).TransRowId = TransRowId1
    '        oOposting.InternalReconciliationOpenTransRows.Add()
    '        oOposting.InternalReconciliationOpenTransRows.Item(1).Selected = SAPbobsCOM.BoYesNoEnum.tYES
    '        oOposting.InternalReconciliationOpenTransRows.Item(1).TransId = TransID2
    '        oOposting.InternalReconciliationOpenTransRows.Item(1).ReconcileAmount = ReconcileAmount * -1
    '        oOposting.InternalReconciliationOpenTransRows.Item(1).TransRowId = TransRowId2

    '        oParam = oReconService.Add(oOposting)

    '    Catch ex As Exception

    '        SBOApplication.MessageBox("Error al reconciliar: ", ex.Message)

    '    End Try

    'End Function

End Class
