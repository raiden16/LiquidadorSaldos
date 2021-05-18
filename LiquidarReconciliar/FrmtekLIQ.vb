Imports System.Drawing

Public Class FrmtekLIQ

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private csFormUID As String
    Private stDocNum As String
    'Friend Monto As Double


    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        Me.stDocNum = stDocNum
    End Sub

    'Private Property stRuta As String 

    '//----- ABRE LA FORMA DENTRO DE LA APLICACION 
    Public Function openForm(ByVal psDirectory As String)
        Dim oRecSetH As SAPbobsCOM.Recordset
        'Dim Monto As Integer

        Try

            oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            csFormUID = "tekReconciliation"
            '//CARGA LA FORMA
            If (loadFormXML(cSBOApplication, csFormUID, psDirectory + "\Forms\" + csFormUID + ".srf") <> 0) Then

                Err.Raise(-1, 1, "")
            End If
            '"que pedo"
            '--- Referencia de Forma
            setForm(csFormUID)

            '---- refresca forma
            coForm.Refresh()
            coForm.Visible = True

        Catch ex As Exception
            If (ex.Message <> "") Then
                cSBOApplication.MessageBox("FrmtekDel. No se pudo iniciar la forma. " & ex.Message)
            End If
            Me.close()
        End Try
    End Function


    '//----- CIERRA LA VENTANA
    Public Function close() As Integer
        close = 0
        coForm.Close()
    End Function


    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Function setForm(ByVal psFormUID As String) As Integer
        Try
            setForm = 0
            '//ESTABLECE LA REFERENCIA A LA FORMA
            coForm = cSBOApplication.Forms.Item(psFormUID)
            '//OBTIENE LA REFERENCIA A LOS USER DATA SOURCES
            setForm = getUserDataSources()
        Catch ex As Exception
            cSBOApplication.MessageBox("FrmtekDel. Al referenciar a la forma. " & ex.Message)
            setForm = -1
        End Try
    End Function


    '//----- OBTIENE LA REFERENCIA A LOS USERDATASOURCES
    Private Function getUserDataSources() As Integer
        'Dim llIndice As Integer
        Try
            coForm.Freeze(True)
            getUserDataSources = 0
            '//SI YA EXISTEN LOS DATASOURCES, SOLO LOS ASOCIA
            If (coForm.DataSources.UserDataSources.Count() > 0) Then
            Else '//EN CASO DE QUE NO EXISTAN, LOS CREA
                getUserDataSources = bindUserDataSources()
            End If
            coForm.Freeze(False)
        Catch ex As Exception
            cSBOApplication.MessageBox("FrmtekDel. Al referenciar los UserDataSources" & ex.Message)
            getUserDataSources = -1
        End Try
    End Function


    '//----- ASOCIA LOS USERDATA A ITEMS
    Private Function bindUserDataSources() As Integer
        Dim loText As SAPbouiCOM.EditText
        Dim loDS As SAPbouiCOM.UserDataSource
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim oGrid As SAPbouiCOM.Grid

        Try
            bindUserDataSources = 0

            loDS = coForm.DataSources.UserDataSources.Add("DateFrom", SAPbouiCOM.BoDataType.dt_DATE) 'Creo el datasources
            loText = coForm.Items.Item("1").Specific  'identifico mi caja de fecha
            loText.DataBind.SetBound(True, "", "DateFrom")   ' uno mi userdatasources a mi caja de fecha

            loDS = coForm.DataSources.UserDataSources.Add("DateTo", SAPbouiCOM.BoDataType.dt_DATE) 'Creo el datasources
            loText = coForm.Items.Item("2").Specific  'identifico mi caja de fecha
            loText.DataBind.SetBound(True, "", "DateTo")   ' uno mi userdatasources a mi caja de fecha

            loDS = coForm.DataSources.UserDataSources.Add("dsAmount", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
            loText = coForm.Items.Item("9").Specific  'identifico mi caja de texto
            loText.DataBind.SetBound(True, "", "dsAmount")   ' uno mi userdatasources a mi caja de texto

            oGrid = coForm.Items.Item("3").Specific
            oDataTable = coForm.DataSources.DataTables.Add("Liquidar")
            oGrid.DataTable = oDataTable

        Catch ex As Exception
            cSBOApplication.MessageBox("FrmtekDel. Al crear los UserDataSources. " & ex.Message)
            bindUserDataSources = -1
        Finally
            loText = Nothing
            loDS = Nothing
            oDataTable = Nothing
            oGrid = Nothing
        End Try
    End Function


    '----- carga los procesos de carga
    Public Function AgregarLineas()
        Dim oGrid As SAPbouiCOM.Grid
        Dim stQuery As String = ""
        Dim oRecSet As SAPbobsCOM.Recordset
        Dim Fdate, Tdate, Amount, Fecha1, Fecha2 As String

        Try
            coForm = cSBOApplication.Forms.Item("tekReconciliation")

            Fdate = coForm.DataSources.UserDataSources.Item("DateFrom").Value
            Tdate = coForm.DataSources.UserDataSources.Item("DateTo").Value
            Amount = coForm.DataSources.UserDataSources.Item("dsAmount").Value

            Fecha1 = ArreglarFechas(Fdate)
            Fecha2 = ArreglarFechas(Tdate)

            oGrid = coForm.Items.Item("3").Specific
            oGrid.DataTable.Clear()

            oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuery = "Select C.""LicTradNum"" as ""RFC"", C.""CardName"" as ""Cliente"",B.""BalDueCred"" as ""Saldo_a_Favor"", D.""PrjName"" as ""Sucursal"", A.""Project"" as ""Cod_Sucursal"", A.""TransId"" as ""Asiento"", '' as ""Liquidar""
                       from OJDT A
                       Inner Join JDT1 B on A.""TransId"" = B.""TransId""
                       INNER JOIN OCRD C on C.""CardCode"" = B.""ShortName""
                       INNER JOIN OPRJ D On A.""Project"" = D.""PrjCode""
                       where B.""BalDueCred""<>0 and C.""validFor""='Y' and C.""CardType""='C' and A.""TransType""=24 and A.""RefDate"" Between '" & Fecha1 & "' and '" & Fecha2 & "' and B.""BalDueCred""<=" & Amount & "
                       Order by C.""LicTradNum"",A.""TransId"""

            oGrid.DataTable.ExecuteQuery(stQuery)

            oGrid.Columns.Item(6).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(4).Editable = False
            oGrid.Columns.Item(5).Editable = False

            Return 0

        Catch ex As Exception

            MsgBox("FrmtekDel. fallo la carga previa de la forma AgregarLineas: " & ex.Message)

        Finally

            oGrid = Nothing

        End Try

    End Function


    Public Function ArreglarFechas(ByVal stFecha As String) As String

        Try
            Dim oRecSetF1 As SAPbobsCOM.Recordset
            Dim stQueryF1 As String
            Dim Fecha As String

            oRecSetF1 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            stQueryF1 = "select substring('" & stFecha & "',7,4)as ""año1"",substring('" & stFecha & "',4,2) as ""mes1"",substring('" & stFecha & "',1,2)as ""dia1"" from dummy;"
            oRecSetF1.DoQuery(stQueryF1)

            If oRecSetF1.RecordCount > 0 Then
                oRecSetF1.MoveFirst()
                Fecha = oRecSetF1.Fields.Item("año1").Value & "-" & oRecSetF1.Fields.Item("mes1").Value & "-" & oRecSetF1.Fields.Item("dia1").Value
            End If

            'MsgBox(Fecha1)
            Return Fecha

        Catch ex As Exception
            cSBOApplication.MessageBox("ArreglasFecha1. " & ex.Message)
        End Try
    End Function


End Class
