Imports System.Windows.Forms

Friend Class CatchingEvents

    Friend WithEvents SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Friend SBOCompany As SAPbobsCOM.Company '//OBJETO COMPAÑIA
    Friend csDirectory As String '//DIRECTORIO DONDE SE ENCUENTRAN LOS .SRF

    Public Sub New()
        MyBase.New()
        SetAplication()
        SetConnectionContext()
        ConnectSBOCompany()

        addMenuItems()

        setFilters()

    End Sub

    '//----- ESTABLECE LA COMUNICACION CON SBO
    Private Sub SetAplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String
        Try
            SboGuiApi = New SAPbouiCOM.SboGuiApi
            sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sConnectionString)
            SBOApplication = SboGuiApi.GetApplication()
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la aplicación SBO " & ex.Message)
            System.Windows.Forms.Application.Exit()
            End
        End Try
    End Sub

    '//----- ESTABLECE EL CONTEXTO DE LA APLICACION
    Private Sub SetConnectionContext()
        Try
            SBOCompany = SBOApplication.Company.GetDICompany
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con el DI")
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End
            'Finally
        End Try
    End Sub

    '//----- CONEXION CON LA BASE DE DATOS
    Private Sub ConnectSBOCompany()
        Dim loRecSet As SAPbobsCOM.Recordset
        Try
            '//ESTABLECE LA CONEXION A LA COMPAÑIA
            csDirectory = My.Application.Info.DirectoryPath
            If (csDirectory = "") Then
                System.Windows.Forms.Application.Exit()
                End
            End If
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la BD. " & ex.Message)
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End
        Finally
            loRecSet = Nothing
        End Try
    End Sub


    Private Sub addMenuItems()
        Dim loForm As SAPbouiCOM.Form = Nothing
        Dim loMenus As SAPbouiCOM.Menus
        Dim loMenusRoot As SAPbouiCOM.Menus
        Dim loMenuItem As SAPbouiCOM.MenuItem

        Try
            '////// Obtiene referencia de la forma Principal de Modulos
            loForm = SBOApplication.Forms.GetForm(169, 1)

            loForm.Freeze(True)

            '////// Obtiene la referencia de los Menus de SBO
            loMenus = SBOApplication.Menus.Item(6).SubMenus

            '////// Adiciona un Nuevo Menu para la Aplicacion de VectorSBO
            If loMenus.Exists("REC01") Then
                loMenus.RemoveEx("REC01")
            End If

            loMenuItem = loMenus.Add("REC01", "Liquidador de Saldos", SAPbouiCOM.BoMenuType.mt_POPUP, loMenus.Count)

            loMenusRoot = loMenuItem.SubMenus

            '////// Adiciona un menu Item
            If loMenusRoot.Exists("REC11") Then
                loMenusRoot.RemoveEx("REC11")
            End If
            loMenuItem = loMenusRoot.Add("REC11", "Liquidar Saldos", SAPbouiCOM.BoMenuType.mt_STRING, loMenusRoot.Count)
            loMenus = loMenuItem.SubMenus

            loForm.Freeze(False)
            loForm.Update()

        Catch ex As Exception
            If (Not loForm Is Nothing) Then
                loForm.Freeze(False)
                loForm.Update()
            End If
            SBOApplication.MessageBox("CatchingEvents. Error al agregar las opciones del menú. " & ex.Message)
            End
        Finally
            loMenus = Nothing
            loMenusRoot = Nothing
            loMenuItem = Nothing
        End Try
    End Sub


    '//----- ESTABLECE FILTROS DE EVENTOS DE LA APLICACION
    Private Sub setFilters()
        Dim lofilter As SAPbouiCOM.EventFilter
        Dim lofilters As SAPbouiCOM.EventFilters

        Try

            lofilters = New SAPbouiCOM.EventFilters
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            lofilter.AddEx("tekDelivery") '////// FORMA UDO DE ENTREGAS
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
            lofilter.AddEx("tekDelivery") '////// FORMA UDO DE ENTREGAS
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
            lofilter.AddEx("tekDelivery") '////// FORMA UDO DE ENTREGAS
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)

            SBOApplication.SetFilter(lofilters)

        Catch ex As Exception
            SBOApplication.MessageBox("SetFilter: " & ex.Message)
        End Try

    End Sub


    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ''// METODOS PARA EVENTOS DE LA APLICACION
    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBOApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                System.Windows.Forms.Application.Exit()
                End
        End Select

    End Sub


    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// CONTROLADOR DE EVENTOS MENU
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub SBOApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.MenuEvent
        Dim otekDel As FrmtekDel

        Try
            '//ANTES DE PROCESAR SBO
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    '//////////////////////////////////SubMenu de Crear traslado inventario////////////////////////
                    Case "DEL11"

                        otekDel = New FrmtekDel
                        otekDel.openForm(csDirectory)

                End Select
            End If

        Catch ex As Exception
            SBOApplication.MessageBox("clsCatchingEvents. MenuEvent " & ex.Message)
        Finally
            'oReservaPedido = Nothing
        End Try
    End Sub


    Private Sub SBOApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBOApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                System.Windows.Forms.Application.Exit()
                End
        End Select
    End Sub


    'Private Sub SBOApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.ItemEvent
    '    Try
    '        If pVal.Before_Action = False And pVal.FormTypeEx <> "" Then

    '            Select Case pVal.FormTypeEx
    '                '////////////////FORMA PARA ACTIVAR LICENCIA
    '                Case "tekDelivery"
    '                    FrmEntregaSBOControllerAfter(FormUID, pVal)
    '            End Select
    '        End If

    '    Catch ex As Exception
    '        SBOApplication.MessageBox("SBOApplication_ItemEvent. ItemEvent " & ex.Message)
    '    Finally
    '    End Try
    'End Sub

    Public Function addDelivery(ByVal FormUID As String, ByVal csDirectory As String)

        Dim oReconService As SAPbobsCOM.InternalReconciliationsService = SBOCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.InternalReconciliationsService)
        Dim openTrans As SAPbobsCOM.InternalReconciliationOpenTrans = oReconService.GetDataInterface(SAPbobsCOM.InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)
        Dim reconParams As SAPbobsCOM.InternalReconciliationParams = oReconService.GetDataInterface(SAPbobsCOM.InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)

        For i = 0 To 10 'gridRecon.DataTable.Rows.Count - 1


            If gridRecon.DataTable.GetValue("Select", gridRecon.GetDataTableRowIndex(i)) = "Y" Then

                SBOApplication.SetStatusBarMessage("Reconcilling Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, False) With openTrans 'For Incoming Payment '1st Line 

            If gridRecon.DataTable.GetValue("Type", gridRecon.GetDataTableRowIndex(i)) = "RC" Then
                    .InternalReconciliationOpenTransRows.Add()
                    .InternalReconciliationOpenTransRows.Item(x).Selected = SAPbobsCOM.BoYesNoEnum.tYES
                    .InternalReconciliationOpenTransRows.Item(x).TransId = gridRecon.DataTable.GetValue("TransId", gridRecon.GetDataTableRowIndex(i)) ' Journal Entry ID: TransId in OJDT 
                    .InternalReconciliationOpenTransRows.Item(x).TransRowId = 1 ' Journal Entry Line Number: Line_ID in JDT1 
                    oIncomPayment = Math.Abs(gridRecon.DataTable.GetValue("Actual Amount", gridRecon.GetDataTableRowIndex(i))) ' MsgBox(oIncomPayment) 'oTotal = oARPayment - oIncomPayment 
                    .InternalReconciliationOpenTransRows.Item(x).ReconcileAmount = oARPayment 'gridRecon.DataTable.GetValue("Actual Amount", gridRecon.GetDataTableRowIndex(i)) 
                    ' This should always be positive value. But one line should be on Credit, one line is Debit. 
                    oTransId = gridRecon.DataTable.GetValue("TransId", gridRecon.GetDataTableRowIndex(i)) ' 
                ElseIf gridRecon.DataTable.GetValue("Type", gridRecon.GetDataTableRowIndex(i)) = "JE" Then
                Else openTrans.CardOrAccount = SAPbobsCOM.CardOrAccountEnum.coaCard .InternalReconciliationOpenTransRows.Add() .InternalReconciliationOpenTransRows.Item(x).Selected = SAPbobsCOM.BoYesNoEnum.tYES .InternalReconciliationOpenTransRows.Item(x).TransId = gridRecon.DataTable.GetValue("TransId", gridRecon.GetDataTableRowIndex(i)) .InternalReconciliationOpenTransRows.Item(x).TransRowId = 0 ' Journal Entry Line Number: Line_ID in JDT1 .InternalReconciliationOpenTransRows.Item(x).ReconcileAmount = Math.Abs(gridRecon.DataTable.GetValue("Actual Amount", gridRecon.GetDataTableRowIndex(i))) ' This should always be positive value. But one line should be on Credit, one line is Debit. Console.WriteLine(Math.Abs(gridRecon.DataTable.GetValue("Actual Amount", gridRecon.GetDataTableRowIndex(i)))) If gridRecon.DataTable.GetValue("Type", gridRecon.GetDataTableRowIndex(i)) = "JE" Then oJeTranID = gridRecon.DataTable.GetValue("TransId", gridRecon.GetDataTableRowIndex(i)) olistPostedJE.Add(oJeTranID) End If End If If Not gridRecon.DataTable.GetValue("Type", gridRecon.GetDataTableRowIndex(i)) = "RC" Then oARPayment += Math.Abs(gridRecon.DataTable.GetValue("Actual Amount", i)) End If x = x + 1 End With End If 

        Next

        Try reconParams = oReconService.Add(openTrans) Catch ex As Exception SAP_APP.SetMessage(ex.ToString, SAPbouiCOM.BoStatusBarMessageType.smt_Error) oWriteText(Now & " - " & "[Err] - " & ex.ToString, True) oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack) Return False End Try Try oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit) Catch ex As Exception End Try

    End Function

End Class
