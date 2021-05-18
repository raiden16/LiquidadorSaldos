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
            lofilter.AddEx("tekReconciliation") '////// FORMA UDO DE ENTREGAS

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
        Dim otekLiq As FrmtekLIQ

        Try
            '//ANTES DE PROCESAR SBO
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    '//////////////////////////////////SubMenu de Crear traslado inventario////////////////////////
                    Case "REC11"

                        otekLiq = New FrmtekLIQ
                        otekLiq.openForm(csDirectory)

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


    Private Sub SBOApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.ItemEvent
        Try
            If pVal.Before_Action = False And pVal.FormTypeEx <> "" Then

                Select Case pVal.FormTypeEx
                    '////////////////FORMA PARA ACTIVAR LICENCIA
                    Case "tekReconciliation"
                        FrmLiquidarSBO(FormUID, pVal)
                End Select
            End If

        Catch ex As Exception
            SBOApplication.MessageBox("SBOApplication_ItemEvent. ItemEvent " & ex.Message)
        Finally
        End Try
    End Sub


    Private Sub FrmLiquidarSBO(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)

        Dim otekLiq As FrmtekLIQ
        Dim oLiquidar As Liquidar
        Dim oRecSetH As SAPbobsCOM.Recordset

        Try

            oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID

                        Case "4"

                            otekLiq = New FrmtekLIQ
                            otekLiq.AgregarLineas()

                        Case "5"

                            oLiquidar = New Liquidar
                            oLiquidar.Liquidar()

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("FrmLiquidarSBO. Error en forma de Panel General. " & ex.Message)
        Finally

        End Try
    End Sub


End Class
