Public Class clsListener
    Inherits Object

    Private ThreadClose As New Threading.Thread(AddressOf CloseApp)
    Private WithEvents _SBO_Application As SAPbouiCOM.Application
    Private _Company As SAPbobsCOM.Company
    Private _Utilities As clsUtilities
    Private _Collection As Hashtable
    Private _LookUpCollection As Hashtable
    Private _FormUID As String
    Private _Log As clsLog_Error
    Private oMenuObject As Object
    Private oFormObject As Object
    Private oItemObject As Object
    Private oSystemForms As Object
    Dim objFilters As SAPbouiCOM.EventFilters
    Dim objFilter As SAPbouiCOM.EventFilter


#Region "New"
    Public Sub New()
        MyBase.New()
        Try
            _Company = New SAPbobsCOM.Company
            _Utilities = New clsUtilities
            _Collection = New Hashtable(10, 0.5)
            _LookUpCollection = New Hashtable(10, 0.5)
            oSystemForms = New clsSystemForms
            _Log = New clsLog_Error

            SetApplication()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Public Properties"

    Public ReadOnly Property SBO_Application() As SAPbouiCOM.Application
        Get
            Return _SBO_Application
        End Get
    End Property

    Public ReadOnly Property Company() As SAPbobsCOM.Company
        Get
            Return _Company
        End Get
    End Property

    Public ReadOnly Property Utilities() As clsUtilities
        Get
            Return _Utilities
        End Get
    End Property

    Public ReadOnly Property Collection() As Hashtable
        Get
            Return _Collection
        End Get
    End Property

    Public ReadOnly Property LookUpCollection() As Hashtable
        Get
            Return _LookUpCollection
        End Get
    End Property

    Public ReadOnly Property Log() As clsLog_Error
        Get
            Return _Log
        End Get
    End Property
#Region "Filter"

    Public Sub SetFilter(ByVal Filters As SAPbouiCOM.EventFilters)
        oApplication.SetFilter(Filters)
    End Sub
    Public Sub SetFilter()
        Try
            ''Form Load
            objFilters = New SAPbouiCOM.EventFilters


        Catch ex As Exception
            Throw ex
        End Try

    End Sub
#End Region

#End Region

#Region "Menu Event"

    Private Sub _SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.FormDataEvent
        Select Case BusinessObjectInfo.FormTypeEx
            Case frm_Budget
                oFormObject = New clsBudget
                oFormObject.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_BudgetEntry
                oFormObject = New clsBudgetEntry
                oFormObject.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_BudgetTemplate
                oFormObject = New clsBudgetTemplate
                oFormObject.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_Contract
                oFormObject = New clsContractAgreement
                oFormObject.FormDataEvent(BusinessObjectInfo, BubbleEvent)

        End Select
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case mnu_ContrctTimeSheet
                        oMenuObject = New clsContractTimesheet
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_PrjReports
                        oMenuObject = New clsprjReports
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case "Z_mnu_PRJ"
                        If pVal.BeforeAction = False Then
                            oApplication.SBO_Application.ActivateMenuItem("8457")
                        End If
                    Case mnu_ChangePassword
                        oMenuObject = New clsChangePassword
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_MyTasks
                        oMenuObject = New clsMyTasks
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Contract
                        oMenuObject = New clsContractAgreement
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_ACCClaim
                        oMenuObject = New clsExpaneEntry_Account
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_BOQDetails
                        oMenuObject = New clsBOQDefinition
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Activity
                        oMenuObject = New clsActivity
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_LeaveType
                        oMenuObject = New clsLeavType
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_EmpPosition
                        oMenuObject = New clsEmployeePosition
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_ExpClaim
                        oMenuObject = New clsExp_Posting
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Loginsetup
                        oMenuObject = New clsloginsetup
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Expances
                        oMenuObject = New clsexpances
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_report
                        oMenuObject = New clsReports
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_EmpTime
                        oMenuObject = New clsEmpTimeSheet
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_PrjTime, mnu_ExpApproval, mnu_LeaveApproval
                        oMenuObject = New clsPrjTime
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Moudule
                        oMenuObject = New clsModule
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Leaverequest
                        oMenuObject = New clsLeaverequest
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Budget
                        oMenuObject = New clsBudget
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_BudgetEntry
                        oMenuObject = New clsBudgetEntry
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_BudgetTemplate
                        oMenuObject = New clsBudgetTemplate
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_ExpEntry
                        oMenuObject = New clsExpaneEntry
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_ADD_ROW, mnu_DELETE_ROW, mnu_Duplicate_Row, mnu_DELETE_ROW, mnu_Delete, mnu_FIRST, mnu_LAST, mnu_PREVIOUS, mnu_NEXT, mnu_FIND, mnu_ADD
                        Dim oform As SAPbouiCOM.Form
                        oform = oApplication.SBO_Application.Forms.ActiveForm()
                        If oform.TypeEx = frm_ExpEntry Then
                            oMenuObject = New clsExpaneEntry
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        ElseIf oform.TypeEx = frm_Budget Then
                            oMenuObject = New clsBudget
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        ElseIf oform.TypeEx = frm_BudgetEntry Then
                            oMenuObject = New clsBudgetEntry
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        ElseIf oform.TypeEx = frm_BudgetTemplate Then
                            oMenuObject = New clsBudgetTemplate
                            oMenuObject.MenuEvent(pVal, BubbleEvent)

                        ElseIf oform.TypeEx = frm_ACCEntry Then
                            oMenuObject = New clsExpaneEntry_Account
                            oMenuObject.MenuEvent(pVal, BubbleEvent)

                        ElseIf oform.TypeEx = frm_Contract Then
                            oMenuObject = New clsContractAgreement
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        ElseIf oform.TypeEx = frm_ContrctTimeSheet Then
                            oMenuObject = New clsContractTimesheet
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        End If

                End Select
            Else
                Select Case pVal.MenuUID
                    Case mnu_ADD_ROW, mnu_DELETE_ROW, mnu_Duplicate_Row, mnu_DELETE_ROW, mnu_Delete, mnu_FIRST, mnu_LAST, mnu_PREVIOUS, mnu_NEXT, mnu_FIND, mnu_ADD
                        Dim oform As SAPbouiCOM.Form
                        oform = oApplication.SBO_Application.Forms.ActiveForm()
                        If oform.TypeEx = frm_ExpEntry Then
                            oMenuObject = New clsExpaneEntry
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        ElseIf oform.TypeEx = frm_Budget Then
                            oMenuObject = New clsBudget
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        ElseIf oform.TypeEx = frm_BudgetEntry Then
                            oMenuObject = New clsBudgetEntry
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        ElseIf oform.TypeEx = frm_BudgetTemplate Then
                            oMenuObject = New clsBudgetTemplate
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        ElseIf oform.TypeEx = frm_ACCEntry Then
                            oMenuObject = New clsExpaneEntry_Account
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        ElseIf oform.TypeEx = frm_Contract Then
                            oMenuObject = New clsContractAgreement
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        ElseIf oform.TypeEx = frm_ContrctTimeSheet Then
                            oMenuObject = New clsContractTimesheet
                            oMenuObject.MenuEvent(pVal, BubbleEvent)

                        End If
                End Select
            End If
        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            oMenuObject = Nothing
        End Try
    End Sub
#End Region

#Region "Item Event"
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.ItemEvent
        Try
            _FormUID = FormUID
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                Select Case pVal.FormTypeEx

                  
                End Select
            End If



            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD And pVal.Before_Action = False Then
                Select Case pVal.FormTypeEx
                    Case "60100"
                        Dim oForm As SAPbouiCOM.Form
                        oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        oApplication.Utilities.AddControls(oForm, "chkPrj", "48", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX, "RIGHT", 0, 0, "48", "Project Employee")
                        Dim och As SAPbouiCOM.CheckBox
                        och = oForm.Items.Item("chkPrj").Specific
                        och.DataBind.SetBound(True, "OHEM", "U_Z_ISPROJECT")
                    Case frm_ContractCFL
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsContractTimeCFL
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_ContrctTimeSheet
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsContractTimesheet
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_DisRule
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDisRule
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_BOMWizard
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsBOMCopy
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_PrjReports
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsprjReports
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ChangePassword
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsChangePassword
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_MyTasks
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsMyTasks
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ProjectMaster
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsProjectMaster
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_BudgetContract
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsBudgetDetails_Contract
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Contract
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsContractAgreement
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case "0"
                        Dim oform As SAPbouiCOM.Form
                        oform = oApplication.SBO_Application.Forms.Item(FormUID)
                        If oform.TypeEx = "0" Then
                            Try
                                Dim ostatic As SAPbouiCOM.StaticText
                                ostatic = oform.Items.Item("4").Specific
                                '  MsgBox(ostatic.Caption)
                                If ostatic.Caption.Contains("Would you like to save changes") = True Then
                                    oform.Items.Item("2").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                End If
                            Catch ex As Exception

                            End Try
                        End If
                    Case frm_EmployeePostion
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsEmployeePosition
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_ACCEntry
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsExpaneEntry_Account
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case "frm_BOQ"
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New ClsBOQ
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case "frm_ACT"
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsCFLActivity
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_BOQDeatils
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsBOQDefinition
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case "392", frm_PORequest, frm_ARCreditNote, frm_SalesQuotation, frm_SalesOrder, frm_Delivery, frm_Return, frm_GoodsIssue, frm_GoodsReceipt, frm_PO, frm_GRPO, frm_APReturn, frm_APInvoice, frm_APDOWN, frm_APDOWNREquest, frm_Reservce, frm_ARInvoice
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDocuments
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ChoosefromList
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsChooseFromList
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ChoosefromList1
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsChooseFromList_BOQ
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Report
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsReports
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_LeaveType
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsLeavType
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Leaverequest
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsLeaverequest
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_LoginSetup
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsloginsetup
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Activity
                        oItemObject = New clsActivity
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_Expances
                        oItemObject = New clsexpances
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_ExpEntry
                        oItemObject = New clsExpaneEntry
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_Module
                        oItemObject = New clsModule
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_Budget
                        oItemObject = New clsBudget
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_BudgetEntry
                        oItemObject = New clsBudgetEntry
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)

                    Case frm_BudgetTemplate
                        oItemObject = New clsBudgetTemplate
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_Timesheet
                        oItemObject = New clsEmpTimeSheet
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)

                    Case frm_Login
                        oItemObject = New clsLogin
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_Approal
                        oItemObject = New clsApproval
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_ProjectTime
                        oItemObject = New clsPrjTime
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                End Select
            End If
            If _Collection.ContainsKey(FormUID) Then
                oItemObject = _Collection.Item(FormUID)
                If oItemObject.IsLookUpOpen And pVal.BeforeAction = True Then
                    _SBO_Application.Forms.Item(oItemObject.LookUpFormUID).Select()
                    BubbleEvent = False
                    Exit Sub
                End If
                _Collection.Item(FormUID).ItemEvent(FormUID, pVal, BubbleEvent)
            End If
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.BeforeAction = False Then
                If _LookUpCollection.ContainsKey(FormUID) Then
                    oItemObject = _Collection.Item(_LookUpCollection.Item(FormUID))
                    If Not oItemObject Is Nothing Then
                        oItemObject.IsLookUpOpen = False
                    End If
                    _LookUpCollection.Remove(FormUID)
                End If

                If _Collection.ContainsKey(FormUID) Then
                    _Collection.Item(FormUID) = Nothing
                    _Collection.Remove(FormUID)
                End If

            End If

        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

#Region "Application Event"
    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles _SBO_Application.AppEvent
        Try
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged
                    oApplication.Utilities.Message("Project Management addon disconnected successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    _Utilities.AddRemoveMenus("RemoveMenus.xml")
                    CloseApp()
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        End Try
    End Sub
#End Region

#Region "Close Application"
    Private Sub CloseApp()
        Try
            If Not _SBO_Application Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_SBO_Application)
            End If

            If Not _Company Is Nothing Then
                If _Company.Connected Then
                    _Company.Disconnect()
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_Company)
            End If

            _Utilities = Nothing
            _Collection = Nothing
            _LookUpCollection = Nothing

            ThreadClose.Sleep(10)
            System.Windows.Forms.Application.Exit()
        Catch ex As Exception
            Throw ex
        Finally
            oApplication = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

#Region "Set Application"
    Private Sub SetApplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        Try
            If Environment.GetCommandLineArgs.Length > 1 Then
                sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
                SboGuiApi = New SAPbouiCOM.SboGuiApi
                SboGuiApi.Connect(sConnectionString)
                _SBO_Application = SboGuiApi.GetApplication()
            Else
                Throw New Exception("Connection string missing.")
            End If

        Catch ex As Exception
            Throw ex
        Finally
            SboGuiApi = Nothing
        End Try
    End Sub
#End Region

#Region "Finalize"
    Protected Overrides Sub Finalize()
        Try
            MyBase.Finalize()
            '            CloseApp()

            oMenuObject = Nothing
            oItemObject = Nothing
            oSystemForms = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Addon Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

    Private Sub _SBO_Application_LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.LayoutKeyEvent

    End Sub
End Class
