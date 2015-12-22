Public Class clsDocuments
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private objMatrix As SAPbouiCOM.Matrix
    Private objForm As SAPbouiCOM.Form
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oCombo As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oRecset As SAPbobsCOM.Recordset
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo, sPath As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False

#Region "Methods"
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
  
#End Region

#Region "Events"



#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                'Case mnu_Activity
                '    If pVal.BeforeAction = False Then
                '        LoadForm()
                '    End If

                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD, mnu_FIND, mnu_ADD_ROW, mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    'If oForm.TypeEx = frm_Activity Then
                    '    If pVal.BeforeAction = False Then

                    '    End If
                    'End If
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
        End Try
    End Sub
#End Region


    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try

            If pVal.FormTypeEx = frm_GoodsIssue Or pVal.FormTypeEx = frm_GoodsReceipt Then
                Select Case pVal.Before_Action
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "13" And pVal.ColUID = "Z_BOQRef" Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "13" And (pVal.ColUID = "U_Z_ACTNAME" Or pVal.ColUID = "U_Z_MDNAME") And pVal.CharPressed = 9 Then
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList
                                    Dim strwhs, strProject, strGirdValue As String
                                    objMatrix = oForm.Items.Item("13").Specific
                                    strwhs = oApplication.Utilities.getMatrixValues(objMatrix, "21", pVal.Row)
                                    If strwhs = "" Then
                                        Exit Sub
                                    End If
                                    strGirdValue = oApplication.Utilities.getMatrixValues(objMatrix, "U_Z_ACTNAME", pVal.Row)
                                    If oApplication.Utilities.CheckModule_Activity(strwhs, "[@Z_PRJ1]", strGirdValue, "U_Z_ACTNAME") = False Then
                                        oApplication.Utilities.SetMatrixValues(objMatrix, "U_Z_ACTNAME", pVal.Row, "")
                                    Else
                                        Exit Sub
                                    End If
                                    If strwhs <> "" Then
                                        clsChooseFromList.ItemUID = "13"
                                        clsChooseFromList.SourceFormUID = FormUID
                                        clsChooseFromList.SourceLabel = pVal.Row
                                        clsChooseFromList.sourceItemCode = oApplication.Utilities.getMatrixValues(objMatrix, "1", pVal.Row)
                                        clsChooseFromList.CFLChoice = "[@Z_PRJ1]" 'oCombo.Selected.Value
                                        clsChooseFromList.choice = "ACTIVITY1"
                                        clsChooseFromList.ItemCode = strwhs
                                        clsChooseFromList.Documentchoice = "" ' oApplication.Utilities.GetDocType(oForm)
                                        clsChooseFromList.sourceColumID = pVal.ColUID
                                        clsChooseFromList.sourcerowId = pVal.Row
                                        clsChooseFromList.BinDescrUID = ""
                                        oApplication.Utilities.LoadForm("CFL.xml", frm_ChoosefromList)
                                        objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                        objChoose.databound(objChooseForm)
                                    End If
                                End If
                        End Select
                End Select
            End If

            If pVal.FormTypeEx = frm_PORequest Or pVal.FormTypeEx = frm_PO Or pVal.FormTypeEx = frm_ARInvoice Or pVal.FormTypeEx = frm_ARCreditNote Or pVal.FormTypeEx = frm_SalesQuotation Or pVal.FormTypeEx = frm_SalesOrder Or pVal.FormTypeEx = frm_Delivery Or pVal.FormTypeEx = frm_Return Or pVal.FormTypeEx = frm_GRPO Or pVal.FormTypeEx = frm_APReturn Or pVal.FormTypeEx = frm_APInvoice Or pVal.FormTypeEx = frm_APCR Or pVal.FormTypeEx = frm_APDOWN Or pVal.FormTypeEx = frm_APDOWNREquest Or pVal.FormTypeEx = frm_Reservce Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And pVal.ColUID = "Z_BOQRef" Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Z_MODNAME" Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "38" Or pVal.ItemUID = "39") And (pVal.ColUID = "U_Z_ACTNAME" Or pVal.ColUID = "U_Z_MDNAME") And pVal.CharPressed = 9 Then
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList
                                    Dim strwhs, strProject, strGirdValue As String
                                    objMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                    If pVal.ItemUID = "38" Then
                                        strwhs = oApplication.Utilities.getMatrixValues(objMatrix, "31", pVal.Row)
                                    Else
                                        strwhs = oApplication.Utilities.getMatrixValues(objMatrix, "4", pVal.Row)
                                    End If

                                    If strwhs = "" Then
                                        Exit Sub
                                    End If
                                    strGirdValue = oApplication.Utilities.getMatrixValues(objMatrix, "U_Z_ACTNAME", pVal.Row)
                                    If oApplication.Utilities.CheckModule_Activity(strwhs, "[@Z_PRJ1]", strGirdValue, "U_Z_ACTNAME") = False Then
                                        oApplication.Utilities.SetMatrixValues(objMatrix, "U_Z_ACTNAME", pVal.Row, "")
                                    Else
                                        Exit Sub
                                    End If
                                    If strwhs <> "" Then
                                        clsChooseFromList.ItemUID = pVal.ItemUID
                                        clsChooseFromList.SourceFormUID = FormUID
                                        clsChooseFromList.SourceLabel = pVal.Row
                                        clsChooseFromList.sourceItemCode = oApplication.Utilities.getMatrixValues(objMatrix, "1", pVal.Row)
                                        clsChooseFromList.CFLChoice = "[@Z_PRJ1]" 'oCombo.Selected.Value
                                        clsChooseFromList.choice = "ACTIVITY1"
                                        clsChooseFromList.ItemCode = strwhs
                                        clsChooseFromList.Documentchoice = "" ' oApplication.Utilities.GetDocType(oForm)
                                        clsChooseFromList.sourceColumID = pVal.ColUID
                                        clsChooseFromList.sourcerowId = pVal.Row
                                        clsChooseFromList.BinDescrUID = ""
                                        oApplication.Utilities.LoadForm("CFL.xml", frm_ChoosefromList)
                                        objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                        objChoose.databound(objChooseForm)
                                    End If
                                End If
                        End Select

                End Select
            End If


            If pVal.FormTypeEx = "392" Then ' Or pVal.FormTypeEx = frm_GRPO Or pVal.FormTypeEx = frm_APReturn Or pVal.FormTypeEx = frm_APInvoice Or pVal.FormTypeEx = frm_APCR Or pVal.FormTypeEx = frm_APDOWN Or pVal.FormTypeEx = frm_APDOWNREquest Or pVal.FormTypeEx = frm_Reservce Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "76" And pVal.ColUID = "Z_BOQRef" Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "76" And pVal.ColUID = "U_Z_MODNAME" Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "76") And (pVal.ColUID = "U_Z_ACTNAME" Or pVal.ColUID = "U_Z_MDNAME") And pVal.CharPressed = 9 Then
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList
                                    Dim strwhs, strProject, strGirdValue As String
                                    objMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                    'If pVal.ItemUID = "76" Then
                                    '    strwhs = oApplication.Utilities.getMatrixValues(objMatrix, "31", pVal.Row)
                                    'Else
                                    '    strwhs = oApplication.Utilities.getMatrixValues(objMatrix, "4", pVal.Row)
                                    'End If
                                    strwhs = oApplication.Utilities.getMatrixValues(objMatrix, "16", pVal.Row)
                                    If strwhs = "" Then
                                        Exit Sub
                                    End If
                                    strGirdValue = oApplication.Utilities.getMatrixValues(objMatrix, "U_Z_ACTNAME", pVal.Row)
                                    If oApplication.Utilities.CheckModule_Activity(strwhs, "[@Z_PRJ1]", strGirdValue, "U_Z_ACTNAME") = False Then
                                        oApplication.Utilities.SetMatrixValues(objMatrix, "U_Z_ACTNAME", pVal.Row, "")
                                    Else
                                        Exit Sub
                                    End If
                                    If strwhs <> "" Then
                                        clsChooseFromList.ItemUID = pVal.ItemUID
                                        clsChooseFromList.SourceFormUID = FormUID
                                        clsChooseFromList.SourceLabel = pVal.Row
                                        clsChooseFromList.sourceItemCode = oApplication.Utilities.getMatrixValues(objMatrix, "1", pVal.Row)
                                        clsChooseFromList.CFLChoice = "[@Z_PRJ1]" 'oCombo.Selected.Value
                                        clsChooseFromList.choice = "ACTIVITY1"
                                        clsChooseFromList.ItemCode = strwhs
                                        clsChooseFromList.Documentchoice = "" ' oApplication.Utilities.GetDocType(oForm)
                                        clsChooseFromList.sourceColumID = pVal.ColUID
                                        clsChooseFromList.sourcerowId = pVal.Row
                                        clsChooseFromList.BinDescrUID = ""
                                        oApplication.Utilities.LoadForm("CFL.xml", frm_ChoosefromList)
                                        objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                        objChoose.databound(objChooseForm)
                                    End If
                                End If
                        End Select

                End Select
            End If

        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#End Region
End Class
