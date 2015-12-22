Public Class clsContractTimesheet
    Inherits clsBase
#Region "Declarations"
    Public Shared ItemUID As String
    Public Shared SourceFormUID As String
    Public Shared SourceLabel As Integer
    Public Shared CFLChoice As String
    Public Shared ItemCode As String
    Public Shared sourceItemCode As String
    Public Shared choice As String
    Public Shared sourceColumID As String
    Public Shared sourcerowId As Integer
    Public Shared BinDescrUID As String
    Public Shared Documentchoice As String
    Public Shared ContractID As String

    Private oDbDatasource As SAPbouiCOM.DBDataSource
    Private Ouserdatasource As SAPbouiCOM.UserDataSource
    Private oConditions As SAPbouiCOM.Conditions
    Private ocondition As SAPbouiCOM.Condition
    Private intRowId As Integer
    Private strRowNum As Integer
    Private i As Integer
    Private oedit As SAPbouiCOM.EditTextColumn
    '   Private oForm As SAPbouiCOM.Form
    Private objSoureceForm As SAPbouiCOM.Form
    Private objForm As SAPbouiCOM.Form
    Private oMatrix, oGrid, oGrid1 As SAPbouiCOM.Grid
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private osourcegrid As SAPbouiCOM.Matrix
    Private oChecbox As SAPbouiCOM.CheckBoxColumn
    Private Const SEPRATOR As String = "~~~"
    Private SelectedRow As Integer
    Private sSearchColumn As String
    Private oItem As SAPbouiCOM.Item
    Public stritemid As SAPbouiCOM.Item
    Private intformmode As SAPbouiCOM.BoFormMode
    Private objGrid As SAPbouiCOM.Grid
    Private objSourcematrix As SAPbouiCOM.Matrix
    Private dtTemp As SAPbouiCOM.DataTable
    Private objStatic As SAPbouiCOM.StaticText
    Private inttable As Integer = 0
    Public strformid As String
    Public strStaticValue As String
    Public strSQL As String
    Private strSelectedItem1 As String
    Private strSelectedItem2 As String
    Private strSelectedItem3 As String
    Private strSelectedItem4 As String
    Private strContractID, strContractName, strCustCntID As String
    Private oRecSet As SAPbobsCOM.Recordset
    '   Private objSBOAPI As ClsSBO
    '   Dim objTransfer As clsTransfer
#End Region

#Region "New"
    '*****************************************************************
    'Type               : Constructor
    'Name               : New
    'Parameter          : 
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Create object for classes.
    '******************************************************************
    Public Sub New()
        '   objSBOAPI = New ClsSBO
        MyBase.New()
    End Sub
#End Region

    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_ContrctTimeSheet) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_ContrctTimeSheet, frm_ContrctTimeSheet)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        AddChooseFromList(oForm)
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        oForm.PaneLevel = 1
        oForm.Freeze(True)
        databind(oForm)
        oForm.Freeze(False)
       
    End Sub

    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition


            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            'oCFLCreationParams.MultiSelection = False
            'oCFLCreationParams.ObjectType = "Z_Module"
            'oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Item("CFL_2")

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_TYPE"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            'oCFLCreationParams.ObjectType = "Z_Module"
            'oCFLCreationParams.UniqueID = "CFL2"
            'oCFL = oCFLs.Add(oCFLCreationParams)
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "U_Z_Status"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "N"
            'oCFL.SetConditions(oCons)
            'oCon = oCons.Add()





        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#Region "Bind Data"
    '******************************************************************
    'Type               : Procedure
    'Name               : BindData
    'Parameter          : Form
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Binding the fields.
    '******************************************************************
    Public Sub databound(ByVal objform As SAPbouiCOM.Form)
        Try
            Dim strSQL As String = ""
            Dim ObjSegRecSet As SAPbobsCOM.Recordset
            ObjSegRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oform = objform
            oform.Freeze(True)

            oGrid = oform.Items.Item("6").Specific
            oGrid1 = oform.Items.Item("7").Specific
            oGrid.DataTable.ExecuteQuery("Select ItemCode,ItemName,' ' 'Select' from OITM where TreeType<>'N'")
            ' oGrid1.DataTable.ExecuteQuery("SELECT T1.[Father], T1.[ChildNum], T1.[Code],T2.[ItemName] FROM OITT T0  INNER JOIN ITT1 T1 ON T0.Code = T1.Father INNER JOIN OITM T2 ON T1.Code = T2.ItemCode  where 1=2 ORDER BY T1.[Father]")
            Dim st As String = "SELECT T1.[Father] 'ParentItem', T1.[Code] 'ItemCode',T2.[ItemName], T2.cardcode 'CardCode',T3.CardName,T2.InvntryUom 'UoM'  ,T1.Quantity 'Quantity',' ' 'Select' FROM OITT T0  INNER JOIN ITT1 T1 ON T0.Code = T1.Father Inner JOIN OITM T2 ON T1.Code = T2.ItemCode left outer join OCRD T3 on T3.CardCode=T2.CardCode "
            oGrid1.DataTable.ExecuteQuery(st & "  where 1=2 ORDER BY T1.[Father]")
            formatGrid(oGrid, "BOM", oform)
            formatGrid(oGrid1, "Item", oform)
            oform.PaneLevel = 1
            oform.Freeze(False)
        Catch ex As Exception
            oApplication.SBO_Application.MessageBox(ex.Message)
        Finally
        End Try
    End Sub
#End Region
    Private Sub databind(ByVal aform As SAPbouiCOM.Form)
        Dim strQuery, StrQuery1, strContractID As String
        strContractID = oApplication.Utilities.getEdittextvalue(aform, "22")
        If strContractID = "" Then
            strContractID = "99999"
        End If
        oGrid = aform.Items.Item("11").Specific
        oGrid1 = aform.Items.Item("12").Specific
        strQuery = "select T0.Code, T1.U_Z_CNTID,T1.U_Z_EMPCODE,T1.U_Z_EMPNAME,T0.U_Z_PRJCODE,U_Z_PRJNAME,U_Z_PRCNAME,U_Z_ACTNAME,U_Z_DATE,U_Z_HOURS,U_Z_REMARKS,U_Z_APPROVED,T0.U_Z_TYPE ,U_Z_EMPAPPROVAL from [@Z_TIM1] T0 inner Join [@Z_OTIM] T1 on T1.Code=T0.U_Z_REFCODE where T0.U_Z_APPROVED='P' and  T1.U_Z_CNTID='" & strContractID & "'"
        StrQuery1 = "select T0.Code, T1.U_Z_CNTID,T1.U_Z_EMPCODE,T1.U_Z_EMPNAME,T0.U_Z_PRJCODE,U_Z_PRJNAME,U_Z_PRCNAME,U_Z_ACTNAME,U_Z_DATE,U_Z_HOURS,U_Z_REMARKS,U_Z_APPROVED,T0.U_Z_TYPE ,U_Z_EMPAPPROVAL from [@Z_TIM1] T0 inner Join [@Z_OTIM] T1 on T1.Code=T0.U_Z_REFCODE where T1.U_Z_CNTID='" & strContractID & "'"
        oGrid.DataTable.ExecuteQuery(strQuery)
        formatGrid(oGrid, "Trans", aform)
        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        oGrid1.DataTable.ExecuteQuery(StrQuery1)
        formatGrid(oGrid1, "Summary", aform)
        aform.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
    End Sub
    Private Sub AssignLineNo1(ByVal aform As SAPbouiCOM.Form, ByVal agrid As SAPbouiCOM.Grid)
        Try
            aform.Freeze(True)
            For count As Integer = 0 To agrid.DataTable.Rows.Count - 1
                agrid.RowHeaders.SetText(count, count + 1)
            Next
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub
    Private Sub formatGrid(ByVal aGrid As SAPbouiCOM.Grid, ByVal aChoice As String, ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        Select Case aChoice
            Case "Trans"
                aGrid.Columns.Item("Code").TitleObject.Caption = "Code"
                aGrid.Columns.Item("Code").Visible = False
                aGrid.Columns.Item("U_Z_CNTID").TitleObject.Caption = "ContractID"
                aGrid.Columns.Item("U_Z_CNTID").Visible = False
                aGrid.Columns.Item("U_Z_EMPCODE").TitleObject.Caption = "Employee Code"
                aGrid.Columns.Item("U_Z_EMPNAME").TitleObject.Caption = "Employee Name"
                aGrid.Columns.Item("U_Z_EMPNAME").Editable = False
                aGrid.Columns.Item("U_Z_PRJCODE").TitleObject.Caption = "Project Code"
                aGrid.Columns.Item("U_Z_PRJNAME").TitleObject.Caption = "Project Name"
                aGrid.Columns.Item("U_Z_PRJNAME").Editable = False
                aGrid.Columns.Item("U_Z_PRCNAME").TitleObject.Caption = "Phase"
                aGrid.Columns.Item("U_Z_ACTNAME").TitleObject.Caption = "Activity"
                aGrid.Columns.Item("U_Z_ACTNAME").Editable = False
                aGrid.Columns.Item("U_Z_DATE").TitleObject.Caption = "Date"
                aGrid.Columns.Item("U_Z_HOURS").TitleObject.Caption = "Hours"
                aGrid.Columns.Item("U_Z_APPROVED").TitleObject.Caption = "Approved"
                aGrid.Columns.Item("U_Z_APPROVED").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                oComboColumn = aGrid.Columns.Item("U_Z_APPROVED")
                oComboColumn.ValidValues.Add("P", "Pending")
                oComboColumn.ValidValues.Add("A", "Approved")
                oComboColumn.ValidValues.Add("D", "Declained")
                oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                oComboColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

                aGrid.Columns.Item("U_Z_REMARKS").TitleObject.Caption = "Remarks"
                aGrid.Columns.Item("U_Z_EMPAPPROVAL").TitleObject.Caption = "Employee Approval Status"
                aGrid.Columns.Item("U_Z_EMPAPPROVAL").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                oComboColumn = aGrid.Columns.Item("U_Z_EMPAPPROVAL")
                oComboColumn.ValidValues.Add("P", "Pending")
                oComboColumn.ValidValues.Add("C", "Confirmed")
                oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                oComboColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                aGrid.Columns.Item("U_Z_TYPE").TitleObject.Caption = "TimeSheet Type"
                aGrid.Columns.Item("U_Z_TYPE").Visible = False
                aGrid.AutoResizeColumns()
                aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single


                '  aGrid.Columns.Item("Quantity").TitleObject.Caption = "Quantity"
            Case "Summary"
                aGrid.Columns.Item("Code").TitleObject.Caption = "Code"
                aGrid.Columns.Item("Code").Visible = False
                aGrid.Columns.Item("U_Z_CNTID").TitleObject.Caption = "ContractID"
                aGrid.Columns.Item("U_Z_CNTID").Visible = False
                aGrid.Columns.Item("U_Z_EMPCODE").TitleObject.Caption = "Employee Code"
                aGrid.Columns.Item("U_Z_EMPCODE").Editable = False
                aGrid.Columns.Item("U_Z_EMPNAME").TitleObject.Caption = "Employee Name"
                aGrid.Columns.Item("U_Z_EMPNAME").Editable = False
                aGrid.Columns.Item("U_Z_PRJCODE").TitleObject.Caption = "Project Code"
                aGrid.Columns.Item("U_Z_PRJCODE").Editable = False
                aGrid.Columns.Item("U_Z_PRJNAME").TitleObject.Caption = "Project Name"
                aGrid.Columns.Item("U_Z_PRJNAME").Editable = False
                aGrid.Columns.Item("U_Z_PRCNAME").TitleObject.Caption = "Phase"
                aGrid.Columns.Item("U_Z_PRCNAME").Editable = False
                aGrid.Columns.Item("U_Z_ACTNAME").TitleObject.Caption = "Activity"
                aGrid.Columns.Item("U_Z_ACTNAME").Editable = False
                aGrid.Columns.Item("U_Z_DATE").TitleObject.Caption = "Date"
                aGrid.Columns.Item("U_Z_DATE").Editable = False
                aGrid.Columns.Item("U_Z_HOURS").TitleObject.Caption = "Hours"
                aGrid.Columns.Item("U_Z_HOURS").Editable = False
                aGrid.Columns.Item("U_Z_APPROVED").TitleObject.Caption = "Approved"
                aGrid.Columns.Item("U_Z_APPROVED").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                oComboColumn = aGrid.Columns.Item("U_Z_APPROVED")
                oComboColumn.ValidValues.Add("P", "Pending")
                oComboColumn.ValidValues.Add("A", "Approved")
                oComboColumn.ValidValues.Add("D", "Declained")
                oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                oComboColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                aGrid.Columns.Item("U_Z_APPROVED").Editable = False
                aGrid.Columns.Item("U_Z_EMPAPPROVAL").Editable = False
                aGrid.Columns.Item("U_Z_REMARKS").TitleObject.Caption = "Remarks"
                aGrid.Columns.Item("U_Z_REMARKS").Editable = False
                aGrid.Columns.Item("U_Z_EMPAPPROVAL").TitleObject.Caption = "Employee Approval Status"
                aGrid.Columns.Item("U_Z_EMPAPPROVAL").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                oComboColumn = aGrid.Columns.Item("U_Z_EMPAPPROVAL")
                oComboColumn.ValidValues.Add("P", "Pending")
                oComboColumn.ValidValues.Add("C", "Confirmed")
                oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                oComboColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                aGrid.Columns.Item("U_Z_EMPAPPROVAL").Editable = False
                aGrid.Columns.Item("U_Z_TYPE").TitleObject.Caption = "TimeSheet Type"
                aGrid.Columns.Item("U_Z_TYPE").Visible = False
                aGrid.AutoResizeColumns()
                aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
        End Select
        AssignLineNo(aform, aGrid)
        aform.Freeze(False)
    End Sub


#Region "validations"
    Private Function validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strCode, strcode1, strProject, strPhase, strActivity, strDate, strEmp As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oGrid = aform.Items.Item("11").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            strCode = oGrid.DataTable.GetValue("U_Z_PRJCODE", intRow)
            If strCode <> "" Then
                strProject = strCode
                strcode1 = oGrid.DataTable.GetValue("U_Z_EMPCODE", intRow)
                If strcode1 = "" Then
                    oApplication.Utilities.Message("Employee Code is missing . Line Number : " & intRow + 1, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oGrid.Columns.Item("U_Z_EMPCODE").Click(intRow)
                    Return False
                Else
                    strPhase = strcode1
                End If
                strcode1 = oGrid.DataTable.GetValue("U_Z_PRCNAME", intRow)
                If strcode1 = "" Then
                    oApplication.Utilities.Message("Phase is missing . Line Number : " & intRow + 1, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oGrid.Columns.Item("U_Z_PRCNAME").Click(intRow)
                    Return False
                Else
                    strPhase = strcode1
                End If

                strcode1 = oGrid.DataTable.GetValue("U_Z_DATE", intRow)
                If strcode1 = "" Then
                    oApplication.Utilities.Message("Date is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oGrid.Columns.Item("U_Z_DATE").Click(intRow)
                    Return False
                End If
                strcode1 = oGrid.DataTable.GetValue("U_Z_HOURS", intRow)
                If strcode1 = "" Then
                    oApplication.Utilities.Message("Hours is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oGrid.Columns.Item("U_Z_HOURS").Click(intRow)
                    Return False
                End If
                If CDbl(strcode1) <= 0 Then
                    oApplication.Utilities.Message("Hours is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oGrid.Columns.Item("U_Z_HOURS").Click(intRow)
                    Return False
                End If

            End If
        Next
        Return True
    End Function
#End Region
#Region "Update On hand Qty"
    Private Sub FillOnhandqty(ByVal strItemcode As String, ByVal strwhs As String, ByVal aGrid As SAPbouiCOM.Grid)
        Dim oTemprec As SAPbobsCOM.Recordset
        oTemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim strBin, strSql As String
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            strBin = aGrid.DataTable.GetValue(0, intRow)
            strSql = "Select isnull(Sum(U_InQty)-sum(U_OutQty),0) from [@DABT_BTRN] where U_Itemcode='" & strItemcode & "' and U_FrmWhs='" & strwhs & "' and U_BinCode='" & strBin & "'"
            oTemprec.DoQuery(strSql)
            Dim dblOnhand As Double
            dblOnhand = oTemprec.Fields.Item(0).Value

            aGrid.DataTable.SetValue(2, intRow, dblOnhand.ToString)
        Next
    End Sub
#End Region

#Region "Get Form"
    '******************************************************************
    'Type               : Function
    'Name               : GetForm
    'Parameter          : FormUID
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Get The Form
    '******************************************************************
    Public Function GetForm(ByVal FormUID As String) As SAPbouiCOM.Form
        Return oApplication.SBO_Application.Forms.Item(FormUID)
    End Function
#End Region

    Private Sub addRow(ByVal aform As SAPbouiCOM.Form)
        If aform.PaneLevel = 2 Then
            aform.Freeze(True)
            oGrid = aform.Items.Item("11").Specific
            If oGrid.DataTable.Rows.Count > 0 Then
                If oGrid.DataTable.GetValue("U_Z_EMPCODE", oGrid.DataTable.Rows.Count - 1) <> "" Then
                    oGrid.DataTable.Rows.Add()
                    oGrid.DataTable.SetValue("U_Z_CNTID", oGrid.DataTable.Rows.Count - 1, oApplication.Utilities.getEdittextvalue(aform, "22"))
                    oGrid.DataTable.SetValue("U_Z_APPROVED", oGrid.DataTable.Rows.Count - 1, "P")
                    oGrid.DataTable.SetValue("U_Z_EMPAPPROVAL", oGrid.DataTable.Rows.Count - 1, "P")

                End If
            Else
                oGrid.DataTable.Rows.Add()
                oGrid.DataTable.SetValue("U_Z_CNTID", oGrid.DataTable.Rows.Count - 1, oApplication.Utilities.getEdittextvalue(aform, "22"))
                oGrid.DataTable.SetValue("U_Z_APPROVED", oGrid.DataTable.Rows.Count - 1, "P")
                oGrid.DataTable.SetValue("U_Z_EMPAPPROVAL", oGrid.DataTable.Rows.Count - 1, "P")
            End If
            oGrid.Columns.Item("U_Z_EMPCODE").Click(oGrid.DataTable.Rows.Count - 1)
            AssignLineNo(aform)
            aform.Freeze(False)
        End If
        Exit Sub
    End Sub

    Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form)
        Dim strLineCode, strLineApproval, strEmpApproval As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oGrid = aForm.Items.Item("11").Specific
        If aForm.PaneLevel = 2 Then
            For introw As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.Rows.IsSelected(introw) Then
                    ' oComboColumn = oGrid.Columns.Item("U_Z_APPROVED")
                    strLineApproval = oGrid.DataTable.GetValue("U_Z_APPROVED", introw) ' oComboColumn.GetSelectedValue(introw).Value
                    If strLineApproval = "P" Or strLineApproval = "" Then
                        strLineCode = oGrid.DataTable.GetValue("Code", introw)
                        If strLineCode <> "" Then
                            oTemp.DoQuery("Update [@Z_TIM1] set Name=Name +'_XD' where code='" & strLineCode & "'")
                        End If
                        oGrid.DataTable.Rows.Remove(introw)
                        AssignLineNo(aForm)
                        Exit Sub
                    Else
                        oApplication.Utilities.Message("You can not delete Approved / Declined time sheet entries", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oGrid = aForm.Items.Item("11").Specific
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next
           
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try

    End Sub


#Region "AddtoUDT"

    Private Function addto_HeaderTable(ByVal aEmpId As String, ByVal aEmpName As String, ByVal aDate As Date, ByVal aContractID As String) As String
        Dim oTemp As SAPbobsCOM.Recordset
        Dim oCode As String
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "Select * from [@Z_OTIM] where U_Z_Type='C' and U_Z_CntID='" & aContractID & "' and U_Z_EMPCODE='" & aEmpId & "' and convert(varchar(10),U_Z_DocDate,105)='" & aDate.ToString("dd-MM-yyyy") & "'"
        oTemp.DoQuery(strSQL)
        If oTemp.RecordCount > 0 Then
            oCode = oTemp.Fields.Item("Code").Value
        Else
            oCode = oApplication.Utilities.getMaxCode("@Z_OTIM", "Code")
            Dim ousertable1 As SAPbobsCOM.UserTable
            ousertable1 = oApplication.Company.UserTables.Item("Z_OTIM")
            ousertable1.Code = oCode
            ousertable1.Name = oCode
            ousertable1.UserFields.Fields.Item("U_Z_EMPCODE").Value = aEmpId
            ousertable1.UserFields.Fields.Item("U_Z_EMPNAME").Value = aEmpName
            ousertable1.UserFields.Fields.Item("U_Z_DocDate").Value = aDate
            ousertable1.UserFields.Fields.Item("U_Z_Type").Value = "C"
            ousertable1.UserFields.Fields.Item("U_Z_CntID").Value = aContractID
            If ousertable1.Add <> 0 Then
                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return ""
            Else
                Return oCode
            End If
        End If
        Return oCode
    End Function
    Private Function AddToUDT_Table(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim stContractID, strCode, strDocref, strEmpID, strLineCode, stdocdate, strProjectName, strEmpName, strProject, strProcess, strdate, strActivtiy, strhours, stremptype, strPrjCode, strAmount As String
        Dim dtDate As Date
        Dim intHours, dblAmount As Double
        Dim oTempRec, otemp As SAPbobsCOM.Recordset
        Dim ousertable As SAPbobsCOM.UserTable
        Dim ocheckbox As SAPbouiCOM.CheckBoxColumn
        Dim oedittext As SAPbouiCOM.EditTextColumn
        Dim dblPercentage As Double
        Dim blnexits As Boolean = False
        Dim blnLines As Boolean = False
        Dim dtFrom, dtTo As Date
        Dim oBPGrid As SAPbouiCOM.Matrix

        oGrid = aform.Items.Item("11").Specific
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        stContractID = oApplication.Utilities.getEdittextvalue(aform, "22")
        ousertable = oApplication.Company.UserTables.Item("Z_OTIM")
        '  oTempRec.DoQuery("Select isnull(firstName,'') + ' ' + isnull(middleName,'')+' '+isnull(lastName,'') from OHEM where ""empID""=" & strEmpID)
        '  strEmpName = oTempRec.Fields.Item(0).Value
        blnLines = True
        If blnLines = False Then
            ' oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Else
            For intRowId As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                strPrjCode = oGrid.DataTable.GetValue("U_Z_PRJCODE", intRowId)
                strEmpID = oGrid.DataTable.GetValue("U_Z_EMPCODE", intRowId)
                strEmpName = oGrid.DataTable.GetValue("U_Z_EMPNAME", intRowId)
                If strPrjCode <> "" And strEmpID <> "" Then
                    strProcess = oGrid.DataTable.GetValue("U_Z_PRCNAME", intRowId)
                    strActivtiy = oGrid.DataTable.GetValue("U_Z_ACTNAME", intRowId)
                    strProjectName = oGrid.DataTable.GetValue("U_Z_PRJNAME", intRowId)
                    dtDate = oGrid.DataTable.GetValue("U_Z_DATE", intRowId)
                    strdate = oGrid.DataTable.GetValue("U_Z_DATE", intRowId)
                    strhours = oGrid.DataTable.GetValue("U_Z_HOURS", intRowId)
                    strLineCode = oGrid.DataTable.GetValue("Code", intRowId)
                    dblAmount = oApplication.Utilities.getDocumentQuantity(strhours)
                    strDocref = addto_HeaderTable(strEmpID, strEmpName, dtDate, stContractID)
                    ousertable = oApplication.Company.UserTables.Item("Z_TIM1")
                    Dim oTe As SAPbobsCOM.Recordset
                    oTe = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If strDocref <> "" Then
                        If strLineCode = "" Then
                            strCode = oApplication.Utilities.getMaxCode("@Z_TIM1", "Code")
                            ousertable.Code = strCode
                            ousertable.Name = strCode
                            ousertable.UserFields.Fields.Item("U_Z_PRJCODE").Value = strPrjCode
                            ousertable.UserFields.Fields.Item("U_Z_PRJNAME").Value = strProjectName
                            ousertable.UserFields.Fields.Item("U_Z_PRCNAME").Value = strProcess
                            ousertable.UserFields.Fields.Item("U_Z_ACTNAME").Value = strActivtiy
                            If strdate <> "" Then
                                ousertable.UserFields.Fields.Item("U_Z_DATE").Value = dtDate
                            End If
                            Dim strBusiness, strsql As String

                            oTe = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            strProject = strPrjCode ' oApplication.Utilities.getMatrixValues(oBPGrid, "V_1", intLoop)
                            strBusiness = strProcess ' oApplication.Utilities.getMatrixValues(oBPGrid, "V_3", intLoop)
                            strActivtiy = strActivtiy '  oApplication.Utilities.getMatrixValues(oBPGrid, "V_4", intLoop)
                            '  strsql = "Select * from [@Z_PRJ1] T0 inner Join [@Z_HPRJ] T1 on T1.Code=T0.U_Z_REFCODE and U_Z_EMPCODE='" & oApplication.Utilities.getEdittextvalue(aform, "4") & "' where U_Z_MODNAME='" & strBusiness & "' and U_Z_ActName='" & strActivtiy & "' and  T1.U_Z_PRJCODE='" & strProject & "')"
                            strsql = "Select * from [@Z_PRJ1] T0 inner Join [@Z_HPRJ] T1 on T1.DocEntry=T0.DocEntry and T0.U_Z_CNTID='" & stContractID & "' where U_Z_MODNAME='" & strBusiness & "' and U_Z_ActName='" & strActivtiy & "' and  T1.U_Z_PRJCODE='" & strProject & "'"
                            oTe.DoQuery(strsql)
                            If oTe.RecordCount > 0 Then
                                '    oApplication.Utilities.SetMatrixValues(oBPGrid, "BdgQty", intLoop, oTe.Fields.Item("U_Z_HOURS").Value)
                                '    oApplication.Utilities.SetMatrixValues(oBPGrid, "Measure", intLoop, oTe.Fields.Item("U_Z_Measure").Value)
                            End If

                            '   ousertable.UserFields.Fields.Item("U_Z_BDGQTY").Value = 0 'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oBPGrid, "BdgQty", intLoop))
                            ' ousertable.UserFields.Fields.Item("U_Z_QUANTITY").Value = 0 ' oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oBPGrid, "Qty", intLoop))
                            '  ousertable.UserFields.Fields.Item("U_Z_MEASURE").Value = 0 ' oApplication.Utilities.getMatrixValues(oBPGrid, "Measure", intLoop)
                            ousertable.UserFields.Fields.Item("U_Z_HOURS").Value = dblAmount
                            ' oCombobox = oBPGrid.Columns.Item("Type").Cells.Item(intLoop).Specific
                            ousertable.UserFields.Fields.Item("U_Z_TYPE").Value = "R" ' oCombobox.Selected.Value
                            Dim strEmpConfirmaion As String
                            oComboColumn = oGrid.Columns.Item("U_Z_EMPAPPROVAL")
                            Try
                                ousertable.UserFields.Fields.Item("U_Z_EmpApproval").Value = oComboColumn.GetSelectedValue(intRowId).Value
                                strEmpConfirmaion = oComboColumn.GetSelectedValue(intRowId).Value
                            Catch ex As Exception
                                ousertable.UserFields.Fields.Item("U_Z_EmpApproval").Value = "P"
                                strEmpConfirmaion = "P"
                            End Try
                            oComboColumn = oGrid.Columns.Item("U_Z_APPROVED")
                            Dim otest As SAPbobsCOM.Recordset
                            otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            otest.DoQuery("SELECT T0.[U_Z_PRJCODE], T0.[U_Z_APPROVAL],* FROM [dbo].[@Z_HPRJ]  T0 where T0.U_Z_PRJCODE='" & strPrjCode & "'")
                            If otest.Fields.Item("U_Z_APPROVAL").Value = "N" Then
                                If strEmpConfirmaion = "C" Then
                                    ousertable.UserFields.Fields.Item("U_Z_Approved").Value = "A" ' oCombobox.Selected.Value
                                Else
                                    ousertable.UserFields.Fields.Item("U_Z_Approved").Value = "P" ' oCombobox.Selected.Value
                                End If
                            Else
                                Try
                                    ousertable.UserFields.Fields.Item("U_Z_Approved").Value = oComboColumn.GetSelectedValue(intRowId).Value
                                Catch ex As Exception
                                    ousertable.UserFields.Fields.Item("U_Z_Approved").Value = "P"
                                End Try

                            End If
                            ousertable.UserFields.Fields.Item("U_Z_RefCode").Value = strDocref
                            ousertable.UserFields.Fields.Item("U_Z_Remarks").Value = oGrid.DataTable.GetValue("U_Z_REMARKS", intRowId)
                            If ousertable.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            Else
                                Dim strCode1, strQuery1 As String
                                Dim oTest12 As SAPbobsCOM.Recordset
                                oTest12 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                strCode1 = strDocref
                                strQuery1 = "Select *  from [@Z_OTIM]  T0 inner Join [@Z_TIM1] T1  on T1.U_Z_REFCODE=T0.Code  where T0.Code='" & strCode1 & "'"
                                oTest12.DoQuery(strQuery1)
                                If oTest12.RecordCount > 0 Then
                                    '               oApplication.Utilities.ClosingProjectActivity(oTest12.Fields.Item("U_Z_EMPCODE").Value, oTest12.Fields.Item("U_Z_PRJCODE").Value, oTest12.Fields.Item("U_Z_PRCNAME").Value, oTest12.Fields.Item("U_Z_ACTNAME").Value, oTest12.Fields.Item("U_Z_DATE").Value)
                                End If

                            End If

                        Else
                            If ousertable.GetByKey(strLineCode) Then
                                ousertable.Code = strLineCode
                                ousertable.Name = strLineCode
                                ousertable.UserFields.Fields.Item("U_Z_PRJCODE").Value = strPrjCode
                                ousertable.UserFields.Fields.Item("U_Z_PRJNAME").Value = strProjectName
                                ousertable.UserFields.Fields.Item("U_Z_PRCNAME").Value = strProcess
                                ousertable.UserFields.Fields.Item("U_Z_ACTNAME").Value = strActivtiy
                                If strdate <> "" Then
                                    ousertable.UserFields.Fields.Item("U_Z_DATE").Value = dtDate
                                End If
                                Dim strBusiness, strsql As String
                                '    Dim oTe As SAPbobsCOM.Recordset
                                oTe = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                strProject = strPrjCode ' oApplication.Utilities.getMatrixValues(oBPGrid, "V_1", intLoop)
                                strBusiness = strProcess ' oApplication.Utilities.getMatrixValues(oBPGrid, "V_3", intLoop)
                                strActivtiy = strActivtiy '  oApplication.Utilities.getMatrixValues(oBPGrid, "V_4", intLoop)
                                '  strsql = "Select * from [@Z_PRJ1] T0 inner Join [@Z_HPRJ] T1 on T1.Code=T0.U_Z_REFCODE and U_Z_EMPCODE='" & oApplication.Utilities.getEdittextvalue(aform, "4") & "' where U_Z_MODNAME='" & strBusiness & "' and U_Z_ActName='" & strActivtiy & "' and  T1.U_Z_PRJCODE='" & strProject & "')"
                                strsql = "Select * from [@Z_PRJ1] T0 inner Join [@Z_HPRJ] T1 on T1.DocEntry=T0.DocEntry and T0.U_Z_CNTID='" & stContractID & "' where U_Z_MODNAME='" & strBusiness & "' and U_Z_ActName='" & strActivtiy & "' and  T1.U_Z_PRJCODE='" & strProject & "'"
                                oTe.DoQuery(strsql)
                                If oTe.RecordCount > 0 Then
                                    '   oApplication.Utilities.SetMatrixValues(oBPGrid, "BdgQty", intLoop, oTe.Fields.Item("U_Z_HOURS").Value)
                                    '     oApplication.Utilities.SetMatrixValues(oBPGrid, "Measure", intLoop, oTe.Fields.Item("U_Z_Measure").Value)
                                End If

                                '  ousertable.UserFields.Fields.Item("U_Z_BDGQTY").Value = 0 'oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oBPGrid, "BdgQty", intLoop))
                                '  ousertable.UserFields.Fields.Item("U_Z_QUANTITY").Value = 0 ' oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oBPGrid, "Qty", intLoop))
                                '    ousertable.UserFields.Fields.Item("U_Z_MEASURE").Value = 0 ' oApplication.Utilities.getMatrixValues(oBPGrid, "Measure", intLoop)
                                ousertable.UserFields.Fields.Item("U_Z_HOURS").Value = dblAmount
                                ' oCombobox = oBPGrid.Columns.Item("Type").Cells.Item(intLoop).Specific
                                ousertable.UserFields.Fields.Item("U_Z_TYPE").Value = "R" ' oCombobox.Selected.Value
                                Dim strEmpConfirmaion As String
                                oComboColumn = oGrid.Columns.Item("U_Z_EMPAPPROVAL")
                                Try
                                    ousertable.UserFields.Fields.Item("U_Z_EmpApproval").Value = oComboColumn.GetSelectedValue(intRowId).Value
                                    strEmpConfirmaion = oComboColumn.GetSelectedValue(intRowId).Value
                                Catch ex As Exception
                                    ousertable.UserFields.Fields.Item("U_Z_EmpApproval").Value = "P"
                                    strEmpConfirmaion = "P"
                                End Try
                                oComboColumn = oGrid.Columns.Item("U_Z_APPROVED")
                                Dim otest As SAPbobsCOM.Recordset
                                otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                otest.DoQuery("SELECT T0.[U_Z_PRJCODE], T0.[U_Z_APPROVAL],* FROM [dbo].[@Z_HPRJ]  T0 where T0.U_Z_PRJCODE='" & strPrjCode & "'")
                                If otest.Fields.Item("U_Z_APPROVAL").Value = "N" Then
                                    If strEmpConfirmaion = "C" Then
                                        ousertable.UserFields.Fields.Item("U_Z_Approved").Value = "A" ' oCombobox.Selected.Value
                                    Else
                                        ousertable.UserFields.Fields.Item("U_Z_Approved").Value = "P" ' oCombobox.Selected.Value
                                    End If
                                Else
                                    Try
                                        ousertable.UserFields.Fields.Item("U_Z_Approved").Value = oComboColumn.GetSelectedValue(intRowId).Value
                                    Catch ex As Exception
                                        ousertable.UserFields.Fields.Item("U_Z_Approved").Value = "P"
                                    End Try
                                End If
                                ousertable.UserFields.Fields.Item("U_Z_RefCode").Value = strDocref
                                ousertable.UserFields.Fields.Item("U_Z_Remarks").Value = oGrid.DataTable.GetValue("U_Z_REMARKS", intRowId)
                                If ousertable.Update <> 0 Then
                                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Return False
                                Else
                                    Dim strCode1, strQuery1 As String
                                    Dim oTest12 As SAPbobsCOM.Recordset
                                    oTest12 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    strCode1 = strDocref
                                    strQuery1 = "Select *  from [@Z_OTIM]  T0 inner Join [@Z_TIM1] T1  on T1.U_Z_REFCODE=T0.Code  where T0.Code='" & strCode1 & "'"
                                    oTest12.DoQuery(strQuery1)
                                    If oTest12.RecordCount > 0 Then
                                        '           oApplication.Utilities.ClosingProjectActivity(oTest12.Fields.Item("U_Z_EMPCODE").Value, oTest12.Fields.Item("U_Z_PRJCODE").Value, oTest12.Fields.Item("U_Z_PRCNAME").Value, oTest12.Fields.Item("U_Z_ACTNAME").Value, oTest12.Fields.Item("U_Z_DATE").Value)
                                    End If

                                End If
                            End If
                        End If
                    End If
                End If
            Next
            oTempRec.DoQuery("Delete from [@Z_TIM1] where Name like '%_XD'") ' and U_Z_RefCode='" & strDocref & "'")
        End If
        oApplication.Utilities.Message("Operation completed successfuly", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Return True
    End Function

#End Region
#Region "FormDataEvent"


#End Region

#Region "Class Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_ContrctTimeSheet
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If

                Case mnu_ADD_ROW
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        addRow(oForm)
                    End If
                Case mnu_Duplicate_Row
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    'If pVal.BeforeAction = True Then
                    '    DuplicateRow(oForm)
                    '    BubbleEvent = False
                    '    Exit Sub
                    'End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        ' RefereshDeleteRow(oForm)
                    Else
                        If oForm.PaneLevel = 2 Then
                            deleterow(oForm)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If

            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "getBOQReference"
    Private Function getBOQReference(ByVal aItemCode As String, ByVal aProject As String, ByVal aProcess As String, ByVal aActivity As String) As String
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("Select isnull(U_Z_BOQREF,'') from [@Z_PRJ2] where U_Z_ItemCode='" & aItemCode & "' and  U_Z_PRJCODE='" & aProject.Replace("'", "''") & "' and U_Z_MODNAME='" & aProcess.Replace("'", "''") & "' and U_Z_ACTNAME='" & aActivity.Replace("'", "''") & "'")
        Return oTest.Fields.Item(0).Value
    End Function
#End Region

    Private Function CopyBOM(ByVal aform As SAPbouiCOM.Form) As Boolean
        frmSourceForm = oApplication.SBO_Application.Forms.Item(SourceFormUID)

        Try
            frmSourceForm.Freeze(True)

            frmSourceGrid = frmSourceForm.Items.Item(ItemUID).Specific
            oGrid = aform.Items.Item("7").Specific
            For intLoop As Integer = 0 To oGrid.Rows.Count - 1
                oChecbox = oGrid.Columns.Item("Select")
                If oChecbox.IsChecked(intLoop) = True Then
                    If frmSourceGrid.DataTable.GetValue("U_Z_ITEMCODE", frmSourceGrid.DataTable.Rows.Count - 1) <> "" Then
                        frmSourceGrid.DataTable.Rows.Add()
                    End If
                    frmSourceGrid.DataTable.SetValue("U_Z_ITEMCODE", frmSourceGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("ItemCode", intLoop))
                    frmSourceGrid.DataTable.SetValue("U_Z_ITEMNAME", frmSourceGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("ItemName", intLoop))
                    frmSourceGrid.DataTable.SetValue("U_Z_UOM", frmSourceGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("UoM", intLoop))
                    frmSourceGrid.DataTable.SetValue("U_Z_VENDOR", frmSourceGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("CardCode", intLoop))
                    frmSourceGrid.DataTable.SetValue("U_Z_VENDORNAME", frmSourceGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("CardName", intLoop))
                    frmSourceGrid.DataTable.SetValue("U_Z_REQQTY", frmSourceGrid.DataTable.Rows.Count - 1, oGrid.DataTable.GetValue("Quantity", intLoop))
                    Dim dblPrice As Double = oApplication.Utilities.getPriceforSupplier(oGrid.DataTable.GetValue("CardCode", intLoop), oGrid.DataTable.GetValue("ItemCode", intLoop))
                    Dim dblQty As Double = oGrid.DataTable.GetValue("Quantity", intLoop)

                    Try
                        frmSourceGrid.DataTable.SetValue("U_Z_UnitPrice", frmSourceGrid.DataTable.Rows.Count - 1, dblPrice)
                    Catch ex As Exception
                        frmSourceGrid.DataTable.SetValue("U_Z_UNITPRICE", frmSourceGrid.DataTable.Rows.Count - 1, dblPrice)
                    End Try

                    Try
                        frmSourceGrid.DataTable.SetValue("U_Z_ESTCOST", frmSourceGrid.DataTable.Rows.Count - 1, dblPrice * dblQty)
                    Catch ex As Exception
                        frmSourceGrid.DataTable.SetValue("U_Z_ESTCOST", frmSourceGrid.DataTable.Rows.Count - 1, dblPrice * dblQty)
                    End Try
                End If
            Next
            AssignLineNo1(frmSourceForm, frmSourceGrid)
            frmSourceForm.Freeze(False)
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            frmSourceForm.Freeze(False)
            Return False
        End Try
    End Function
    Private Sub AssignLineNo(ByVal aform As SAPbouiCOM.Form, ByVal agrid As SAPbouiCOM.Grid)
        Try
            aform.Freeze(True)
            For count As Integer = 0 To agrid.DataTable.Rows.Count - 1
                agrid.RowHeaders.SetText(count, count + 1)
            Next
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

    Private Sub SelectAll(ByVal aform As SAPbouiCOM.Form, ByVal achoice As Boolean)
        aform.Freeze(True)
        oGrid = aform.Items.Item("7").Specific
        oChecbox = oGrid.Columns.Item("Select")
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oChecbox.Check(intRow, achoice)
        Next
        aform.Freeze(False)
    End Sub
#Region "Item Event"

    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Select Case pVal.BeforeAction
            Case True
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                        If pVal.ItemUID = "21" Then
                            If oForm.PaneLevel <> 1 Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED
                        oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                        If pVal.ItemUID = "21" Then
                            If oForm.PaneLevel <> 1 Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                End Select
           
            Case False
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                        If pVal.ItemUID = "11" And pVal.ColUID = "U_Z_PRJCODE" And pVal.CharPressed = 9 Then
                            Dim objChoose As New clsContractTimeCFL
                            Dim strwhs, strGridValue As String
                            oGrid = oForm.Items.Item("11").Specific
                            strwhs = oGrid.DataTable.GetValue(pVal.ColUID, pVal.Row)
                            strGridValue = oApplication.Utilities.getEdittextvalue(oForm, "22")
                            If strwhs <> "" Then
                                If oApplication.Utilities.CheckProject_Contract(strwhs, strGridValue) = True Then
                                    Exit Sub
                                Else
                                    Dim strProject, strActivtiy, strBusiness, strsql As String
                                    Dim oTe As SAPbobsCOM.Recordset
                                    oTe = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, "")
                                    oGrid.DataTable.SetValue("U_Z_PRJNAME", pVal.Row, "")
                                End If
                            Else
                                oGrid.DataTable.SetValue("U_Z_PRJNAME", pVal.Row, "")
                            End If
                            If strGridValue <> "" Then
                                clsContractTimeCFL.ItemUID = pVal.ItemUID
                                clsContractTimeCFL.SourceFormUID = FormUID
                                clsContractTimeCFL.SourceLabel = pVal.Row
                                clsContractTimeCFL.CFLChoice = "[@Z_PRJ1]" 'oCombo.Selected.Value
                                clsContractTimeCFL.choice = "Project"
                                clsContractTimeCFL.ContractID = strGridValue
                                clsContractTimeCFL.sourceItemCode = ""
                                clsContractTimeCFL.Documentchoice = "" ' oApplication.Utilities.GetDocType(oForm)
                                clsContractTimeCFL.sourceColumID = pVal.ColUID
                                clsContractTimeCFL.sourcerowId = pVal.Row
                                clsContractTimeCFL.BinDescrUID = "" ' oApplication.Utilities.getEdittextvalue(oForm, "4")
                                oApplication.Utilities.LoadForm("Contract_CFL.xml", frm_ContractCFL)
                                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                objChoose.databound(oForm)
                            End If
                        End If

                        If pVal.ItemUID = "11" And (pVal.ColUID = "U_Z_PRCNAME") And pVal.CharPressed = 9 Then

                            Dim objChoose As New clsContractTimeCFL
                            Dim strwhs, strGridValue, strProject As String
                            oGrid = oForm.Items.Item("11").Specific
                            strwhs = oGrid.DataTable.GetValue(pVal.ColUID, pVal.Row)
                            strGridValue = oApplication.Utilities.getEdittextvalue(oForm, "22")
                            strProject = oGrid.DataTable.GetValue("U_Z_PRJCODE", pVal.Row)
                            If strProject = "" Then
                                Exit Sub
                            End If
                            If strwhs <> "" Then
                                If oApplication.Utilities.CheckModule_Contract(strProject, strGridValue, strwhs) = True Then
                                    Exit Sub
                                Else
                                    Dim strActivtiy, strBusiness, strsql As String
                                    Dim oTe As SAPbobsCOM.Recordset
                                    oTe = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, "")
                                    oGrid.DataTable.SetValue("U_Z_ACTNAME", pVal.Row, "")
                                End If
                            Else
                                oGrid.DataTable.SetValue("U_Z_ACTNAME", pVal.Row, "")
                            End If
                            strwhs = oGrid.DataTable.GetValue(pVal.ColUID, pVal.Row)
                            If strwhs = "" And strGridValue <> "" Then
                                clsContractTimeCFL.ItemUID = pVal.ItemUID
                                clsContractTimeCFL.SourceFormUID = FormUID
                                clsContractTimeCFL.SourceLabel = pVal.Row
                                clsContractTimeCFL.CFLChoice = strProject  'oCombo.Selected.Value
                                clsContractTimeCFL.choice = "ACTIVITY"
                                clsContractTimeCFL.ContractID = strGridValue
                                clsContractTimeCFL.sourceItemCode = ""
                                clsContractTimeCFL.Documentchoice = "" ' oApplication.Utilities.GetDocType(oForm)
                                clsContractTimeCFL.sourceColumID = pVal.ColUID
                                clsContractTimeCFL.sourcerowId = pVal.Row
                                clsContractTimeCFL.BinDescrUID = "" ' oApplication.Utilities.getEdittextvalue(oForm, "4")
                                oApplication.Utilities.LoadForm("Contract_CFL.xml", frm_ContractCFL)
                                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                objChoose.databound(oForm)
                            End If
                        End If

                        If pVal.ItemUID = "11" And (pVal.ColUID = "U_Z_EMPCODE") And pVal.CharPressed = 9 Then

                            Dim objChoose As New clsContractTimeCFL
                            Dim strwhs, strGridValue, strProject As String
                            oGrid = oForm.Items.Item("11").Specific
                            strwhs = oGrid.DataTable.GetValue(pVal.ColUID, pVal.Row)
                            strGridValue = oApplication.Utilities.getEdittextvalue(oForm, "22")
                            strProject = oGrid.DataTable.GetValue("U_Z_EMPCODE", pVal.Row)
                            If strProject <> "" Then
                                If oApplication.Utilities.CheckEMP_Contract(strProject, strGridValue) = True Then
                                    '  Exit Sub
                                Else
                                    Dim strActivtiy, strBusiness, strsql As String
                                    Dim oTe As SAPbobsCOM.Recordset
                                    oTe = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, "")
                                    oGrid.DataTable.SetValue("U_Z_EMPNAME", pVal.Row, "")
                                End If
                            Else
                                oGrid.DataTable.SetValue("U_Z_EMPNAME", pVal.Row, "")
                            End If
                            strProject = oGrid.DataTable.GetValue(pVal.ColUID, pVal.Row)
                            If strProject = "" And strGridValue <> "" Then
                                clsContractTimeCFL.ItemUID = pVal.ItemUID
                                clsContractTimeCFL.SourceFormUID = FormUID
                                clsContractTimeCFL.SourceLabel = pVal.Row
                                clsContractTimeCFL.CFLChoice = strProject  'oCombo.Selected.Value
                                clsContractTimeCFL.choice = "EMP"
                                clsContractTimeCFL.ContractID = strGridValue
                                clsContractTimeCFL.sourceItemCode = ""
                                clsContractTimeCFL.Documentchoice = "" ' oApplication.Utilities.GetDocType(oForm)
                                clsContractTimeCFL.sourceColumID = pVal.ColUID
                                clsContractTimeCFL.sourcerowId = pVal.Row
                                clsContractTimeCFL.BinDescrUID = "" ' oApplication.Utilities.getEdittextvalue(oForm, "4")
                                oApplication.Utilities.LoadForm("Contract_CFL.xml", frm_ContractCFL)
                                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                objChoose.databound(oForm)
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        oform = oApplication.SBO_Application.Forms.Item(FormUID)
                        If pVal.ItemUID = "8" Then
                            ' SelectAll(oform, True)
                        End If
                        If pVal.ItemUID = "9" Then
                            ' SelectAll(oform, False)
                        End If
                        If pVal.ItemUID = "4" Then
                            If oForm.PaneLevel = 1 Then
                                If oApplication.Utilities.getEdittextvalue(oForm, "21") = "" Then
                                    oApplication.Utilities.Message("Sub Contractor Details is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                End If
                            End If
                            oForm.PaneLevel = oForm.PaneLevel + 1
                            databind(oForm)
                        End If
                        If pVal.ItemUID = "3" Then
                            If oForm.PaneLevel = 2 Or oForm.PaneLevel = 3 Then
                                oForm.PaneLevel = 1
                            End If
                            '  databind(oForm)
                        End If
                        If pVal.ItemUID = "5" Then
                            If validation(oForm) = False Then
                                Exit Sub
                            End If
                            If AddToUDT_Table(oForm) = True Then
                                databind(oForm)
                            End If
                        End If
                        If pVal.ItemUID = "7" Then
                            oForm.PaneLevel = 2
                        End If
                        If pVal.ItemUID = "8" Then
                            oForm.PaneLevel = 3
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        Dim val1 As String
                        Dim sCHFL_ID, val As String
                        Dim intChoice As Integer
                        Dim codebar As String
                        Try
                            oCFLEvento = pVal
                            sCHFL_ID = oCFLEvento.ChooseFromListUID
                            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                            If (oCFLEvento.BeforeAction = False) Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                oDataTable = oCFLEvento.SelectedObjects
                                intChoice = 0
                                oForm.Freeze(True)
                                If pVal.ItemUID = "21" Then
                                    val1 = oDataTable.GetValue("U_Z_CARDCODE", 0)
                                    val = oDataTable.GetValue("DocEntry", 0)
                                    Try
                                        oApplication.Utilities.setEdittextvalue(oForm, "22", val)
                                        oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val1)
                                    Catch ex As Exception
                                    End Try
                                End If
                                oForm.Freeze(False)
                            End If
                        Catch ex As Exception
                          
                            oForm.Freeze(False)
                        End Try

                End Select

        End Select
    End Sub
#End Region

End Class
