Public Class clsprjReports
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oCombobox As SAPbouiCOM.ComboBoxColumn
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oCheckBoxColumn As SAPbouiCOM.CheckBoxColumn
    Private obutton As SAPbouiCOM.Button
    Private oRecset As SAPbobsCOM.Recordset
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oitems As SAPbouiCOM.Item
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

    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_PrjReports) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_PrjReports, frm_PrjReports)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        AddChooseFromList(oForm)
        oForm.DataSources.UserDataSources.Add("empFrom", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("empTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("dtFrom", SAPbouiCOM.BoDataType.dt_DATE)
        oForm.DataSources.UserDataSources.Add("dtTo", SAPbouiCOM.BoDataType.dt_DATE)
        oForm.DataSources.UserDataSources.Add("prjFrom", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("prjto", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oApplication.Utilities.setUserDatabind(oForm, "9", "empfrom")
        oEditText = oForm.Items.Item("9").Specific
        oEditText.ChooseFromListUID = "CFL1"
        oEditText.ChooseFromListAlias = "empID"
        oApplication.Utilities.setUserDatabind(oForm, "11", "empTo")
        oEditText = oForm.Items.Item("11").Specific
        oEditText.ChooseFromListUID = "CFL2"
        oEditText.ChooseFromListAlias = "empID"
        oApplication.Utilities.setUserDatabind(oForm, "13", "dtfrom")
        oApplication.Utilities.setUserDatabind(oForm, "15", "dtTo")

        oApplication.Utilities.setUserDatabind(oForm, "17", "prjFrom")
        oEditText = oForm.Items.Item("17").Specific



        oEditText.ChooseFromListUID = "CFL3"
        oEditText.ChooseFromListAlias = "U_Z_PRJCODE"



        'databind(oForm)
        oCombo = oForm.Items.Item("7").Specific
        Dim oCombo1 As SAPbouiCOM.ComboBox
        oCombo1 = oForm.Items.Item("7").Specific
        For intRow As Integer = oCombo1.ValidValues.Count - 1 To 0 Step -1
            oCombo1.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        oCombo1.ValidValues.Add("", "")
        oCombo1.ValidValues.Add("TA", "Time and Attendance")
        oCombo1.ValidValues.Add("PRS", "Project Status ")
        oCombo1.ValidValues.Add("PRJS", "Project Summary ")
        oCombo1.ValidValues.Add("RS", "Resource Utilization ")
        oCombo1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        oCombo1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oForm.PaneLevel = 1
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


            oCFL = oCFLs.Item("CFL3")
            Dim intEmp As String = oApplication.Utilities.getLoggedonEmployee()
            If oApplication.Utilities.CheckSuperUser(intEmp) = False Then
                oCons = oCFL.GetConditions()
                oCon = oCons.Add()
                oCon.Alias = "U_Z_EMPID"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = intEmp
                oCFL.SetConditions(oCons)
                oCon = oCons.Add()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub LoadForm(ByVal aEmpId As String, ByVal aChoice As String)
        oForm = oApplication.Utilities.LoadForm(xml_Report, frm_Report)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataSources.UserDataSources.Add("empFrom", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("empTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("dtFrom", SAPbouiCOM.BoDataType.dt_DATE)
        oForm.DataSources.UserDataSources.Add("dtTo", SAPbouiCOM.BoDataType.dt_DATE)
        oForm.DataSources.UserDataSources.Add("prjFrom", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("prjto", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("ActCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

        oApplication.Utilities.setUserDatabind(oForm, "9", "empfrom")
        oEditText = oForm.Items.Item("9").Specific
        oEditText.ChooseFromListUID = "CFL4"
        oEditText.ChooseFromListAlias = "empID"
        oEditText.String = aEmpId
        oApplication.Utilities.setUserDatabind(oForm, "12", "empTo")
        oEditText = oForm.Items.Item("12").Specific
        oEditText.ChooseFromListUID = "CFL5"
        oEditText.ChooseFromListAlias = "empID"
        oEditText.String = aEmpId
        oApplication.Utilities.setUserDatabind(oForm, "18", "dtfrom")
        oApplication.Utilities.setUserDatabind(oForm, "20", "dtTo")
        oApplication.Utilities.setUserDatabind(oForm, "102", "ActCode")
        oEditText = oForm.Items.Item("102").Specific
        oEditText.ChooseFromListUID = "CFL_7"
        oEditText.ChooseFromListAlias = "FormatCode"
        'databind(oForm)
        oCombo = oForm.Items.Item("4").Specific
        oCombo.ValidValues.Add("", "")
        oCombo.ValidValues.Add("Exp", "Expenses")
        oCombo.ValidValues.Add("TIME", "Time Sheet")
        oCombo.ValidValues.Add("LEV", "Leave Request")
        oCombo.ValidValues.Add("TA", "Employee Time & Attendance")
        oCombo.ValidValues.Add("Prj", "Project Wise Time Sheet")
        oCombo.ValidValues.Add("Prjs", "Project Summery")
        oCombo.ValidValues.Add("PrjM", "Project Material Status")

        oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oCombo = oForm.Items.Item("36").Specific
        oCombo.ValidValues.Add("", "")
        oCombo.ValidValues.Add("B", "Business Partner")
        oCombo.ValidValues.Add("E", "Expenses")
        oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oForm.Items.Item("36").DisplayDesc = True

        oForm.Items.Item("4").DisplayDesc = True
        oForm.Items.Item("6").DisplayDesc = True
        oForm.Items.Item("15").DisplayDesc = True
        oForm.Items.Item("7").Visible = False
        ' oForm.Items.Item("17").Visible = False
        fillCombo(oForm)
        If EntryChoice = "Reports" Then
            If aChoice = "Super" Then
                oForm.Items.Item("9").Enabled = True
                oForm.Items.Item("12").Enabled = True
            Else
                oForm.Items.Item("9").Enabled = False
                oForm.Items.Item("12").Enabled = False
            End If
            oForm.Items.Item("101").Visible = False
            oForm.Items.Item("102").Visible = False
            oForm.Items.Item("103").Visible = False
            oForm.Items.Item("104").Visible = False
            oForm.Items.Item("3").Visible = True
            oForm.Items.Item("4").Visible = True
            oForm.Items.Item("115").Visible = False
            oForm.Items.Item("36").Visible = False
        ElseIf EntryChoice = "Posting" Then
            If aChoice = "Super" Then
                oForm.Items.Item("9").Enabled = True
                oForm.Items.Item("12").Enabled = True
            Else
                oForm.Items.Item("9").Enabled = False
                oForm.Items.Item("12").Enabled = False
            End If
            oForm.Title = "Expense Accounting "
            oForm.Items.Item("101").Visible = True
            oForm.Items.Item("102").Visible = True
            oForm.Items.Item("103").Visible = True
            oForm.Items.Item("104").Visible = True
            oForm.Items.Item("3").Visible = False
            oForm.Items.Item("4").Visible = False
            oForm.Items.Item("115").Visible = True
            oForm.Items.Item("36").Visible = True
        End If

        oForm.PaneLevel = 0
        oForm.Freeze(False)
    End Sub

    Private Sub fillCombo(ByVal aForm As SAPbouiCOM.Form)
        Dim oCombo1 As SAPbouiCOM.ComboBox
        Dim strString, strCaption As String
        Dim otemp As SAPbobsCOM.Recordset
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strString = "Select U_Z_PrjCode,U_Z_PrjName from [@Z_HPRJ] order by DocEntry"
        oCombo1 = aForm.Items.Item("6").Specific
        For intRow As Integer = oCombo1.ValidValues.Count - 1 To 0 Step -1
            oCombo1.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
        Next

        otemp.DoQuery(strString)
        oCombo1.ValidValues.Add("", "")
        For intRow As Integer = 0 To otemp.RecordCount - 1
            oCombo1.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
            otemp.MoveNext()
        Next
        oCombo1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        aForm.Items.Item("6").DisplayDesc = True

        oCombo1 = aForm.Items.Item("15").Specific
        For intRow As Integer = oCombo1.ValidValues.Count - 1 To 0 Step -1
            oCombo1.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
        Next

        otemp.DoQuery(strString)
        oCombo1.ValidValues.Add("", "")
        For intRow As Integer = 0 To otemp.RecordCount - 1
            oCombo1.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
            otemp.MoveNext()
        Next
        oCombo1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        aForm.Items.Item("15").DisplayDesc = True
    End Sub

#Region "Databind Summary"
    Private Sub DatabindSummary(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)

            Dim strString1, strString2, strreporttype, strCode, strString, strFromEMP, strToEMP, strFromPRJ, strToPRJ, strfromdate, strToDate, strEMPCondition, strDateCondition, strProjectCondition As String
            Dim dtFromdate, dtTodate As Date

            strFromEMP = oApplication.Utilities.getEdittextvalue(aForm, "9")
            strToEMP = oApplication.Utilities.getEdittextvalue(aForm, "12")

            strfromdate = oApplication.Utilities.getEdittextvalue(aForm, "18")
            strToDate = oApplication.Utilities.getEdittextvalue(aForm, "20")

            If strfromdate <> "" Then
                dtFromdate = oApplication.Utilities.GetDateTimeValue(strfromdate)
            End If
            If strToDate <> "" Then
                dtTodate = oApplication.Utilities.GetDateTimeValue(strToDate)
            End If
            oCombo = aForm.Items.Item("4").Specific
            strreporttype = oCombo.Selected.Value

            oCombo = aForm.Items.Item("6").Specific
            strFromPRJ = oCombo.Selected.Value
            oCombo = aForm.Items.Item("15").Specific
            strToPRJ = oCombo.Selected.Value

            strEMPCondition = ""
            strDateCondition = ""
            strProjectCondition = ""
            If strFromEMP <> "" And strToEMP <> "" Then
                strEMPCondition = " U_Z_EMPCODE between '" & strFromEMP & "' and '" & strToEMP & "'"
            ElseIf strFromEMP <> "" And strToEMP = "" Then
                strEMPCondition = " U_Z_EMPCODE >= '" & strFromEMP & "'"
            ElseIf strFromEMP = "" And strToEMP <> "" Then
                strEMPCondition = " U_Z_EMPCODE <= '" & strToEMP & "'"
            Else
                strEMPCondition = " 1=1"
            End If

            If strFromPRJ <> "" And strToPRJ <> "" Then
                strProjectCondition = " U_Z_PRJCODE between '" & strFromPRJ & "' and '" & strToPRJ & "'"
            ElseIf strFromPRJ <> "" And strToPRJ = "" Then
                strProjectCondition = " U_Z_PRJCODE >= '" & strFromPRJ & "'"
            ElseIf strFromPRJ = "" And strToPRJ <> "" Then
                strProjectCondition = " U_Z_PRJCODE <= '" & strToPRJ & "'"
            Else
                strProjectCondition = " 1=1"
            End If


            Dim strDateCondition1 As String = ""

            If strfromdate <> "" And strToDate <> "" Then
                strDateCondition = " U_Z_DocDate > = '" & dtFromdate.ToString("yyyy-MM-dd") & "' and U_Z_DocDate <= '" & dtTodate.ToString("yyyy-MM-dd") & "'"
                strDateCondition1 = " U_Z_DocDate > = '" & dtFromdate.ToString("yyyy-dd-MM") & "' and U_Z_DocDate <= '" & dtTodate.ToString("yyyy-dd-MM") & "'"
            ElseIf strfromdate <> "" And strToDate = "" Then
                strDateCondition = " U_Z_DocDate >= '" & dtFromdate.ToString("yyyy-MM-dd") & "'"
                strDateCondition1 = " U_Z_DocDate >= '" & dtFromdate.ToString("yyyy-dd-MM") & "'"
            ElseIf strfromdate = "" And strToDate <> "" Then
                strDateCondition = " U_Z_DocDate <= '" & dtTodate.ToString("yyyy-MM-dd") & "'"
                strDateCondition1 = " U_Z_DocDate <= '" & dtTodate.ToString("yyyy-dd-MM") & "'"
            Else
                strDateCondition = " 1=1"
                strDateCondition1 = " 1=1"
            End If

            Dim oTestRs As SAPbobsCOM.Recordset
            Dim s As String
            Try
                s = "Select * from [@Z_OEXP] where U_Z_DocDate > ='2011-01-04' and U_Z_DocDate  <='2011-30-04'"
                oTestRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTestRs.DoQuery(s)
                strDateCondition = strDateCondition1
            Catch ex As Exception
                s = "Select * from [@Z_OTIM] where " & strDateCondition
                oTestRs.DoQuery(s)
            End Try

            Dim strCondition, strGroupcondition, st1, strActvityquery As String
            strCondition = strEMPCondition & " and " & strProjectCondition & " and " & strDateCondition
            strString = "SELECT T0.[U_Z_PRJCODE],T3.PrjName ,T0.[U_Z_PRCNAME],sum(T0.U_Z_HOURS) FROM [dbo].[@Z_TIM1]  T0 inner join [@Z_OTIM] T1 on T0.U_Z_REFCODE=T1.CODE "
            strString = strString & " inner join OPRJ T3 on T3.PrjCode=T0.U_Z_PrjCode where " & strCondition & " group by T0.[U_Z_PRJCODE],T3.PrjName,T0.[U_Z_PRCNAME]  order by T0.[U_Z_PRJCODE],T0.[U_Z_PRCNAME]"
            strActvityquery = "SELECT T0.[U_Z_PRJCODE],T0.[U_Z_PRCNAME],T0.[U_Z_ACTNAME],isnull(sum(T0.U_Z_HOURS),0) FROM [dbo].[@Z_TIM1]  T0  inner join [@Z_OTIM] T1 on T0.U_Z_REFCODE=T1.CODE and T0.[U_Z_APPROVED]<>'P' where " & strCondition
            'strActvityquery = "SELECT T0.[U_Z_PRJCODE],T0.[U_Z_PRCNAME],T0.[U_Z_ACTNAME],isnull(sum(U_Z_HOURS),0) FROM [dbo].[@Z_TIM1]  T0  inner join [@Z_OTIM] T1 on T0.U_Z_REFCODE=T1.CODE  where " & strCondition

            Dim oRS, otemp As SAPbobsCOM.Recordset
            Dim strColun, stTempQuery, strModule As String
            oGrid = aForm.Items.Item("26").Specific
            oGrid.DataTable.Rows.Clear()
            oRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery(strString)

            For intRow As Integer = 0 To oRS.RecordCount - 1
                oGrid.DataTable.Rows.Add()
                oGrid.DataTable.SetValue("PrjCode", oGrid.DataTable.Rows.Count - 1, oRS.Fields.Item(0).Value)
                oGrid.DataTable.SetValue("PrjName", oGrid.DataTable.Rows.Count - 1, oRS.Fields.Item(1).Value)
                oGrid.DataTable.SetValue("PrcName", oGrid.DataTable.Rows.Count - 1, oRS.Fields.Item(2).Value)
                strModule = oRS.Fields.Item(2).Value
                otemp.DoQuery("Select isnull(sum(U_Z_Days),0) from [@Z_PRJ1] where U_Z_ModName='" & strModule.Replace("'", "''") & "' and  DocEntry= (Select docEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & oRS.Fields.Item(0).Value & "') group by U_Z_ModName")
                oGrid.DataTable.SetValue("Hours", oGrid.DataTable.Rows.Count - 1, otemp.Fields.Item(0).Value)
                For intLoop As Integer = 4 To oGrid.Columns.Count - 1
                    strColun = oGrid.DataTable.Columns.Item(intLoop).Name
                    stTempQuery = strActvityquery
                    stTempQuery = stTempQuery & " and T0.[U_Z_PRJCODE]='" & oRS.Fields.Item(0).Value & "' and T0.[U_Z_PRCNAME]='" & strModule.Replace("'", "''") & "' and  T0.[U_Z_ACTNAME]='" & strColun.Replace("'", "''") & "' Group by T0.[U_Z_PRJCODE],T0.[U_Z_PRCNAME],T0.[U_Z_ACTNAME]  order by T0.[U_Z_PRJCODE],T0.[U_Z_PRCNAME]"
                    otemp.DoQuery(stTempQuery)
                    oGrid.DataTable.SetValue(strColun, oGrid.DataTable.Rows.Count - 1, otemp.Fields.Item(3).Value)
                Next
                oRS.MoveNext()
            Next
            Dim intColumcount As Integer
            intColumcount = oGrid.Columns.Count - 1
            oGrid.Columns.Item(3).TitleObject.Caption = "Estimated Days"
            oGrid.Columns.Item(0).TitleObject.Caption = "Project Code"
            oGrid.Columns.Item(1).TitleObject.Caption = "Project  Name"
            oGrid.Columns.Item(2).TitleObject.Caption = "Phase"
            oForm.Items.Item("26").Enabled = False
            For intRow As Integer = 2 To oGrid.Columns.Count - 1
                oEditTextColumn = oGrid.Columns.Item(intRow)
                ' oGrid.Columns.Item(intRow).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            Next
            oGrid.CollapseLevel = 1
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        aForm.Freeze(False)
    End Sub
#End Region

#Region "DataBind"
    Private Sub databind(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Dim strString1, strString2, strreporttype, strCode, strString, strFromEMP, strToEMP, strFromPRJ, strToPRJ, strfromdate, strToDate, strEMPCondition, strDateCondition, strProjectCondition As String
            Dim dtFromdate, dtTodate As Date

            strFromEMP = oApplication.Utilities.getEdittextvalue(aForm, "9")
            strToEMP = oApplication.Utilities.getEdittextvalue(aForm, "11")

            strfromdate = oApplication.Utilities.getEdittextvalue(aForm, "13")
            strToDate = oApplication.Utilities.getEdittextvalue(aForm, "15")

            If strfromdate <> "" Then
                dtFromdate = oApplication.Utilities.GetDateTimeValue(strfromdate)
                '   dtFromdate = CDate(strfromdate)
            End If
            If strToDate <> "" Then
                dtTodate = oApplication.Utilities.GetDateTimeValue(strToDate)
            End If
            'oCombo = aForm.Items.Item("4").Specific
            'strreporttype = oCombo.Selected.Value
            oCombo = aForm.Items.Item("7").Specific
            strreporttype = oCombo.Selected.Value
            If strreporttype = "" Then
                oApplication.Utilities.Message("Report type missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Exit Sub
            End If


            strFromPRJ = oApplication.Utilities.getEdittextvalue(aForm, "17")
            Dim strSQL As String
            Dim strExmEmpCondition As String = ""
            strEMPCondition = ""
            strDateCondition = ""
            strProjectCondition = ""
            If strFromEMP <> "" And strToEMP <> "" Then
                strEMPCondition = " Convert(Decimal,U_Z_EMPCODE) between " & CDbl(strFromEMP) & " and " & CDbl(strToEMP)
                strExmEmpCondition = " isnull(T1.U_Z_EMPID,'0') between " & CDbl(strFromEMP) & " and " & CDbl(strToEMP)
            ElseIf strFromEMP <> "" And strToEMP = "" Then
                strEMPCondition = " Convert(Decimal,U_Z_EMPCODE) >= " & CDbl(strFromEMP)
                strExmEmpCondition = "  isnull(T1.U_Z_EMPID,'0') >= " & CDbl(strFromEMP)
            ElseIf strFromEMP = "" And strToEMP <> "" Then
                strEMPCondition = " Convert(Decimal,U_Z_EMPCODE) <= " & CDbl(strToEMP)
                strExmEmpCondition = "  isnull(T1.U_Z_EMPID,'0') <= " & CDbl(strToEMP)
            Else
                strEMPCondition = " 1=1"
                strExmEmpCondition = " 1=1"
            End If
            Dim projectMaterialcond As String
           
            If strFromPRJ = "" Then
                strProjectCondition = " 1=1"
                projectMaterialcond = "1=1"
            Else
                strProjectCondition = "X.U_Z_PRJCODE='" & strFromPRJ & "'"
                projectMaterialcond = "U_Z_PRJCODE='" & strFromPRJ & "'"
            End If

            Dim strDateCondition1 As String = ""
            Dim strDateCondition2 As String = ""
            Dim strexpquery As String = ", 0/1.0 'Estimated Expenses' , 0/1.0 'Acutal Approved Expenses', 0/1.0 'Variance in Expenses', 0/1.0 'Pening for Approval Expenses', x.EstimatedCost-x.EstimatedCost 'Total Project Cost',x.EstimatedCost-x.EstimatedCost 'Total Actual Cost',x.EstimatedCost-x.EstimatedCost 'Total Variance Cost',"

            If strfromdate <> "" And strToDate <> "" Then
                If strfromdate = strToDate Then
                    strDateCondition = " U_Z_DocDate > = '" & dtFromdate.ToString("yyyy-MM-dd") & "' and U_Z_DocDate <= '" & dtTodate.ToString("yyyy-MM-dd") & "'"
                    strDateCondition1 = "('" & dtFromdate.ToString("yyyy-MM-dd") & "' between T1.U_Z_FromDate and T1.U_Z_ToDate) " ' (T1.U_Z_FromDate > = '" & dtFromdate.ToString("yyyy-MM-dd") & "' and T1.U_Z_FromDate <= '" & dtTodate.ToString("yyyy-MM-dd") & "')"
                    ' strDateCondition1 = " U_Z_DocDate > = '" & dtFromdate.ToString("MM-dd-yyyy") & "' and U_Z_DocDate <= '" & dtTodate.ToString("MM-dd-yyyy") & "'"
                    strDateCondition2 = "('" & dtFromdate.ToString("yyyy-MM-dd") & "' between T1.U_Z_FromDate and T1.U_Z_ToDate) " ' (T1.U_Z_FromDate > = '" & dtFromdate.ToString("yyyy-MM-dd") & "' and T1.U_Z_FromDate <= '" & dtTodate.ToString("yyyy-MM-dd") & "')"

                Else
                    strDateCondition = " U_Z_DocDate > = '" & dtFromdate.ToString("yyyy-MM-dd") & "' and U_Z_DocDate <= '" & dtTodate.ToString("yyyy-MM-dd") & "'"
                    strDateCondition1 = " (T1.U_Z_FromDate > = '" & dtFromdate.ToString("yyyy-MM-dd") & "' and T1.U_Z_FromDate <= '" & dtTodate.ToString("yyyy-MM-dd") & "')"
                    ' strDateCondition1 = " U_Z_DocDate > = '" & dtFromdate.ToString("MM-dd-yyyy") & "' and U_Z_DocDate <= '" & dtTodate.ToString("MM-dd-yyyy") & "'"
                    strDateCondition2 = " (T1.U_Z_ToDate > = '" & dtFromdate.ToString("yyyy-MM-dd") & "' and T1.U_Z_ToDate <= '" & dtTodate.ToString("yyyy-MM-dd") & "')"

                End If
               
            ElseIf strfromdate <> "" And strToDate = "" Then
                strDateCondition = " U_Z_DocDate >= '" & dtFromdate.ToString("yyyy-MM-dd") & "'"
                strDateCondition1 = "(T1.U_Z_FromDate >= '" & dtFromdate.ToString("yyyy-MM-dd") & "')"
                strDateCondition2 = "(T1.U_Z_ToDate >= '" & dtFromdate.ToString("yyyy-MM-dd") & "')"

            ElseIf strfromdate = "" And strToDate <> "" Then
                strDateCondition = " U_Z_DocDate <= '" & dtTodate.ToString("yyyy-MM-dd") & "'"
                strDateCondition1 = "(T1.U_Z_FromDate <= '" & dtTodate.ToString("yyyy-MM-dd") & "')"
                strDateCondition2 = "(T1.U_Z_ToDate <= '" & dtFromdate.ToString("yyyy-MM-dd") & "')"

            Else
                strDateCondition = " 1=1"
                strDateCondition1 = " 1=1"
                strDateCondition2 = " 1=1"
            End If
            Dim oCheckbox As SAPbouiCOM.CheckBox
            oCheckbox = aForm.Items.Item("21").Specific
            Dim strExcludeCondition As String
            If oCheckbox.Checked = True Then
                strExcludeCondition = "(Select PrjCode from OPRJ where isnull(U_Z_INTERNAL,'N')='N')"
            Else
                strExcludeCondition = "(Select PrjCode from OPRJ where isnull(U_Z_INTERNAL,'N')<>'X')"
            End If

            Dim st5 As String
            Dim blnSuperUser As Boolean
            Dim intUser As String = oApplication.Utilities.getLoggedonEmployee()
            If oApplication.Utilities.CheckSuperUser(intUser) = False Then
                blnSuperUser = False
            Else
                blnSuperUser = True
            End If
            Dim strEMPIDS As String = oApplication.Utilities.getEmpIDforMangers_Reports(oApplication.Company.UserName)
            Select Case strreporttype
                Case "TA" 'Time and Attendance
                    'strSQL = "SELECT  T0.[U_Z_DOCDATE],T0.[U_Z_EMPCODE] AS 'EMPLOYEE ID', T0.[U_Z_EMPNAME],  [U_Z_TYPE]  as 'Project Code/ Leave Type' ,' ' 'U_Z_PRJNAME' ,' '   AS 'PHASE' ,' '  AS 'ACTIVITY',  T0.[U_Z_DAYS] 'Days/Hours' FROM [@Z_OLEV]  T0  where " & strEMPCondition & "   and    ( " & strDateCondition & ")"
                    'strSQL = strSQL & "  union all"
                    'strSQL = strSQL & " SELECT T1.[U_Z_DATE],T0.[U_Z_EMPCODE] AS 'EMPLOYEE ID', T0.[U_Z_EMPNAME],  T1.[U_Z_PRJCODE] as 'Project Code/ Leave Type',T1.[U_Z_PRJNAME], T1.[U_Z_PRCNAME]  AS 'PHASE', T1.[U_Z_ACTNAME] AS 'ACTIVITY',T1.[U_Z_HOURS] 'Days/Hours'  FROM [@Z_OTIM]  T0  inner join  [@Z_TIM1]  T1 on T0.Code=T1.U_Z_REFCODE   where " & strEMPCondition & "   and    ( " & strDateCondition & ")"

                    If blnSuperUser = True Then
                        strSQL = "SELECT  T0.[U_Z_DOCDATE],T0.[U_Z_EMPCODE] AS 'EMPLOYEE ID', T0.[U_Z_EMPNAME],  [U_Z_TYPE]  as 'Project Code/ Leave Type' ,' ' 'U_Z_PRJNAME' ,' '   AS 'PHASE' ,' '  AS 'ACTIVITY', T0.[U_Z_DAYS] 'Days', T0.[U_Z_DAYS] * 8.0 'Hours' FROM [@Z_OLEV]  T0  where " & strEMPCondition & "   and    ( " & strDateCondition & ")"
                        strSQL = strSQL & "  union all"
                        strSQL = strSQL & " SELECT T1.[U_Z_DATE],T0.[U_Z_EMPCODE] AS 'EMPLOYEE ID', T0.[U_Z_EMPNAME],  T1.[U_Z_PRJCODE] as 'Project Code/ Leave Type',T1.[U_Z_PRJNAME], T1.[U_Z_PRCNAME]  AS 'PHASE', T1.[U_Z_ACTNAME] AS 'ACTIVITY',T1.[U_Z_HOURS] /8.0 'Days',T1.[U_Z_HOURS] 'Hours'  FROM [@Z_OTIM]  T0  inner join  [@Z_TIM1]  T1 on T0.Code=T1.U_Z_REFCODE   where T1.U_Z_PRJCODE In (" & strExcludeCondition & ") and  " & strEMPCondition & "   and    ( " & strDateCondition & ")"

                    Else
                        strSQL = "SELECT  T0.[U_Z_DOCDATE],T0.[U_Z_EMPCODE] AS 'EMPLOYEE ID', T0.[U_Z_EMPNAME],  [U_Z_TYPE]  as 'Project Code/ Leave Type' ,' ' 'U_Z_PRJNAME' ,' '   AS 'PHASE' ,' '  AS 'ACTIVITY',  T0.[U_Z_DAYS] 'Days',T0.[U_Z_DAYS] * 8.0 'Hours' FROM [@Z_OLEV]  T0  where T0.[U_Z_EMPCODE] in (" & strEMPIDS & ") and  " & strEMPCondition & "   and    ( " & strDateCondition & ")"
                        strSQL = strSQL & "  union all"
                        strSQL = strSQL & " SELECT T1.[U_Z_DATE],T0.[U_Z_EMPCODE] AS 'EMPLOYEE ID', T0.[U_Z_EMPNAME],  T1.[U_Z_PRJCODE] as 'Project Code/ Leave Type',T1.[U_Z_PRJNAME], T1.[U_Z_PRCNAME]  AS 'PHASE', T1.[U_Z_ACTNAME] AS 'ACTIVITY',T1.[U_Z_HOURS] /8.0 'Days',T1.[U_Z_HOURS] 'Hours'  FROM [@Z_OTIM]  T0  inner join  [@Z_TIM1]  T1 on T0.Code=T1.U_Z_REFCODE   where T0.[U_Z_EMPCODE] in (" & strEMPIDS & ") and  T1.U_Z_PRJCODE In (" & strExcludeCondition & ") and  " & strEMPCondition & "   and    ( " & strDateCondition & ")"

                    End If


                    strSQL = "select X.U_Z_DOCDATE 'Date',X.[EMPLOYEE ID] 'Employee ID',x.U_Z_EMPNAME 'Employee Name',x.[Project Code/ Leave Type] 'Project Code/LeaveType',x.U_Z_PRJNAME 'Project Name',x.PHASE 'Phase',x.ACTIVITY ,x.[Days] 'Days',x.Hours ' Hours' from (" & strSQL & ") X  order by x.U_Z_DOCDATE ,x.[EMPLOYEE ID],x.[Project Code/ Leave Type]"
                Case "PRS"  'Project Status'
                    'strSQL = "select X.U_Z_PRJCODE 'Project Code',x.U_Z_PRJNAME 'Project Name',x.U_Z_CARDNAME 'Customer Name' ,x.U_Z_MODNAME 'Phase',isnull(x.TotalHours,0) 'Estimated Hours',isnull(x.TotalDays,0) 'Estimated Days',isnull(SUM(T2.U_Z_HOURS),0) 'Actual Hours',isnull(x.EstimatedCost,0) 'EstimatedCost',isnull(x.EstimatedCost,0) 'ActualCost',isnull(x.EstimatedCost,0) -isnull(x.EstimatedCost,0) 'Variance In Cost','Completed' 'Status' from "
                    ' strSQL = strSQL & " ( select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T0.U_Z_CARDCODE,T0.U_Z_CARDNAME  , T1.U_Z_MODNAME ,Sum(T1.U_Z_HOURS) 'TotalHours' ,Sum(T1.U_Z_Days) 'TotalDays',Sum(T1.U_Z_AMOUNT) 'EstimatedCost'  from  [@Z_HPRJ]  T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where U_Z_TYPE='R' group by T0.U_Z_CARDCODE , T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T0.U_Z_CARDNAME , T1.U_Z_MODNAME ) x Left Outer Join [@Z_TIM1] T2 on T2.U_Z_PRJCODE=x.U_Z_PRJCODE and T2.U_Z_PRCNAME =X.U_Z_MODNAME where " & strProjectCondition & "  group by X.U_Z_PRJCODE,x.U_Z_PRJNAME,x.U_Z_MODNAME,x.U_Z_CARDNAME,x.TotalHours,x.TotalDays,x.EstimatedCost order by  X.U_Z_PRJCODE,x.U_Z_PRJNAME,x.U_Z_MODNAME"
                    If blnSuperUser = True Then
                        strSQL = "select X.U_Z_PRJCODE 'Project Code',x.U_Z_PRJNAME 'Project Name',x.U_Z_CARDNAME 'Customer Name' ,x.U_Z_MODNAME 'Phase',isnull(x.TotalDays,0) 'Estimated Days',isnull(SUM(T2.U_Z_HOURS),0)/8 'Actual Days',isnull(x.TotalDays,0)-(isnull(SUM(T2.U_Z_HOURS),0))/8 'Variance In Days',isnull(x.EstimatedCost,0) 'EstimatedCost',isnull(x.EstimatedCost,0) 'ActualCost',isnull(x.EstimatedCost,0) -isnull(x.EstimatedCost,0) 'Variance In Cost','Completed' 'Status' from "
                        strSQL = strSQL & " ( select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T0.U_Z_CARDCODE,T0.U_Z_CARDNAME  , T1.U_Z_MODNAME ,Sum(T1.U_Z_HOURS) 'TotalHours' ,Sum(T1.U_Z_Days) 'TotalDays',Sum(T1.U_Z_AMOUNT) 'EstimatedCost'  from  [@Z_HPRJ]  T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where (U_Z_TYPE='R' or U_Z_TYPE='E'  or U_Z_TYPE='I') group by T0.U_Z_CARDCODE , T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T0.U_Z_CARDNAME , T1.U_Z_MODNAME ) x Left Outer Join [@Z_TIM1] T2 on T2.U_Z_PRJCODE=x.U_Z_PRJCODE and T2.U_Z_PRCNAME =X.U_Z_MODNAME where x.U_Z_PRJCODE In (" & strExcludeCondition & ") and  " & strProjectCondition & "  group by X.U_Z_PRJCODE,x.U_Z_PRJNAME,x.U_Z_MODNAME,x.U_Z_CARDNAME,x.TotalHours,x.TotalDays,x.EstimatedCost order by  X.U_Z_PRJCODE,x.U_Z_PRJNAME,x.U_Z_MODNAME"

                    Else
                        strSQL = "select X.U_Z_PRJCODE 'Project Code',x.U_Z_PRJNAME 'Project Name',x.U_Z_CARDNAME 'Customer Name' ,x.U_Z_MODNAME 'Phase',isnull(x.TotalDays,0) 'Estimated Days',isnull(SUM(T2.U_Z_HOURS),0)/8 'Actual Days',isnull(x.TotalDays,0)-(isnull(SUM(T2.U_Z_HOURS),0))/8 'Variance In Days',isnull(x.EstimatedCost,0) 'EstimatedCost',isnull(x.EstimatedCost,0) 'ActualCost',isnull(x.EstimatedCost,0) -isnull(x.EstimatedCost,0) 'Variance In Cost','Completed' 'Status' from "
                        strSQL = strSQL & " ( select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T0.U_Z_CARDCODE,T0.U_Z_CARDNAME  , T1.U_Z_MODNAME ,Sum(T1.U_Z_HOURS) 'TotalHours' ,Sum(T1.U_Z_Days) 'TotalDays',Sum(T1.U_Z_AMOUNT) 'EstimatedCost'  from  [@Z_HPRJ]  T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where (U_Z_TYPE='R' or U_Z_TYPE='E or U_Z_TYPE='I' ) and T0.U_Z_EMPID='" & intUser & "'  group by T0.U_Z_CARDCODE , T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T0.U_Z_CARDNAME , T1.U_Z_MODNAME ) x Left Outer Join [@Z_TIM1] T2 on T2.U_Z_PRJCODE=x.U_Z_PRJCODE and T2.U_Z_PRCNAME =X.U_Z_MODNAME where x.U_Z_PRJCODE In (" & strExcludeCondition & ") and  " & strProjectCondition & "  group by X.U_Z_PRJCODE,x.U_Z_PRJNAME,x.U_Z_MODNAME,x.U_Z_CARDNAME,x.TotalHours,x.TotalDays,x.EstimatedCost order by  X.U_Z_PRJCODE,x.U_Z_PRJNAME,x.U_Z_MODNAME"

                    End If



                Case "PRJS" ' Project Summary
                    'strSQL = "Select  X.U_Z_PRJCODE 'Project Code',x.U_Z_PRJNAME 'Project Name',x.U_Z_CARDNAME 'Customer Name' ,isnull((x.TotalHours),0) 'Estimated Hours',isnull(SUM(T2.U_Z_HOURS),0) 'Actual Hours',isnull((x.TotalHours),0) - isnull(SUM(T2.U_Z_HOURS),0) 'Variance',(isnull((x.TotalHours),0) - isnull(SUM(T2.U_Z_HOURS),0))/8 'Variance in Days','                                                        ' 'Current Phase' from  "
                    'strSQL = strSQL & "( select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T0.U_Z_CARDCODE,T0.U_Z_CARDNAME,  Sum(T1.U_Z_HOURS) 'TotalHours' ,Sum(T1.U_Z_Days) 'TotalDays',Sum(T1.U_Z_AMOUNT) 'EstimatedCost'  from  [@Z_HPRJ]  T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where U_Z_TYPE='R' group by T0.U_Z_CARDCODE , T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T0.U_Z_CARDNAME ) x "
                    'strSQL = strSQL & "Left Outer Join [@Z_TIM1] T2 on T2.U_Z_PRJCODE=x.U_Z_PRJCODE   where " & strProjectCondition & "  group by X.U_Z_PRJCODE,x.U_Z_PRJNAME,x.U_Z_CARDNAME,x.TotalHours order by  X.U_Z_PRJCODE,x.U_Z_PRJNAME"

                    If blnSuperUser = True Then
                        strSQL = "Select  X.U_Z_PRJCODE 'Project Code',x.U_Z_PRJNAME 'Project Name',x.U_Z_CARDNAME 'Customer Name' ,isnull((x.TotalDays),0) 'Estimated Days',isnull(SUM(T2.U_Z_HOURS),0) /8  'Actual Days',isnull((x.TotalDays),0) - (isnull(SUM(T2.U_Z_HOURS),0)/8) 'Variance In Days',x.EstimatedCost, x.EstimatedCost-x.EstimatedCost 'Actual Cost',x.EstimatedCost-x.EstimatedCost 'Variance In Cost'" & strexpquery & " '                                                        ' 'Current Phase' from  "
                        strSQL = strSQL & "( select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T0.U_Z_CARDCODE,T0.U_Z_CARDNAME,  Sum(T1.U_Z_HOURS) 'TotalHours' ,Sum(T1.U_Z_Days) 'TotalDays',Sum(T1.U_Z_AMOUNT) 'EstimatedCost'  from  [@Z_HPRJ]  T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where (U_Z_TYPE='R' or U_Z_TYPE='E'  or U_Z_TYPE='I') group by T0.U_Z_CARDCODE , T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T0.U_Z_CARDNAME ) x "
                        strSQL = strSQL & "Left Outer Join [@Z_TIM1] T2 on T2.U_Z_PRJCODE=x.U_Z_PRJCODE   where X.U_Z_PRJCODE In (" & strExcludeCondition & ") and  " & strProjectCondition & "  group by X.U_Z_PRJCODE,x.U_Z_PRJNAME,x.U_Z_CARDNAME,x.TotalDays,x.EstimatedCost order by  X.U_Z_PRJCODE,x.U_Z_PRJNAME"

                    Else
                        strSQL = "Select  X.U_Z_PRJCODE 'Project Code',x.U_Z_PRJNAME 'Project Name',x.U_Z_CARDNAME 'Customer Name' ,isnull((x.TotalDays),0) 'Estimated Days',isnull(SUM(T2.U_Z_HOURS),0) /8  'Actual Days',isnull((x.TotalDays),0) - (isnull(SUM(T2.U_Z_HOURS),0)/8) 'Variance In Days',x.EstimatedCost, x.EstimatedCost-x.EstimatedCost 'Actual Cost',x.EstimatedCost-x.EstimatedCost 'Variance In Cost'" & strexpquery & " '                                                        ' 'Current Phase' from  "
                        strSQL = strSQL & "( select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T0.U_Z_CARDCODE,T0.U_Z_CARDNAME,  Sum(T1.U_Z_HOURS) 'TotalHours' ,Sum(T1.U_Z_Days) 'TotalDays',Sum(T1.U_Z_AMOUNT) 'EstimatedCost'  from  [@Z_HPRJ]  T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where (U_Z_TYPE='R' or U_Z_TYPE='E'  or U_Z_TYPE='I') and T0.U_Z_EMPID='" & intUser & "' group by T0.U_Z_CARDCODE , T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T0.U_Z_CARDNAME ) x "
                        strSQL = strSQL & "Left Outer Join [@Z_TIM1] T2 on T2.U_Z_PRJCODE=x.U_Z_PRJCODE   where X.U_Z_PRJCODE In (" & strExcludeCondition & ") and  " & strProjectCondition & "  group by X.U_Z_PRJCODE,x.U_Z_PRJNAME,x.U_Z_CARDNAME,x.TotalDays,x.EstimatedCost order by  X.U_Z_PRJCODE,x.U_Z_PRJNAME"
                    End If
                Case "RS" 'Resource Utilization
                    If blnSuperUser = True Then
                        strSQL = "select T1.U_Z_EMPID,T1.U_Z_POSITION ,T1.U_Z_FROMDATE,T1.U_Z_TODATE ,T0.U_Z_CARDNAME,T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T1.U_Z_Modname,T1.U_Z_ACTNAME ,  T1.U_Z_HOURS/8 'TotalHours' , T1.U_Z_HOURS 'ActualHours' ,T1.U_Z_HOURS 'Variance' from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry  where T0.U_Z_PRJCODE In (" & strExcludeCondition & ") and  (" & strExmEmpCondition & ") and (" & projectMaterialcond & ")  and (" & strDateCondition1 & " or " & strDateCondition2 & ")   Order By T1.U_Z_EMPID,T1.U_Z_FROMDATE,T0.U_Z_PRJCODE"

                    Else
                        strSQL = "select T1.U_Z_EMPID,T1.U_Z_POSITION ,T1.U_Z_FROMDATE,T1.U_Z_TODATE ,T0.U_Z_CARDNAME,T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T1.U_Z_Modname,T1.U_Z_ACTNAME ,  T1.U_Z_HOURS/8 'TotalHours' , T1.U_Z_HOURS 'ActualHours' ,T1.U_Z_HOURS 'Variance' from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry  where T1.U_Z_EMPID in (" & strEMPIDS & ") and  T0.U_Z_PRJCODE In (" & strExcludeCondition & ") and  (" & strExmEmpCondition & ") and (" & projectMaterialcond & ")  and (" & strDateCondition1 & " or " & strDateCondition2 & ")   Order By T1.U_Z_EMPID,T1.U_Z_FROMDATE,T0.U_Z_PRJCODE"
                    End If
            End Select
            oGrid = aForm.Items.Item("18").Specific
            oGrid.DataTable.ExecuteQuery(strSQL)
            If strreporttype = "PRS" Then
                ProjectStatus(oGrid)
            ElseIf strreporttype = "PRJS" Then
                ProjectSummary(oGrid)
            ElseIf strreporttype = "RS" Then
                EmployeeUtilization(oGrid)
            End If
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next
            oGrid.Columns.Item("RowsHeader").TitleObject.Caption = "#"
            FormatGrid(oGrid, strreporttype)
            aForm.Freeze(False)
            Exit Sub




            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#Region "Format Grid"
    Private Sub FormatGrid(ByVal agrid As SAPbouiCOM.Grid, ByVal aType As String)
        oGrid = agrid
        Select Case aType
            Case "TA"
                oGrid.Columns.Item(0).TitleObject.Caption = "Date"
                oGrid.Columns.Item(0).TitleObject.Sortable = True

                oGrid.Columns.Item(1).TitleObject.Caption = "Employee ID"
                oEditTextColumn = oGrid.Columns.Item(1)
                oEditTextColumn.LinkedObjectType = "171"
                oGrid.Columns.Item(1).TitleObject.Sortable = True

                oGrid.Columns.Item(2).TitleObject.Caption = "Employee Name"
                oGrid.Columns.Item(3).TitleObject.Caption = "Project Code / Leave Type"
                oGrid.Columns.Item(3).TitleObject.Sortable = True

                oGrid.Columns.Item(4).TitleObject.Caption = "Project Name"
                oGrid.Columns.Item(5).TitleObject.Caption = "Phase"
                oGrid.Columns.Item(6).TitleObject.Sortable = True

                oGrid.Columns.Item(6).TitleObject.Caption = "Activity"
                oGrid.Columns.Item(7).TitleObject.Caption = "Days "
                oGrid.Columns.Item(7).TitleObject.Sortable = True

                oGrid.Columns.Item(8).TitleObject.Caption = " Hours"
                oGrid.Columns.Item(8).TitleObject.Sortable = True
                oEditTextColumn = oGrid.Columns.Item(7)
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

                oEditTextColumn = oGrid.Columns.Item(8)
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto


            Case "PRS"
                strSQL = "select X.U_Z_PRJCODE 'Project Code',x.U_Z_PRJNAME 'Project Name',x.U_Z_CARDNAME 'Customer Name' ,x.U_Z_MODNAME 'Phase',x.TotalHours 'Estimated Hours',x.TotalDays 'Estimated Days',SUM(T2.U_Z_HOURS) 'Actual Hours',x.EstimatedCost,0 'ActualCost','Completed' 'Status' from "
                oGrid.Columns.Item(0).TitleObject.Caption = "Project Code"
                oGrid.Columns.Item(0).TitleObject.Sortable = True
                oEditTextColumn = oGrid.Columns.Item(0)
                oEditTextColumn.LinkedObjectType = "63"
                oEditTextColumn = oGrid.Columns.Item("Customer Name")
                oEditTextColumn.LinkedObjectType = "2"
                oGrid.Columns.Item("ActualCost").Visible = True
                oEditTextColumn = oGrid.Columns.Item("Estimated Days")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = oGrid.Columns.Item("Variance In Days")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = oGrid.Columns.Item("Actual Days")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = oGrid.Columns.Item("EstimatedCost")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = oGrid.Columns.Item("ActualCost")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = oGrid.Columns.Item("Variance In Cost")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto


            Case "PRJS"

                strSQL = "Select  X.U_Z_PRJCODE 'Project Code',x.U_Z_PRJNAME 'Project Name',x.U_Z_CARDNAME 'Customer Name' ,(x.TotalHours) 'Estimated Hours',SUM(T2.U_Z_HOURS) 'Actual Hours',(x.TotalHours) - SUM(T2.U_Z_HOURS) 'Variance',((x.TotalHours) - SUM(T2.U_Z_HOURS))/8 'Variance in Days','                                                        ' 'Current Phase' from  "
                oGrid.Columns.Item(0).TitleObject.Caption = "Project Code"
                oGrid.Columns.Item(0).TitleObject.Sortable = True
                oEditTextColumn = oGrid.Columns.Item(0)
                oEditTextColumn.LinkedObjectType = "63"
                oEditTextColumn = oGrid.Columns.Item("Customer Name")
                oEditTextColumn.LinkedObjectType = "2"

                oEditTextColumn = oGrid.Columns.Item("Estimated Days")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = oGrid.Columns.Item("Variance In Days")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = oGrid.Columns.Item("Actual Days")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = oGrid.Columns.Item("Variance In Cost")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = oGrid.Columns.Item("EstimatedCost")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = oGrid.Columns.Item("Actual Cost")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto


                oEditTextColumn = oGrid.Columns.Item("Estimated Expenses")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = oGrid.Columns.Item("Acutal Approved Expenses")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = oGrid.Columns.Item("Variance in Expenses")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto


                oEditTextColumn = oGrid.Columns.Item("Total Project Cost")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = oGrid.Columns.Item("Total Actual Cost")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = oGrid.Columns.Item("Total Variance Cost")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto


                oEditTextColumn = oGrid.Columns.Item("Pening for Approval Expenses")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            Case "RS"
                ' strSQL = "select T1.U_Z_EMPID,T1.U_Z_POSITION ,T1.U_Z_FROMDATE,T1.U_Z_TODATE ,T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T1.U_Z_Modname,T1.U_Z_ACTNAME , T0.U_Z_CARDNAME, T1.U_Z_HOURS 'TotalHours' , T1.U_Z_HOURS 'ActualHours' ,T1.U_Z_HOURS 'Variance' from [@Z_HPRJ] T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry  where (" & strExmEmpCondition & ") and (" & projectMaterialcond & ")  and (" & strDateCondition1 & " or " & strDateCondition2 & ")   Order By T1.U_Z_EMPID,T1.U_Z_FROMDATE,T0.U_Z_PRJCODE"
                oGrid.Columns.Item("U_Z_EMPID").TitleObject.Caption = "Employee ID"
                oGrid.Columns.Item("U_Z_EMPID").TitleObject.Sortable = True
                oEditTextColumn = oGrid.Columns.Item("U_Z_EMPID")
                oEditTextColumn.LinkedObjectType = "171"
                oGrid.Columns.Item("U_Z_POSITION").TitleObject.Caption = "Employee Name"
                oGrid.Columns.Item("U_Z_FROMDATE").TitleObject.Caption = "From Date"
                oGrid.Columns.Item("U_Z_FROMDATE").TitleObject.Sortable = True
                oGrid.Columns.Item("U_Z_TODATE").TitleObject.Caption = "To Date"
                oGrid.Columns.Item("U_Z_TODATE").TitleObject.Sortable = True
                oGrid.Columns.Item("U_Z_PRJCODE").TitleObject.Caption = "Project Code"
                oGrid.Columns.Item("U_Z_PRJCODE").TitleObject.Sortable = True
                oEditTextColumn = oGrid.Columns.Item("U_Z_PRJCODE")
                oEditTextColumn.LinkedObjectType = "63"
                oGrid.Columns.Item("U_Z_PRJNAME").TitleObject.Caption = "Project Name"
                oGrid.Columns.Item("U_Z_Modname").TitleObject.Caption = "Phase"
                oGrid.Columns.Item("U_Z_ACTNAME").TitleObject.Caption = "Activity"
                oGrid.Columns.Item("U_Z_CARDNAME").TitleObject.Caption = "Customer Name"
                oGrid.Columns.Item("TotalHours").TitleObject.Caption = "Total Days"
                oEditTextColumn = oGrid.Columns.Item("TotalHours")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oGrid.Columns.Item("ActualHours").TitleObject.Caption = "Actual Days"
                oEditTextColumn = oGrid.Columns.Item("ActualHours")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oGrid.Columns.Item("ActualHours").TitleObject.Sortable = True
                oGrid.Columns.Item("Variance").TitleObject.Caption = "Variance In Days"
                oEditTextColumn = oGrid.Columns.Item("Variance")
                oGrid.Columns.Item("Variance").TitleObject.Sortable = True

                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

        End Select
        oGrid.AutoResizeColumns()
        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
    End Sub
#End Region

#Region "Project Status"
    Private Sub ProjectStatus(ByVal aGrid As SAPbouiCOM.Grid)
        Dim oTempRS As SAPbobsCOM.Recordset
        Dim strSQL, strProject, strPhase As String
        Dim dblEstimatedcost, dblAcutalcost, dblVarianceCost As Double
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            strProject = aGrid.DataTable.GetValue(0, intRow)
            strPhase = aGrid.DataTable.GetValue(3, intRow)
            strSQL = " select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T1.U_Z_MODNAME,T1.U_Z_STATUS, Count(*) 'Count' from  [@Z_HPRJ]  T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where  U_Z_TYPE='R' and T1.U_Z_STATUS<>'C' and U_Z_PRJCODE='" & strProject & "' and U_Z_MODNAME='" & strPhase & "' group by T0.U_Z_PRJCODE,T0.U_Z_PRJNAME, T1.U_Z_MODNAME ,T1.U_Z_STATUS"
            oTempRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempRS.DoQuery(strSQL)
            If oTempRS.RecordCount > 0 Then
                If oTempRS.Fields.Item(3).Value = "I" Then
                    aGrid.DataTable.SetValue("Status", intRow, "InProcess")
                Else
                    aGrid.DataTable.SetValue("Status", intRow, "Pending")
                End If
                '  aGrid.DataTable.SetValue("Status", intRow, oTempRS.Fields.Item(3).Value)
            Else

                aGrid.DataTable.SetValue("Status", intRow, "Completed")
            End If

            strSQL = "select sum(isnull(T3.U_HR_RATE,0) * isnull(T1.U_Z_HOURS,0) /1) from [@Z_OTIM] T0 inner join [@Z_TIM1] T1 on T1.U_Z_REFCODE=T0.Code inner Join OHEM T3 on convert(Varchar,T3.empID)=T0.U_Z_EMPCODE "
            strSQL = strSQL & "  where T1.U_Z_PRJCODE='" & strProject & "'  and T1.U_Z_PRCNAME='" & strPhase & "'"
            oTempRS.DoQuery(strSQL)
            dblAcutalcost = oTempRS.Fields.Item(0).Value
            dblEstimatedcost = oGrid.DataTable.GetValue("EstimatedCost", intRow)

            Dim dblactualexpenses As Double
            strSQL = "SELECT T0.[MainCurncy] FROM OADM T0"
            oTempRS.DoQuery(strSQL)
            Dim strLocalCurrency As String = oTempRS.Fields.Item(0).Value

            strSQL = "select isnull(Sum(Convert(Decimal,REPLACE(U_Z_LOCAMT,'" & strLocalCurrency & "',''))),0) from [@Z_EXP1]  where U_Z_Approved='A' and  U_Z_PRJCODE ='" & strProject & "' and U_Z_MODNAME='" & strPhase & "'"
            oTempRS.DoQuery(strSQL)
            If oTempRS.RecordCount > 0 Then
                dblactualexpenses = oTempRS.Fields.Item(0).Value
            Else
                dblactualexpenses = 0
            End If

            dblAcutalcost = dblAcutalcost + dblactualexpenses

            dblVarianceCost = dblEstimatedcost - dblAcutalcost
            aGrid.DataTable.SetValue("ActualCost", intRow, dblAcutalcost)
            aGrid.DataTable.SetValue("Variance In Cost", intRow, dblVarianceCost)
            ' aGrid.RowHeaders.SetText(intRow, intRow + 1)
        Next

    End Sub

    Private Sub ProjectSummary(ByVal aGrid As SAPbouiCOM.Grid)
        Dim oTempRS As SAPbobsCOM.Recordset
        Dim strSQL, strProject, strPhase As String
        Dim dblEstimatedcost, dblAcutalcost, dblVarianceCost As Double

        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            strProject = aGrid.DataTable.GetValue(0, intRow)
            '   strPhase = aGrid.DataTable.GetValue(3, intRow)
            strSQL = " select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T1.U_Z_MODNAME,T1.U_Z_STATUS, Count(*) 'Count' from  [@Z_HPRJ]  T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where  U_Z_TYPE='R' and T1.U_Z_STATUS<>'C' and U_Z_PRJCODE='" & strProject & "' group by T0.U_Z_PRJCODE,T0.U_Z_PRJNAME, T1.U_Z_MODNAME ,T1.U_Z_STATUS"
            oTempRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempRS.DoQuery(strSQL)
            If oTempRS.RecordCount > 0 Then
                If oTempRS.Fields.Item(3).Value = "I" Then
                    aGrid.DataTable.SetValue("Current Phase", intRow, oTempRS.Fields.Item(2).Value)
                Else
                    aGrid.DataTable.SetValue("Current Phase", intRow, oTempRS.Fields.Item(2).Value)
                End If
                '  aGrid.DataTable.SetValue("Status", intRow, oTempRS.Fields.Item(3).Value)
            Else
                aGrid.DataTable.SetValue("Current Phase", intRow, "Completed")
            End If


            strSQL = " select T0.U_Z_PRJCODE, Sum(U_Z_AMOUNT) 'Count' from  [@Z_HPRJ]  T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where  U_Z_TYPE='R' or U_Z_TYPE='I' and U_Z_PRJCODE='" & strProject & "' group by T0.U_Z_PRJCODE"
            oTempRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempRS.DoQuery(strSQL)
            If oTempRS.RecordCount > 0 Then
                dblEstimatedcost = oTempRS.Fields.Item(1).Value
            Else
                dblEstimatedcost = 0
            End If
            oGrid.DataTable.SetValue("EstimatedCost", intRow, dblEstimatedcost)

            strSQL = "select sum(isnull(T3.U_HR_RATE,0) * isnull(T1.U_Z_HOURS,0) /1) from [@Z_OTIM] T0 inner join [@Z_TIM1] T1 on T1.U_Z_REFCODE=T0.Code inner Join OHEM T3 on convert(nvarchar,T3.empID)=T0.U_Z_EMPCODE "
            strSQL = strSQL & "  where T1.U_Z_PRJCODE='" & strProject & "'" '  and T1.U_Z_PRCNAME='" & strPhase & "'"
            oTempRS.DoQuery(strSQL)
            Dim dblTotalProjectCost, dblActualProjectCost, dblVarianceProjectCost As Double
            dblAcutalcost = oTempRS.Fields.Item(0).Value
            dblActualProjectCost = dblAcutalcost

          

            dblEstimatedcost = oGrid.DataTable.GetValue("EstimatedCost", intRow)
            dblTotalProjectCost = dblEstimatedcost
            dblVarianceCost = dblEstimatedcost - dblAcutalcost
            aGrid.DataTable.SetValue("Actual Cost", intRow, oTempRS.Fields.Item(0).Value)
            aGrid.DataTable.SetValue("Variance In Cost", intRow, dblVarianceCost)

            Dim dblestimatedExpense, dblactualexpenses, dbleVarianceExpenses As Double
            ' strSQL = " select T0.U_Z_PRJCODE, Sum(U_Z_AMOUNT) 'Count' from  [@Z_HPRJ]  T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where  U_Z_TYPE='E' and U_Z_PRJCODE='" & strProject & "' group by T0.U_Z_PRJCODE"
            strSQL = " select T0.U_Z_PRJCODE, Sum(U_Z_AMOUNT) 'Count' from  [@Z_HPRJ]  T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where  U_Z_TYPE='E' and U_Z_PRJCODE='" & strProject & "' group by T0.U_Z_PRJCODE"
            oTempRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempRS.DoQuery(strSQL)
            If oTempRS.RecordCount > 0 Then
                dblEstimatedcost = oTempRS.Fields.Item(1).Value
            Else
                dblEstimatedcost = 0
            End If
            strSQL = "SELECT T0.[MainCurncy] FROM OADM T0"
            oTempRS.DoQuery(strSQL)
            Dim strLocalCurrency As String = oTempRS.Fields.Item(0).Value

            strSQL = "select isnull(Sum(Convert(Decimal,REPLACE(U_Z_LOCAMT,'" & strLocalCurrency & "',''))),0) from [@Z_EXP1]  where U_Z_Approved='A' and  U_Z_PRJCODE ='" & strProject & "'"
            oTempRS.DoQuery(strSQL)
            If oTempRS.RecordCount > 0 Then
                dblactualexpenses = oTempRS.Fields.Item(0).Value
            Else
                dblactualexpenses = 0
            End If
            dblTotalProjectCost = dblTotalProjectCost + dblEstimatedcost
            dblActualProjectCost = dblActualProjectCost + dblactualexpenses
            dblVarianceCost = dblEstimatedcost - dblactualexpenses
            aGrid.DataTable.SetValue("Estimated Expenses", intRow, dblEstimatedcost)
            aGrid.DataTable.SetValue("Acutal Approved Expenses", intRow, dblactualexpenses)
            aGrid.DataTable.SetValue("Variance in Expenses", intRow, dblVarianceCost)


            aGrid.DataTable.SetValue("Total Project Cost", intRow, dblTotalProjectCost)
            aGrid.DataTable.SetValue("Total Actual Cost", intRow, dblActualProjectCost)
            dblActualProjectCost = dblTotalProjectCost - dblActualProjectCost
            aGrid.DataTable.SetValue("Total Variance Cost", intRow, dblActualProjectCost)


            strSQL = "select isnull(Sum(Convert(Decimal,REPLACE(U_Z_LOCAMT,'" & strLocalCurrency & "',''))),0) from [@Z_EXP1]  where U_Z_Approved='P' and  U_Z_PRJCODE ='" & strProject & "'"
            oTempRS.DoQuery(strSQL)
            If oTempRS.RecordCount > 0 Then
                dblactualexpenses = oTempRS.Fields.Item(0).Value
            Else
                dblactualexpenses = 0
            End If
            aGrid.DataTable.SetValue("Pening for Approval Expenses", intRow, dblactualexpenses)
        Next
    End Sub
    Private Sub EmployeeUtilization(ByVal aGrid As SAPbouiCOM.Grid)
        Dim oTempRS As SAPbobsCOM.Recordset
        Dim strSQL, strProject, strPhase, strActvity, strEmpID As String
        Dim dblTotalhours, dblAcutalhours, dblVariance As Double
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            strEmpID = aGrid.DataTable.GetValue(0, intRow)
            strProject = aGrid.DataTable.GetValue("U_Z_PRJCODE", intRow)
            strActvity = aGrid.DataTable.GetValue("U_Z_ACTNAME", intRow)
            strPhase = aGrid.DataTable.GetValue("U_Z_Modname", intRow)
            dblTotalhours = aGrid.DataTable.GetValue("TotalHours", intRow)
            strSQL = " select T0.U_Z_PRJCODE,T0.U_Z_PRJNAME,T1.U_Z_MODNAME,T1.U_Z_STATUS, Count(*) 'Count' from  [@Z_HPRJ]  T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where  U_Z_TYPE='R' and T1.U_Z_STATUS<>'C' and U_Z_PRJCODE='" & strProject & "' group by T0.U_Z_PRJCODE,T0.U_Z_PRJNAME, T1.U_Z_MODNAME ,T1.U_Z_STATUS"
            strSQL = "select SUM(T1.U_Z_HOURS) from [@Z_OTIM] T0 inner join [@Z_TIM1] T1 on T1.U_Z_REFCODE=T0.Code where T0.U_Z_EMPCODE='" & strEmpID & "' and  T1.U_Z_PRJCODE='" & strProject & "' and  T1.U_Z_PRCNAME ='" & strPhase & "' and T1.U_Z_ACTNAME='" & strActvity & "'"
            oTempRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempRS.DoQuery(strSQL)
            dblAcutalhours = oTempRS.Fields.Item(0).Value
            dblAcutalhours = dblAcutalhours / 8
            dblVariance = dblTotalhours - dblAcutalhours
            aGrid.DataTable.SetValue("ActualHours", intRow, dblAcutalhours)
            aGrid.DataTable.SetValue("Variance", intRow, dblVariance)
            ' aGrid.RowHeaders.SetText(intRow, intRow + 1)
        Next

    End Sub
#End Region

    Private Function JournalQuery(ByVal aForm As SAPbouiCOM.Form) As String
        Try
            aForm.Freeze(True)
            Dim strString1, strString2, strreporttype, strCode, strString, strFromEMP, strToEMP, strFromPRJ, strToPRJ, strfromdate, strToDate, strEMPCondition, strDateCondition, strProjectCondition As String
            Dim dtFromdate, dtTodate As Date

            strFromEMP = oApplication.Utilities.getEdittextvalue(aForm, "9")
            strToEMP = oApplication.Utilities.getEdittextvalue(aForm, "12")

            strfromdate = oApplication.Utilities.getEdittextvalue(aForm, "18")
            strToDate = oApplication.Utilities.getEdittextvalue(aForm, "20")

            If strfromdate <> "" Then
                dtFromdate = oApplication.Utilities.GetDateTimeValue(strfromdate)
            End If
            If strToDate <> "" Then
                dtTodate = oApplication.Utilities.GetDateTimeValue(strToDate)
            End If
            'oCombo = aForm.Items.Item("4").Specific
            'strreporttype = oCombo.Selected.Value
            If aForm.Title = "Reports" Then
                oCombo = aForm.Items.Item("4").Specific
                strreporttype = oCombo.Selected.Value
                If strreporttype = "" Then
                    oApplication.Utilities.Message("Report type missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Exit Function
                End If
            Else
                oCombo = aForm.Items.Item("36").Specific
                strreporttype = oApplication.Utilities.getEdittextvalue(aForm, "102")
                If strreporttype = "" And oCombo.Selected.Value = "E" Then
                    oApplication.Utilities.Message("Account is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Exit Function
                Else
                    oitems = aForm.Items.Item("29")
                    obutton = oitems.Specific
                    If obutton.Caption.Contains("Copy to Journal") Then
                        strreporttype = "Posting"
                    Else
                        strreporttype = "PostJE"
                    End If
                End If
            End If


            oCombo = aForm.Items.Item("6").Specific
            strFromPRJ = oCombo.Selected.Value
            oCombo = aForm.Items.Item("15").Specific
            strToPRJ = oCombo.Selected.Value

            strEMPCondition = ""
            strDateCondition = ""
            strProjectCondition = ""
            If strFromEMP <> "" And strToEMP <> "" Then
                strEMPCondition = " Convert(Decimal,U_Z_EMPCODE) between " & CDbl(strFromEMP) & " and " & CDbl(strToEMP)
            ElseIf strFromEMP <> "" And strToEMP = "" Then
                strEMPCondition = " Convert(Decimal,U_Z_EMPCODE) >= " & CDbl(strFromEMP)
            ElseIf strFromEMP = "" And strToEMP <> "" Then
                strEMPCondition = " Convert(Decimal,U_Z_EMPCODE) <= " & CDbl(strToEMP)
            Else
                strEMPCondition = " 1=1"
            End If

            If strFromPRJ <> "" And strToPRJ <> "" Then
                strProjectCondition = " U_Z_PRJCODE between '" & strFromPRJ & "' and '" & strToPRJ & "'"
            ElseIf strFromPRJ <> "" And strToPRJ = "" Then
                strProjectCondition = " U_Z_PRJCODE >= '" & strFromPRJ & "'"
            ElseIf strFromPRJ = "" And strToPRJ <> "" Then
                strProjectCondition = " U_Z_PRJCODE <= '" & strToPRJ & "'"
            Else
                strProjectCondition = " 1=1"
            End If

            'If strfromdate <> "" And strToDate <> "" Then
            '     strDateCondition = " U_Z_DocDate between '" & dtFromdate.ToString("yyyy-MM-dd") & "' and '" & dtTodate.ToString("yyyy-MM-dd") & "'"
            'ElseIf strfromdate <> "" And strToDate = "" Then
            '     strDateCondition = " U_Z_DocDate>= '" & dtFromdate.ToString("yyyy-MM-dd") & "'"
            'ElseIf strfromdate = "" And strToDate <> "" Then
            '    strDateCondition = " U_Z_DocDate <= '" & dtTodate.ToString("yyyy-MM-dd") & "'"
            'Else
            '    strDateCondition = " 1=1"
            'End If

            Dim strDateCondition1 As String = ""

            If strfromdate <> "" And strToDate <> "" Then
                strDateCondition = " U_Z_DocDate > = '" & dtFromdate.ToString("yyyy-MM-dd") & "' and U_Z_DocDate <= '" & dtTodate.ToString("yyyy-MM-dd") & "'"
                strDateCondition1 = " U_Z_DocDate > = '" & dtFromdate.ToString("yyyy-dd-MM") & "' and U_Z_DocDate <= '" & dtTodate.ToString("yyyy-dd-MM") & "'"
                ' strDateCondition1 = " U_Z_DocDate > = '" & dtFromdate.ToString("MM-dd-yyyy") & "' and U_Z_DocDate <= '" & dtTodate.ToString("MM-dd-yyyy") & "'"

            ElseIf strfromdate <> "" And strToDate = "" Then
                strDateCondition = " U_Z_DocDate >= '" & dtFromdate.ToString("yyyy-MM-dd") & "'"
                strDateCondition1 = " U_Z_DocDate >= '" & dtFromdate.ToString("yyyy-dd-MM") & "'"

            ElseIf strfromdate = "" And strToDate <> "" Then
                strDateCondition = " U_Z_DocDate <= '" & dtTodate.ToString("yyyy-MM-dd") & "'"
                strDateCondition1 = " U_Z_DocDate <= '" & dtTodate.ToString("yyyy-dd-MM") & "'"

            Else
                strDateCondition = " 1=1"
                strDateCondition1 = " 1=1"
            End If

            Dim oTestRs As SAPbobsCOM.Recordset
            Dim s As String
            Try
                s = "Select * from [@Z_OEXP] T0 where T0.U_Z_DocDate > ='2011-01-04' and T0.U_Z_DocDate  <='2011-30-04'"
                oTestRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTestRs.DoQuery(s)
                strDateCondition = strDateCondition1
            Catch ex As Exception
                s = "Select * from [@Z_OTIM] where " & strDateCondition
                oTestRs.DoQuery(s)
            End Try

            Dim strCondition, strGroupcondition, st1, st2, strJEPostingCondition As String
            strCondition = strEMPCondition & " and " & strDateCondition & " and " & strProjectCondition & " Order by U_Z_EMPCODE,U_Z_DocDate,U_Z_PRJCODE"
            strJEPostingCondition = strEMPCondition & " and " & strDateCondition & " and " & strProjectCondition
            st2 = strEMPCondition & " and " & strDateCondition & " and " & strProjectCondition
            strGroupcondition = strEMPCondition & " and " & strDateCondition & " and " & strProjectCondition & "  group by T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], T1.[U_Z_PRJCODE], T1.[U_Z_PRCNAME]  Order by U_Z_EMPCODE,U_Z_PRJCODE"
            st1 = strEMPCondition & " and " & strDateCondition & " and " & strProjectCondition & "  group by T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], T1.[U_Z_PRJCODE], T1.[U_Z_ACTNAME]  Order by U_Z_EMPCODE,U_Z_PRJCODE"

            'DatabindSummary(aForm)
            aForm.Freeze(False)
            Return strJEPostingCondition
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Function




#Region "Populate project Hours"

    Private Sub PouplateProjectHours(ByVal aGrid As SAPbouiCOM.Grid)
        Dim oTemp, RecActual As SAPbobsCOM.Recordset
        Dim strSql, strProject, strProcess, strActivity As String
        Dim dblEstimated, dblActual, dblVariance As Double
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            strProject = oGrid.DataTable.GetValue(0, intRow)
            If strProject <> "" Then
                strProcess = oGrid.DataTable.GetValue(2, intRow)
                strActivity = oGrid.DataTable.GetValue(3, intRow)
                dblActual = oGrid.DataTable.GetValue(5, intRow)
                strSql = "SELECT T0.[U_Z_PRJCODE],T0.[U_Z_PRJNAME], T0.[U_Z_PRJNAME], T1.[U_Z_MODNAME], T1.[U_Z_ACTNAME], T1.[U_Z_DAYS], T1.[U_Z_HOURS] FROM [dbo].[@Z_HPRJ]  T0  inner join  [dbo].[@Z_PRJ1]  T1 on T0.DocEntry=T1.DocEntry"
                strSql = strSql & " where T0.[U_Z_PRJCODE]='" & strProject & "' and T1.[U_Z_MODNAME]='" & strProcess & "' and T1.[U_Z_ACTNAME]='" & strActivity & "'"
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp.DoQuery(strSql)
                If oTemp.RecordCount > 0 Then
                    dblEstimated = oTemp.Fields.Item(6).Value
                    dblVariance = dblEstimated - dblActual
                    oGrid.DataTable.SetValue(1, intRow, oTemp.Fields.Item(1).Value)
                    oGrid.DataTable.SetValue(4, intRow, dblEstimated)
                    oGrid.DataTable.SetValue(6, intRow, dblVariance)
                End If

            End If
        Next
    End Sub

    Private Sub DisplayProjectwiseReport(ByVal aGrid As SAPbouiCOM.Grid, ByVal strCondition As String)
        Dim oTemp, RecActual As SAPbobsCOM.Recordset
        Dim strSql, strProject, strProcess, strActivity As String
        Dim dblEstimated, dblActual, dblVariance As Double
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            strProject = oGrid.DataTable.GetValue(0, intRow)
            If strProject <> "" Then
                strProcess = oGrid.DataTable.GetValue(2, intRow)
                strActivity = oGrid.DataTable.GetValue(3, intRow)
                ' dblEstimated = oGrid.DataTable.GetValue(4, intRow)
                dblEstimated = oGrid.DataTable.GetValue(6, intRow)
                ' strSql = "SELECT T0.[U_Z_PRJCODE],T0.[U_Z_PRJNAME], T0.[U_Z_PRJNAME], T1.[U_Z_MODNAME], T1.[U_Z_ACTNAME], T1.[U_Z_DAYS], T1.[U_Z_HOURS] FROM [dbo].[@Z_HPRJ]  T0  inner join  [dbo].[@Z_PRJ1]  T1 on T0.DocEntry=T1.DocEntry"
                'strSql = strSql & " where T0.[U_Z_PRJCODE]='" & strProject & "' and T1.[U_Z_MODNAME]='" & strProcess & "' and T1.[U_Z_ACTNAME]='" & strActivity & "'"
                strSql = "SELECT  T1.[U_Z_PRJCODE],T1.[U_Z_PRJCODE], T1.[U_Z_PRCNAME], T1.[U_Z_ACTNAME],Sum(T1.U_Z_HOURS)  FROM [dbo].[@Z_OTIM]  T0  inner join  [dbo].[@Z_TIM1]  T1 on T0.Code=T1.U_Z_REFCODE and T1.[U_Z_APPROVED]='A'   where " & strCondition & " and T1.[U_Z_PRJCODE]='" & strProject.Replace("'", "''") & "' and T1.[U_Z_PRCNAME]='" & strProcess.Replace("'", "''") & "' and T1.[U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' group by T1.U_Z_PRJCODE,T1.U_Z_PRCNAME,T1.U_Z_ACTNAME "
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp.DoQuery(strSql)
                If oTemp.RecordCount > 0 Then
                    ' dblActual = oTemp.Fields.Item(4).Value
                    dblActual = oTemp.Fields.Item(4).Value
                    dblVariance = dblEstimated - dblActual
                    ' oGrid.DataTable.SetValue(1, intRow, oTemp.Fields.Item(1).Value)
                Else
                    dblActual = 0
                    dblVariance = dblEstimated - dblActual
                End If
                '  oGrid.DataTable.SetValue(5, intRow, (dblActual))
                ' oGrid.DataTable.SetValue(6, intRow, (dblVariance))

                oGrid.DataTable.SetValue(7, intRow, (dblActual))
                oGrid.DataTable.SetValue(8, intRow, (dblVariance))

                ' strSql = "SELECT   U_Z_ACTNAME,SUm(LineTOtal),Project from PCH1 where  U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' group by PROJECT,U_Z_ACTNAME "


                strSql = "select sum(LineTotal) from ("
                strSql = strSql & " select LineTotal from PCH1 where  baseType<>20 and     U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                strSql = strSql & " select LineTotal 'LineTotal' from PDN1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                strSql = strSql & " select LineTotal *-1 'LineTotal' from RPD1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                'strSql = strSql & " select Debit-Credit 'LineTotal' from JDT1 T0 inner Join OJDT T1 on T1.TransId=T0.TransId where T1.TransType=30 and T0.U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  T0.[PROJECT]='" & strProject.Replace("'", "''") & "' and  T0.[U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                strSql = strSql & " select LineTotal *-1 'LineTotal' from RPC1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' ) as x"


                Dim dblJDT As Double
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp.DoQuery(strSql)

                'dblEstimated = oGrid.DataTable.GetValue(7, intRow)
                dblEstimated = oGrid.DataTable.GetValue(9, intRow)

                If oTemp.RecordCount > 0 Then
                    ' dblActual = oTemp.Fields.Item(1).Value
                    dblActual = oTemp.Fields.Item(0).Value
                    '   dblVariance = dblEstimated - dblActual
                    ' oGrid.DataTable.SetValue(1, intRow, oTemp.Fields.Item(1).Value)
                Else
                    dblActual = 0
                    '  dblVariance = dblEstimated - dblActual
                End If



                strSql = " SELECT U_Z_MDNAME, U_Z_ACTNAME , SUm(Debit-Credit),Project from JDT1 where  U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' group by PROJECT,U_Z_MDNAME,U_Z_ACTNAME "

                oTemp.DoQuery(strSql)
                If oTemp.RecordCount > 0 Then
                    dblActual = dblActual + oTemp.Fields.Item(2).Value
                Else
                    dblActual = dblActual
                End If
                dblVariance = dblEstimated - dblActual
                'oGrid.DataTable.SetValue(8, intRow, (dblActual))
                ' oGrid.DataTable.SetValue(9, intRow, (dblVariance))
                oGrid.DataTable.SetValue(10, intRow, (dblActual))
                oGrid.DataTable.SetValue(11, intRow, (dblVariance))


                strSql = "select sum(LineTotal) from ("
                strSql = strSql & " select LineTotal from INV1 where     U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                '  strSql = strSql & " select LineTotal 'LineTotal' from PDN1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                'strSql = strSql & " select LineTotal *-1 'LineTotal' from RIN1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                'strSql = strSql & " select Debit-Credit 'LineTotal' from JDT1 T0 inner Join OJDT T1 on T1.TransId=T0.TransId where T1.TransType=30 and T0.U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  T0.[PROJECT]='" & strProject.Replace("'", "''") & "' and  T0.[U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                strSql = strSql & " select LineTotal *-1 'LineTotal' from RIN1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' ) as x"
                oTemp.DoQuery(strSql)
                If oTemp.RecordCount > 0 Then
                    dblActual = oTemp.Fields.Item(0).Value
                Else
                    dblActual = 0
                End If
                oGrid.DataTable.SetValue("InvoicedAmt", intRow, dblActual)

            End If
        Next
    End Sub

    Private Sub DisplayProjectwiseReport_Material(ByVal aGrid As SAPbouiCOM.Grid, ByVal strCondition As String)
        Dim oTemp, RecActual As SAPbobsCOM.Recordset
        Dim strSql, strProject, strProcess, strActivity, strItemCode As String
        Dim dblEstimated, dblActual, dblVariance As Double
        Dim dblMaterialOrdered, dblMaterialReceived, dblOrderPending, dblMaterialPending, dblCostVariance, dblestimatedcost As Double
        Dim dblOrdQty, dblOrdCost, dblRecQty, dblRecCost, dblReleaseQty, dblReleaseCost As Double
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            ' strProject = oGrid.DataTable.GetValue(0, intRow)
            strProject = oGrid.DataTable.GetValue("Project", intRow)
            If strProject <> "" Then
                strProcess = oGrid.DataTable.GetValue("Business", intRow)
                strActivity = oGrid.DataTable.GetValue("Activity", intRow)
                ' dblEstimated = oGrid.DataTable.GetValue(4, intRow)
                dblEstimated = oGrid.DataTable.GetValue("Qty", intRow)
                dblestimatedcost = oGrid.DataTable.GetValue("Cost", intRow)
                strItemCode = oGrid.DataTable.GetValue("ItemCode", intRow)
                ' strSql = "SELECT T0.[U_Z_PRJCODE],T0.[U_Z_PRJNAME], T0.[U_Z_PRJNAME], T1.[U_Z_MODNAME], T1.[U_Z_ACTNAME], T1.[U_Z_DAYS], T1.[U_Z_HOURS] FROM [dbo].[@Z_HPRJ]  T0  inner join  [dbo].[@Z_PRJ1]  T1 on T0.DocEntry=T1.DocEntry"
                'strSql = strSql & " where T0.[U_Z_PRJCODE]='" & strProject & "' and T1.[U_Z_MODNAME]='" & strProcess & "' and T1.[U_Z_ACTNAME]='" & strActivity & "'"

                'strSql = "SELECT   U_Z_ACTNAME,SUm(LineTOtal),Project from POR1 where  U_Z_MDNAME='" & strProcess & "' and  [PROJECT]='" & strProject & "' and  [U_Z_ACTNAME]='" & strActivity & "' group by PROJECT,U_Z_ACTNAME "
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oTemp.DoQuery(strSql)
                Dim strBOQRef As String
                strSql = "Select isnull(T0.U_Z_BOQ,''),* from [@Z_PRJ1] T0 inner Join [@Z_HPRJ] T1 on T0.DocEntry=T1.DocEntry where T1.U_Z_PrjCode='" & strProject & "' and T0.U_Z_MODNAME='" & strProcess.Replace("'", "''") & "' and U_Z_Actname='" & strActivity.Replace("'", "''") & "'"
                oTemp.DoQuery(strSql)
                If oTemp.RecordCount > 0 Then
                    strBOQRef = oTemp.Fields.Item(0).Value
                Else
                    strBOQRef = ""
                End If
                'dblEstimated = oGrid.DataTable.GetValue(7, intRow)
                dblEstimated = oGrid.DataTable.GetValue("Cost", intRow)

                If strBOQRef <> "" Then
                    strSql = "Select sum(LineTotal),sum(Quantity) from POR1 T0 inner Join OPOR T1 on T1.DocEntry=T0.DocEntry and T1.CANCELED <>'Y' where  T0.ItemCode='" & strItemCode & "' and  T0.U_Z_BOQREF='" & strBOQRef & "'"
                    oTemp.DoQuery(strSql)
                    If oTemp.RecordCount > 0 Then
                        dblOrdCost = oTemp.Fields.Item(0).Value
                        dblOrdQty = oTemp.Fields.Item(1).Value
                    Else
                        dblOrdCost = 0
                        dblOrdQty = 0
                        ' dblVariance = dblEstimated - dblActual
                    End If
                    oGrid.DataTable.SetValue("OrdQty", intRow, dblOrdQty)
                    oGrid.DataTable.SetValue("OrdCost", intRow, dblOrdCost)

                    strSql = "select sum(LineTotal),sum(Quantity) from ("
                    strSql = strSql & " select Quantity 'Quantity',LineTotal 'LineTotal' from PCH1 where  baseType<>20 and Itemcode='" & strItemCode & "' and isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                    strSql = strSql & " select Quantity  'Quantity',LineTotal 'LineTotal' from PDN1 where Itemcode='" & strItemCode & "' and isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                    strSql = strSql & " select Quantity *-1  'Quantity',LineTotal *-1 'LineTotal' from RPD1 where Itemcode='" & strItemCode & "' and isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                    strSql = strSql & " select Quantity *-1  'Quantity',LineTotal *-1 'LineTotal' from RPC1 where Itemcode='" & strItemCode & "' and isnull(U_Z_BOQREF,'')='" & strBOQRef & "' ) as x"
                    oTemp.DoQuery(strSql)

                    If oTemp.RecordCount > 0 Then
                        dblRecQty = oTemp.Fields.Item(1).Value
                        dblRecCost = oTemp.Fields.Item(0).Value
                    Else
                        dblRecCost = 0
                        dblRecQty = 0
                    End If
                    oGrid.DataTable.SetValue("RecQty", intRow, dblRecQty)
                    oGrid.DataTable.SetValue("RecCost", intRow, dblRecCost)

                    strSql = "select sum(LineTotal),sum(Quantity) from ("
                    strSql = strSql & " select Quantity 'Quantity',LineTotal 'LineTotal' from IGE1 where Itemcode='" & strItemCode & "' and isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                    strSql = strSql & " select Quantity *-1  'Quantity',LineTotal *-1 'LineTotal' from IGN1 where Itemcode='" & strItemCode & "' and isnull(U_Z_BOQREF,'')='" & strBOQRef & "' ) as x"
                    oTemp.DoQuery(strSql)
                    If oTemp.RecordCount > 0 Then
                        dblRecQty = oTemp.Fields.Item(1).Value
                        dblRecCost = oTemp.Fields.Item(0).Value
                    Else
                        dblRecQty = 0
                        dblRecCost = 0
                        ' dblVariance = dblEstimated - dblActual
                    End If
                    oGrid.DataTable.SetValue("RelQty", intRow, dblRecQty)
                    oGrid.DataTable.SetValue("RelCost", intRow, dblRecCost)
                End If

            End If
        Next
    End Sub

    Private Sub DisplayProjectwiseReport_Summary(ByVal aGrid As SAPbouiCOM.Grid, ByVal strCondition As String)
        Dim oTemp, RecActual As SAPbobsCOM.Recordset
        Dim strSql, strProject, strProcess, strActivity As String
        Dim dblEstimated, dblActual, dblVariance As Double
        Dim dblMaterialOrdered, dblMaterialReceived, dblOrderPending, dblMaterialPending, dblCostVariance, dblestimatedcost As Double
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            ' strProject = oGrid.DataTable.GetValue(0, intRow)
            strProject = oGrid.DataTable.GetValue("Project", intRow)
            If strProject <> "" Then
                strProcess = oGrid.DataTable.GetValue("Business", intRow)
                strActivity = oGrid.DataTable.GetValue("Activity", intRow)
                ' dblEstimated = oGrid.DataTable.GetValue(4, intRow)
                dblEstimated = oGrid.DataTable.GetValue("EstimatedHours", intRow)
                dblestimatedcost = oGrid.DataTable.GetValue("EstimatedCost", intRow)
                If strProcess <> "Expenses" Then
                    strSql = "SELECT  T1.[U_Z_PRJCODE],T1.[U_Z_PRJCODE], T1.[U_Z_PRCNAME], T1.[U_Z_ACTNAME],Sum(T1.U_Z_HOURS),sum(T1.U_Z_Quantity)  FROM [dbo].[@Z_OTIM]  T0  inner join  [dbo].[@Z_TIM1]  T1 on T0.Code=T1.U_Z_REFCODE and T1.[U_Z_APPROVED]='A'   where " & strCondition & " and T1.[U_Z_PRJCODE]='" & strProject.Replace("'", "''") & "' and T1.[U_Z_PRCNAME]='" & strProcess.Replace("'", "''") & "' and T1.[U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' group by T1.U_Z_PRJCODE,T1.U_Z_PRCNAME,T1.U_Z_ACTNAME "
                    '  strSql = "SELECT  T0.[Project],T0.[Project], 0,0,Sum(T0.Debit),0  FROM [dbo].[JDT1]  T0   where  T0.[Project]='" & strProject & "' group by T0.Project"
                Else
                    strSql = "SELECT  T0.[Project],T0.[Project], 0,0,Sum(T0.Debit),0  FROM [dbo].[JDT1]  T0   where  T0.[Project]='" & strProject.Replace("'", "''") & "' group by T0.Project"
                End If

                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp.DoQuery(strSql)
                Dim dblEstiamtedQty, dblActualQty As Double
                dblEstiamtedQty = oGrid.DataTable.GetValue("Qty", intRow)

                If oTemp.RecordCount > 0 Then
                    ' dblActual = oTemp.Fields.Item(4).Value
                    dblActual = oTemp.Fields.Item(4).Value
                    dblVariance = dblEstimated - dblActual
                    dblActualQty = oTemp.Fields.Item(5).Value
                    ' oGrid.DataTable.SetValue(1, intRow, oTemp.Fields.Item(1).Value)
                Else
                    dblActual = 0
                    dblActualQty = 0
                    dblVariance = dblEstimated - dblActual
                End If
                If dblActualQty >= 0 Then
                    dblEstiamtedQty = dblEstiamtedQty * dblActualQty / 100
                End If
                oGrid.DataTable.SetValue("Percentage", intRow, dblEstiamtedQty)
                '  oGrid.DataTable.SetValue(5, intRow, (dblActual))
                ' oGrid.DataTable.SetValue(6, intRow, (dblVariance))

                'oGrid.DataTable.SetValue(7, intRow, (dblActual))
                'oGrid.DataTable.SetValue(8, intRow, (dblVariance))

                oGrid.DataTable.SetValue("ActualHours", intRow, (dblActual))
                oGrid.DataTable.SetValue("VarianceHours", intRow, (dblVariance))

                strSql = "SELECT   U_Z_ACTNAME,SUm(LineTOtal),Project from PCH1 where  U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' group by PROJECT,U_Z_ACTNAME "
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp.DoQuery(strSql)
                Dim strBOQRef As String
                strSql = "Select isnull(T0.U_Z_BOQ,''),* from [@Z_PRJ1] T0 inner Join [@Z_HPRJ] T1 on T0.DocEntry=T1.DocEntry where T1.U_Z_PrjCode='" & strProject.Replace("'", "''") & "' and T0.U_Z_MODNAME='" & strProcess.Replace("'", "''") & "' and U_Z_Actname='" & strActivity.Replace("'", "''") & "'"
                oTemp.DoQuery(strSql)
                If oTemp.RecordCount > 0 Then
                    strBOQRef = oTemp.Fields.Item(0).Value
                Else
                    strBOQRef = ""
                End If
                'dblEstimated = oGrid.DataTable.GetValue(7, intRow)
                dblEstimated = oGrid.DataTable.GetValue("EstimatedCost", intRow)

                ' strBOQRef = ""
                If strBOQRef <> "" Then
                    strSql = "Select sum(LineTotal) from POR1 T0 inner Join OPOR T1 on T1.DocEntry=T0.DocEntry and T1.CANCELED <>'Y' where T0.U_Z_BOQREF='" & strBOQRef & "'"
                    'strSql = "Select sum(LineTotal) from POR1 T0 inner Join OPOR T1 on T1.DocEntry=T0.DocEntry and T1.CANCELED <>'Y' where T0.U_Z_BOQREF='" & strBOQRef & "'"
                    oTemp.DoQuery(strSql)
                    If oTemp.RecordCount > 0 Then
                        dblMaterialOrdered = oTemp.Fields.Item(0).Value
                    Else
                        dblMaterialOrdered = 0
                        ' dblVariance = dblEstimated - dblActual
                    End If
                    oGrid.DataTable.SetValue("MatrialOrdered", intRow, dblMaterialOrdered)
                    strSql = "select sum(LineTotal) from ("
                    strSql = strSql & " select LineTotal from PCH1 where  baseType<>20 and  isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                    strSql = strSql & " select LineTotal 'LineTotal' from PDN1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                    strSql = strSql & " select LineTotal *-1 'LineTotal' from RPD1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                    strSql = strSql & " select Debit-Credit 'LineTotal' from JDT1 T0 inner Join OJDT T1 on T1.TransId=T0.TransId where T1.TransType=30 and isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                    strSql = strSql & " select LineTotal *-1 'LineTotal' from RPC1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' ) as x"
                    oTemp.DoQuery(strSql)

                    If oTemp.RecordCount > 0 Then
                        dblMaterialReceived = oTemp.Fields.Item(0).Value
                    Else
                        dblMaterialReceived = 0
                        ' dblVariance = dblEstimated - dblActual
                    End If
                    oGrid.DataTable.SetValue("MaterialReceived", intRow, dblMaterialReceived)



                    'strSql = "Select sum(LineTotal) from POR1 T0 inner Join OPOR T1 on T1.DocEntry=T0.DocEntry and T1.CANCELED <>'Y' where T0.U_Z_BOQREF='" & strBOQRef & "'"
                    'oTemp.DoQuery(strSql)
                    'If oTemp.RecordCount > 0 Then
                    '    dblOrderPending = oTemp.Fields.Item(0).Value
                    'Else
                    '    dblOrderPending = 0
                    '    ' dblVariance = dblEstimated - dblActual
                    'End If
                    dblOrderPending = dblestimatedcost - dblMaterialOrdered
                    oGrid.DataTable.SetValue("OrderPending", intRow, dblOrderPending)
                    strSql = "select sum(LineTotal) from ("
                    strSql = strSql & " select LineTotal from POR1 T0 inner Join OPOR T1  on T1.DocEntry=T0.DocEntry and T1.DocStatus='O' where isnull(U_Z_BOQREF,'')='" & strBOQRef & "'"
                    'strSql = strSql & " select LineTotal from PDN1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                    'strSql = strSql & " select LineTotal *-1 from RPD1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                    'strSql = strSql & " select LineTotal *-1 from RPC1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "'
                    strSql = strSql & " ) as x"
                    oTemp.DoQuery(strSql)

                    If oTemp.RecordCount > 0 Then
                        dblMaterialPending = oTemp.Fields.Item(0).Value
                    Else
                        dblMaterialPending = 0
                        ' dblVariance = dblEstimated - dblActual
                    End If
                    oGrid.DataTable.SetValue("MaterialPending", intRow, dblMaterialPending)
                    dblCostVariance = dblestimatedcost - dblMaterialReceived
                    oGrid.DataTable.SetValue("Cost Variance", intRow, dblCostVariance)


                    strSql = "select sum(LineTotal) from ("
                    strSql = strSql & " select LineTotal from inv1 where    isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                    '   strSql = strSql & " select LineTotal 'LineTotal' from PDN1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                    '  strSql = strSql & " select LineTotal *-1 'LineTotal' from RPD1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                    '   strSql = strSql & " select Debit-Credit 'LineTotal' from JDT1 T0 inner Join OJDT T1 on T1.TransId=T0.TransId where T1.TransType=30 and isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                    strSql = strSql & " select LineTotal *-1 'LineTotal' from RIN1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' ) as x"
                    oTemp.DoQuery(strSql)

                    If oTemp.RecordCount > 0 Then
                        dblMaterialReceived = oTemp.Fields.Item(0).Value
                    Else
                        dblMaterialReceived = 0
                        ' dblVariance = dblEstimated - dblActual
                    End If
                    oGrid.DataTable.SetValue("ReceivedAmt", intRow, dblMaterialReceived)

                Else
                    strSql = "Select sum(LineTotal) from POR1 T0 inner Join OPOR T1 on T1.DocEntry=T0.DocEntry and T1.CANCELED <>'Y' where    U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  T0.Project='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' group by T0.PROJECT,U_Z_ACTNAME "
                    oTemp.DoQuery(strSql)
                    If oTemp.RecordCount > 0 Then
                        dblMaterialOrdered = oTemp.Fields.Item(0).Value
                    Else
                        dblMaterialOrdered = 0
                        ' dblVariance = dblEstimated - dblActual
                    End If
                    oGrid.DataTable.SetValue("MatrialOrdered", intRow, dblMaterialOrdered)
                    strSql = "select sum(LineTotal) from ("
                    strSql = strSql & " select LineTotal from PCH1 where  baseType<>20 and     U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                    strSql = strSql & " select LineTotal 'LineTotal' from PDN1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                    strSql = strSql & " select LineTotal *-1 'LineTotal' from RPD1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                    strSql = strSql & " select Debit-Credit 'LineTotal' from JDT1 T0 inner Join OJDT T1 on T1.TransId=T0.TransId where T1.TransType=30 and T0.U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  T0.[PROJECT]='" & strProject.Replace("'", "''") & "' and  T0.[U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                    strSql = strSql & " select LineTotal *-1 'LineTotal' from RPC1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' ) as x"
                    oTemp.DoQuery(strSql)

                    If oTemp.RecordCount > 0 Then
                        dblMaterialReceived = oTemp.Fields.Item(0).Value
                    Else
                        dblMaterialReceived = 0
                        ' dblVariance = dblEstimated - dblActual
                    End If
                    oGrid.DataTable.SetValue("MaterialReceived", intRow, dblMaterialReceived)



                    'strSql = "Select sum(LineTotal) from POR1 T0 inner Join OPOR T1 on T1.DocEntry=T0.DocEntry and T1.CANCELED <>'Y' where T0.U_Z_BOQREF='" & strBOQRef & "'"
                    'oTemp.DoQuery(strSql)
                    'If oTemp.RecordCount > 0 Then
                    '    dblOrderPending = oTemp.Fields.Item(0).Value
                    'Else
                    '    dblOrderPending = 0
                    '    ' dblVariance = dblEstimated - dblActual
                    'End If
                    dblOrderPending = dblestimatedcost - dblMaterialOrdered
                    oGrid.DataTable.SetValue("OrderPending", intRow, dblOrderPending)
                    strSql = "select sum(LineTotal) from ("
                    strSql = strSql & " select LineTotal from POR1 T0 inner Join OPOR T1  on T1.DocEntry=T0.DocEntry and T1.DocStatus='O' where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  T0.[PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "'"
                    'strSql = strSql & " select LineTotal from PDN1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                    'strSql = strSql & " select LineTotal *-1 from RPD1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                    'strSql = strSql & " select LineTotal *-1 from RPC1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "'
                    strSql = strSql & " ) as x"
                    oTemp.DoQuery(strSql)

                    If oTemp.RecordCount > 0 Then
                        dblMaterialPending = oTemp.Fields.Item(0).Value
                    Else
                        dblMaterialPending = 0
                        ' dblVariance = dblEstimated - dblActual
                    End If
                    oGrid.DataTable.SetValue("MaterialPending", intRow, dblMaterialPending)
                    dblCostVariance = dblestimatedcost - dblMaterialReceived
                    oGrid.DataTable.SetValue("Cost Variance", intRow, dblCostVariance)

                    strSql = "select sum(LineTotal) from ("
                    strSql = strSql & " select LineTotal from INV1 where     U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                    '   strSql = strSql & " select LineTotal 'LineTotal' from PDN1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                    '  strSql = strSql & " select LineTotal *-1 'LineTotal' from RPD1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                    '  strSql = strSql & " select Debit-Credit 'LineTotal' from JDT1 T0 inner Join OJDT T1 on T1.TransId=T0.TransId where T1.TransType=30 and T0.U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  T0.[PROJECT]='" & strProject.Replace("'", "''") & "' and  T0.[U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                    strSql = strSql & " select LineTotal *-1 'LineTotal' from RIN1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' ) as x"
                    oTemp.DoQuery(strSql)

                    If oTemp.RecordCount > 0 Then
                        dblMaterialReceived = oTemp.Fields.Item(0).Value
                    Else
                        dblMaterialReceived = 0
                        ' dblVariance = dblEstimated - dblActual
                    End If
                    oGrid.DataTable.SetValue("ReceivedAmt", intRow, dblMaterialReceived)
                End If


            End If
        Next
    End Sub
#End Region

#Region "Get Columns"
    Private Function CreateColumns(ByVal aGrid As SAPbouiCOM.Grid, ByVal aForm As SAPbouiCOM.Form) As SAPbouiCOM.DataTable
        Dim dtGridColumn As SAPbouiCOM.DataTable
        Dim oRS As SAPbobsCOM.Recordset
        oRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim strPrfrom, strPrTo As String
        oCombo = aForm.Items.Item("6").Specific
        strPrfrom = oCombo.Selected.Value
        oCombo = aForm.Items.Item("15").Specific
        strPrTo = oCombo.Selected.Value

        Dim st, cond As String
        If strPrfrom <> "" Then
            cond = "T1.U_Z_PRJCODE >='" & strPrfrom & "'"
        Else
            cond = "1=1"
        End If

        If strPrTo <> "" Then
            cond = cond & " and T1.U_Z_PRJCODE <='" & strPrTo & "'"
        Else
            cond = cond & " and 1=1"
        End If
        st = "  select U_Z_ACTNAME,COUNT(*)  from [@Z_PRJ1] T0 inner Join [@Z_HPRJ] T1 on T1.DocEntry =t0.DocEntry  where " & cond & "  group by T0.U_Z_ACTNAME "
        oRS.DoQuery(st)
        '   oRS.DoQuery("Select U_Z_ActName from [@Z_Activity] order by Code")
        dtGridColumn = aGrid.DataTable
        dtGridColumn.Clear()
        ' dtGridColumn.Columns.Add("EmpId", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtGridColumn.Columns.Add("PrjCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtGridColumn.Columns.Add("PrjName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtGridColumn.Columns.Add("PrcName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)

        'dtGridColumn.Columns.Add("Start", SAPbouiCOM.BoFieldsType.ft_Date)
        'dtGridColumn.Columns.Add("End", SAPbouiCOM.BoFieldsType.ft_Date)
        dtGridColumn.Columns.Add("Hours", SAPbouiCOM.BoFieldsType.ft_Float)
        For intRow As Integer = 0 To oRS.RecordCount - 1
            Try
                dtGridColumn.Columns.Add(oRS.Fields.Item(0).Value, SAPbouiCOM.BoFieldsType.ft_Float)
            Catch ex As Exception
            End Try
            oRS.MoveNext()
        Next
        Return dtGridColumn
    End Function
#End Region

#Region "FormatGrids"
    Private Sub FormatGrids(ByVal aForm As SAPbouiCOM.Form, ByVal aChoice As String, ByVal strQuery As String, Optional ByVal strQuery1 As String = "", Optional ByVal strQuery2 As String = "")

        'oGrid = aForm.Items.Item("17").Specific
        'oGrid.DataTable = CreateColumns(oGrid)
        Dim aGrid As SAPbouiCOM.Grid
        'aGrid = oGrid
        If aChoice.ToString = "LEV" Then
            oGrid = aForm.Items.Item("7").Specific
            oGrid.DataTable.ExecuteQuery(strQuery)

            '   SELECT T0.[Code], T0.[Name], T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], T0.[U_Z_DOCDATE], T0.[U_Z_SUBEMP], T0.[U_Z_TYPE], T0.[U_Z_FROMDATE], 
            'T0.[U_Z_TODATE], T0.[U_Z_DAYS], T0.[U_Z_REASON], T0.[U_Z_APPROVED], T0.[U_Z_REMARKS] FROM [dbo].[@Z_OLEV]  T0
            aGrid = oGrid
            aGrid.Columns.Item("Code").Visible = False
            aGrid.Columns.Item("Name").Visible = False
            aGrid.Columns.Item("U_Z_EMPCODE").Visible = True
            aGrid.Columns.Item("U_Z_EMPCODE").TitleObject.Caption = "Employee Code"
            aGrid.Columns.Item("U_Z_EMPCODE").Editable = False
            oEditTextColumn = aGrid.Columns.Item("U_Z_EMPCODE")
            oEditTextColumn.LinkedObjectType = "171"
            aGrid.Columns.Item("U_Z_EMPNAME").Visible = True
            aGrid.Columns.Item("U_Z_EMPNAME").TitleObject.Caption = "Employee Name"
            aGrid.Columns.Item("U_Z_EMPNAME").Editable = False
            aGrid.Columns.Item("U_Z_DOCDATE").TitleObject.Caption = "Requested Date"
            aGrid.Columns.Item("U_Z_DOCDATE").Editable = False
            aGrid.Columns.Item("U_Z_TYPE").TitleObject.Caption = "Leave Type"
            aGrid.Columns.Item("U_Z_TYPE").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            Dim otemp As SAPbobsCOM.Recordset
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("Select U_Z_Name,U_Z_Name from [@Z_LeaveType] ")
            oCombobox = aGrid.Columns.Item("U_Z_TYPE")
            oCombobox.ValidValues.Add("", "")
            For intRow As Integer = 0 To otemp.RecordCount - 1
                oCombobox.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both


            aGrid.Columns.Item(5).Editable = False

            '   SELECT T0.[Code], T0.[Name], T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], T0.[U_Z_DOCDATE], T0.[U_Z_SUBEMP], T0.[U_Z_TYPE], T0.[U_Z_FROMDATE], 
            'T0.[U_Z_TODATE], T0.[U_Z_DAYS], T0.[U_Z_REASON], T0.[U_Z_APPROVED], T0.[U_Z_REMARKS] FROM [dbo].[@Z_OLEV]  T0

            aGrid.Columns.Item("U_Z_FROMDATE").TitleObject.Caption = "From Date"
            aGrid.Columns.Item("U_Z_FROMDATE").Editable = False
            aGrid.Columns.Item("U_Z_TODATE").TitleObject.Caption = "To date"
            aGrid.Columns.Item("U_Z_TODATE").Editable = False
            aGrid.Columns.Item("U_Z_DAYS").TitleObject.Caption = "Number of Days"
            aGrid.Columns.Item("U_Z_DAYS").Editable = False
            aGrid.Columns.Item("U_Z_DAYS").Editable = False
            aGrid.Columns.Item("U_Z_REASON").TitleObject.Caption = "Reason"
            aGrid.Columns.Item("U_Z_REASON").Editable = False
            aGrid.Columns.Item("U_Z_APPROVED").TitleObject.Caption = "Approved"
            aGrid.Columns.Item("U_Z_APPROVED").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oCombobox = aGrid.Columns.Item("U_Z_APPROVED")
            oCombobox.ValidValues.Add("P", "Pending")
            oCombobox.ValidValues.Add("A", "Approved")
            oCombobox.ValidValues.Add("D", "Declined")
            oCombobox.TitleObject.Caption = "Approval Status"
            oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
            aGrid.Columns.Item("U_Z_APPROVED").Editable = False
            aGrid.Columns.Item("U_Z_REMARKS").TitleObject.Caption = "Remarks"
            aGrid.Columns.Item("U_Z_REMARKS").Editable = False
            aGrid.Columns.Item("U_Z_SUBEMP").TitleObject.Caption = "SubLevel EmployeeName"
            aGrid.Columns.Item("U_Z_SUBEMP").Editable = False

            aGrid.AutoResizeColumns()
            aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            aForm.PaneLevel = 2
            aForm.Items.Item("7").Enabled = False
            aForm.Items.Item("29").Visible = False
            aForm.Items.Item("30").Visible = False
            aForm.Items.Item("34").Visible = False
            aForm.Items.Item("35").Visible = False
        ElseIf aChoice.ToUpper = "EXP" Then
            oGrid = aForm.Items.Item("7").Specific
            oGrid.DataTable.ExecuteQuery(strQuery)
            oGrid.Columns.Item(0).TitleObject.Caption = "Employee Code"
            oEditTextColumn = oGrid.Columns.Item(0)
            oEditTextColumn.LinkedObjectType = "171"
            oGrid.Columns.Item(0).TitleObject.Sortable = True
            oGrid.Columns.Item(0).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(1).TitleObject.Caption = "Employee Name"
            oGrid.Columns.Item(1).TitleObject.Sortable = True
            oGrid.Columns.Item(1).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(2).TitleObject.Caption = "Document Date"
            oGrid.Columns.Item(2).TitleObject.Sortable = True
            oGrid.Columns.Item(2).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(3).TitleObject.Caption = "Expenses Name"
            oGrid.Columns.Item(3).TitleObject.Sortable = True
            oGrid.Columns.Item(3).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(4).TitleObject.Caption = "Expenses Type"
            oGrid.Columns.Item(4).TitleObject.Sortable = True
            oGrid.Columns.Item(4).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(4).Visible = False
            oGrid.Columns.Item(5).TitleObject.Caption = "Project Code"
            oEditTextColumn = oGrid.Columns.Item(5)
            oEditTextColumn.LinkedObjectType = "63"
            oGrid.Columns.Item(5).TitleObject.Sortable = True
            oGrid.Columns.Item(5).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(6).TitleObject.Caption = "Project Name"
            oGrid.Columns.Item(6).Editable = False

            oGrid.Columns.Item(9).TitleObject.Caption = "Currency"
            oGrid.Columns.Item(10).TitleObject.Caption = "Txn Currency Amount"
            oEditTextColumn = oGrid.Columns.Item(10)
            ' oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oGrid.Columns.Item(11).TitleObject.Caption = "Amount in Local Currency"
            oGrid.Columns.Item(12).TitleObject.Caption = "Amount in Local Currency"
            oEditTextColumn = oGrid.Columns.Item(12)
            oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oGrid.Columns.Item(11).Visible = False
            oGrid.Columns.Item(13).Visible = False
            oGrid.Columns.Item(14).TitleObject.Caption = "Remarks"
            oGrid.Columns.Item(15).TitleObject.Caption = "Select"
            oGrid.Columns.Item(16).TitleObject.Caption = "Approval Status"
            oGrid.Columns.Item(14).TitleObject.Sortable = True
            oGrid.Columns.Item(14).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item("U_Z_DocNum").TitleObject.Caption = "Journal Remarks"
            oEditTextColumn = oGrid.Columns.Item("U_Z_DocNum")
            oEditTextColumn.LinkedObjectType = "30"
            oGrid.Columns.Item("U_Z_MODNAME").TitleObject.Caption = "Phase"
            oGrid.Columns.Item("U_Z_MODNAME").Editable = False
            oGrid.Columns.Item("U_Z_ACTNAME").TitleObject.Caption = "Activity"
            oGrid.Columns.Item("U_Z_ACTNAME").Editable = False
            oGrid.AutoResizeColumns()
            aForm.PaneLevel = 2
            aForm.Items.Item("7").Enabled = False
            aForm.Items.Item("29").Visible = False
            aForm.Items.Item("30").Visible = False
            aForm.Items.Item("34").Visible = False
            aForm.Items.Item("35").Visible = False
        ElseIf aChoice.ToUpper = "POSTING" Then
            oGrid = aForm.Items.Item("7").Specific
            oGrid.DataTable.ExecuteQuery(strQuery)
            oGrid.Columns.Item(0).TitleObject.Caption = "Employee Code"
            oEditTextColumn = oGrid.Columns.Item(0)
            oEditTextColumn.LinkedObjectType = "171"
            ' oGrid.Columns.Item(0).TitleObject.Sortable = True
            oGrid.Columns.Item(0).Editable = False
            '   oGrid.Columns.Item(0).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(1).TitleObject.Caption = "Employee Name"
            oGrid.Columns.Item(1).Editable = False
            ' oGrid.Columns.Item(1).TitleObject.Sortable = True
            '  oGrid.Columns.Item(1).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(2).TitleObject.Caption = "Document Date"
            ' oGrid.Columns.Item(2).TitleObject.Sortable = True
            oGrid.Columns.Item(2).Editable = False
            ' oGrid.Columns.Item(2).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(3).TitleObject.Caption = "Expenses Name"
            '  oGrid.Columns.Item(3).TitleObject.Sortable = True
            oGrid.Columns.Item(3).Editable = False
            '  oGrid.Columns.Item(3).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(4).TitleObject.Caption = "Expenses Type"
            ' oGrid.Columns.Item(4).TitleObject.Sortable = True
            '  oGrid.Columns.Item(4).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(4).Visible = False
            oGrid.Columns.Item(4).Editable = False
            oGrid.Columns.Item(5).TitleObject.Caption = "Project Code"
            oEditTextColumn = oGrid.Columns.Item(5)
            oEditTextColumn.LinkedObjectType = "63"
            ' oGrid.Columns.Item(5).TitleObject.Sortable = True
            ' oGrid.Columns.Item(5).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(5).Editable = False
            oGrid.Columns.Item(6).TitleObject.Caption = "Project Name"
            oGrid.Columns.Item(6).Editable = False
            oGrid.Columns.Item(6).Editable = False
            oGrid.Columns.Item(7).TitleObject.Caption = "Currency"
            oGrid.Columns.Item(7).Editable = False
            oGrid.Columns.Item(8).TitleObject.Caption = "Txn Currency Amount"
            oEditTextColumn = oGrid.Columns.Item(8)
            ' oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oGrid.Columns.Item(8).Editable = False
            oGrid.Columns.Item(9).TitleObject.Caption = "Amount in Local Currency"
            oGrid.Columns.Item(9).Editable = False
            oGrid.Columns.Item(10).TitleObject.Caption = "Amount in Local Currency"
            oEditTextColumn = oGrid.Columns.Item(10)
            oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oGrid.Columns.Item(10).Editable = False
            oGrid.Columns.Item(9).Visible = False
            oGrid.Columns.Item(11).Visible = False
            oGrid.Columns.Item(12).TitleObject.Caption = "Remarks"
            oGrid.Columns.Item(12).Editable = False
            'oGrid.Columns.Item(13).TitleObject.Caption = "Approval Status"
            'oGrid.Columns.Item(13).TitleObject.Sortable = True
            'oGrid.Columns.Item(13).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(13).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item(13).Editable = True
            oGrid.Columns.Item(14).TitleObject.Caption = "Code"
            oGrid.Columns.Item(14).Editable = False
            oGrid.AutoResizeColumns()
            aForm.PaneLevel = 2
            '  aForm.Items.Item("7").Enabled = False
            aForm.Items.Item("21").Enabled = False
            aForm.Items.Item("29").Visible = True
            aForm.Items.Item("30").Visible = True
            aForm.Items.Item("34").Visible = True
            aForm.Items.Item("35").Visible = True
        ElseIf aChoice.ToUpper = "POSTJE" Then

            'SELECT  T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], convert(varchar(10),T0.[U_Z_DOCDATE],105),T1.[U_Z_EXPNAME], T1.[U_Z_PRJCODE],T1.[U_Z_PRJNAME],T3.[U_Z_ActCode],T4.[AcctName],T1.[U_Z_LocAmount]
            oGrid = aForm.Items.Item("7").Specific
            oGrid.DataTable.ExecuteQuery(strQuery)
            oGrid.Columns.Item(0).TitleObject.Caption = "Employee Code"
            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(1).TitleObject.Caption = "Employee Name"
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(2).TitleObject.Caption = "Document Date"
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(3).TitleObject.Caption = "Expense Name"
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(4).TitleObject.Caption = "Project Code"
            oGrid.Columns.Item(4).Editable = False
            oGrid.Columns.Item(5).TitleObject.Caption = "Project Name"
            oGrid.Columns.Item(5).Editable = False
            oGrid.Columns.Item(6).TitleObject.Caption = "Format Code"
            oGrid.Columns.Item(6).Editable = False
            oGrid.Columns.Item(7).TitleObject.Caption = "Account Name"
            oGrid.Columns.Item(7).Editable = False
            oGrid.Columns.Item(8).TitleObject.Caption = "Amount"
            oGrid.Columns.Item(8).Editable = False
            oGrid.Columns.Item(9).TitleObject.Caption = "Ref1"
            oGrid.Columns.Item(9).Editable = False
            oGrid.Columns.Item(10).TitleObject.Caption = "Remakrs"
            oGrid.Columns.Item(10).Editable = False
            oGrid.AutoResizeColumns()
            aForm.PaneLevel = 2
            '  aForm.Items.Item("7").Enabled = False
            aForm.Items.Item("21").Enabled = True
            aForm.Items.Item("29").Visible = True
            aForm.Items.Item("30").Visible = False
            aForm.Items.Item("34").Visible = False
            aForm.Items.Item("35").Visible = True
        ElseIf aChoice.ToUpper = "PRJ" Then
            oGrid = aForm.Items.Item("7").Specific
            oGrid.DataTable.ExecuteQuery(strQuery)
            oGrid.Columns.Item(0).TitleObject.Caption = "Project Code"
            oGrid.Columns.Item(0).TitleObject.Sortable = True
            oGrid.Columns.Item(0).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(0).Visible = False
            oGrid.Columns.Item(1).TitleObject.Caption = "Project Name"
            oGrid.Columns.Item(1).TitleObject.Sortable = True
            oGrid.Columns.Item(1).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oGrid.Columns.Item(2).TitleObject.Caption = "Phase"
            oGrid.Columns.Item(2).TitleObject.Sortable = True
            oGrid.Columns.Item(2).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(3).TitleObject.Caption = "Activity"
            oGrid.Columns.Item(3).TitleObject.Sortable = True
            oGrid.Columns.Item(3).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)


            oGrid.Columns.Item(4).TitleObject.Caption = "Employee ID"
            oGrid.Columns.Item(4).Visible = False
            oGrid.Columns.Item(5).TitleObject.Caption = "Employee Name"

            oGrid.Columns.Item(6).TitleObject.Caption = "Estimated Hours"
            oGrid.Columns.Item(6).TitleObject.Sortable = True
            oGrid.Columns.Item(6).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(7).TitleObject.Caption = "Actual Hours"
            oGrid.Columns.Item(7).TitleObject.Sortable = True
            oGrid.Columns.Item(7).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(7).Visible = True
            oGrid.Columns.Item(8).TitleObject.Caption = "Variance"
            oGrid.Columns.Item(8).TitleObject.Sortable = True
            oGrid.Columns.Item(8).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)


            oGrid.Columns.Item(9).TitleObject.Caption = "Estimated Cost"
            oGrid.Columns.Item(9).TitleObject.Sortable = True
            oGrid.Columns.Item(9).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(9).Visible = True
            oGrid.Columns.Item(10).TitleObject.Caption = "Actual Cost"
            oGrid.Columns.Item(10).TitleObject.Sortable = True
            oGrid.Columns.Item(10).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oGrid.Columns.Item(11).TitleObject.Caption = "Variance"
            oGrid.Columns.Item(11).TitleObject.Sortable = True
            oGrid.Columns.Item(11).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item("InvoicedAmt").TitleObject.Caption = "Invoiced Amount"
            oGrid.Columns.Item("InvoicedAmt").TitleObject.Sortable = True
            oGrid.Columns.Item("InvoicedAmt").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)


            oGrid.CollapseLevel = 3
            oGrid.AutoResizeColumns()
            '  PouplateProjectHours(oGrid)
            aForm.PaneLevel = 2
            aForm.Items.Item("7").Enabled = False
            aForm.Items.Item("29").Visible = False
            aForm.Items.Item("35").Visible = False
            aForm.Items.Item("30").Visible = False
            aForm.Items.Item("34").Visible = False

        ElseIf aChoice.ToUpper = "PRJS" Then
            oGrid = aForm.Items.Item("7").Specific
            oGrid.DataTable.ExecuteQuery(strQuery)
            oGrid.Columns.Item("Project").TitleObject.Caption = "Project Code"
            oGrid.Columns.Item("Project").TitleObject.Sortable = True
            oGrid.Columns.Item("Project").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item("Project").Visible = False
            oGrid.Columns.Item("ProjectName").TitleObject.Caption = "Project Name"
            oGrid.Columns.Item("ProjectName").TitleObject.Sortable = True
            oGrid.Columns.Item("ProjectName").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oGrid.Columns.Item("Business").TitleObject.Caption = "Phase"
            oGrid.Columns.Item("Business").TitleObject.Sortable = True
            oGrid.Columns.Item("Business").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item("Activity").TitleObject.Caption = "Activity"
            oGrid.Columns.Item("Activity").TitleObject.Sortable = True
            oGrid.Columns.Item("Activity").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item("Type").TitleObject.Caption = "Activity Type"


            oGrid.Columns.Item("EmpID").TitleObject.Caption = "Employee ID"
            oGrid.Columns.Item("EmpID").Visible = False
            oGrid.Columns.Item("Position").TitleObject.Caption = "Employee Name"
            oGrid.Columns.Item("Position").Visible = False



            oGrid.Columns.Item("FromDate").TitleObject.Caption = "Start Date"
            oGrid.Columns.Item("FromDate").Visible = False
            oGrid.Columns.Item("EndDate").TitleObject.Caption = "End Date"
            oGrid.Columns.Item("EndDate").Visible = False

            oGrid.Columns.Item("EstimatedHours").TitleObject.Caption = "Estimated Hours / Expense"
            oGrid.Columns.Item("EstimatedHours").TitleObject.Sortable = True
            oGrid.Columns.Item("EstimatedHours").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)


            oGrid.Columns.Item("ActualHours").TitleObject.Caption = "Actual Hours  / Expenses"
            oGrid.Columns.Item("ActualHours").TitleObject.Sortable = True
            oGrid.Columns.Item("ActualHours").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item("ActualHours").Visible = True
            oGrid.Columns.Item("VarianceHours").TitleObject.Caption = "Variance Hours  / Expenses"
            oGrid.Columns.Item("VarianceHours").TitleObject.Sortable = True
            oGrid.Columns.Item("VarianceHours").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item("Qty").TitleObject.Caption = "Quantity"
            oGrid.Columns.Item("Percentage").TitleObject.Caption = "Completion %"
            oGrid.Columns.Item("Status").TitleObject.Caption = "Status"
            oGrid.Columns.Item("CmpDate").TitleObject.Caption = "Completion Date"
            oGrid.Columns.Item("EstimatedCost").TitleObject.Caption = "Estimated Cost"
            oGrid.Columns.Item("EstimatedCost").TitleObject.Sortable = True
            oGrid.Columns.Item("EstimatedCost").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item("EstimatedCost").Visible = True
            oGrid.Columns.Item("MatrialOrdered").TitleObject.Caption = "Material Ordered"
            oGrid.Columns.Item("MatrialOrdered").TitleObject.Sortable = True
            oGrid.Columns.Item("MatrialOrdered").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)



            oGrid.Columns.Item("MaterialReceived").TitleObject.Caption = "Material Received"
            oGrid.Columns.Item("MaterialReceived").TitleObject.Sortable = True
            oGrid.Columns.Item("MaterialReceived").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oGrid.Columns.Item("OrderPending").TitleObject.Caption = "Order Pending"
            oGrid.Columns.Item("OrderPending").TitleObject.Sortable = True
            oGrid.Columns.Item("OrderPending").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oGrid.Columns.Item("MaterialPending").TitleObject.Caption = "Material Pending"
            oGrid.Columns.Item("MaterialPending").TitleObject.Sortable = True
            oGrid.Columns.Item("MaterialPending").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oGrid.Columns.Item("Cost Variance").TitleObject.Caption = "Cost Variance"
            oGrid.Columns.Item("Cost Variance").TitleObject.Sortable = True
            oGrid.Columns.Item("Cost Variance").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oGrid.Columns.Item("ReceivedAmt").TitleObject.Caption = "Invoiced Amount"
            oGrid.Columns.Item("ReceivedAmt").TitleObject.Sortable = True
            oGrid.Columns.Item("ReceivedAmt").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)


            oGrid.CollapseLevel = 3
            oGrid.AutoResizeColumns()
            '  PouplateProjectHours(oGrid)
            aForm.PaneLevel = 2
            aForm.Items.Item("7").Enabled = False
            aForm.Items.Item("29").Visible = False
            aForm.Items.Item("35").Visible = False
            aForm.Items.Item("30").Visible = False
            aForm.Items.Item("34").Visible = False

        ElseIf aChoice.ToUpper = "PRJM" Then
            oGrid = aForm.Items.Item("7").Specific
            oGrid.DataTable.ExecuteQuery(strQuery)
            oGrid.Columns.Item("Project").TitleObject.Caption = "Project Code"
            oGrid.Columns.Item("Project").TitleObject.Sortable = True
            oGrid.Columns.Item("Project").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item("Project").Visible = False
            oGrid.Columns.Item("PrjName").TitleObject.Caption = "Project Name"
            oGrid.Columns.Item("PrjName").TitleObject.Sortable = True
            oGrid.Columns.Item("PrjName").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oGrid.Columns.Item("Business").TitleObject.Caption = "Phase"
            oGrid.Columns.Item("Business").TitleObject.Sortable = True
            oGrid.Columns.Item("Business").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item("Activity").TitleObject.Caption = "Activity"
            oGrid.Columns.Item("Activity").TitleObject.Sortable = True
            oGrid.Columns.Item("Activity").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oGrid.Columns.Item("Qty").TitleObject.Caption = "Estimated Quantity"
            oGrid.Columns.Item("ItemCode").TitleObject.Caption = "ItemCode"
            oGrid.Columns.Item("UOM").TitleObject.Caption = "UoM"
            oGrid.Columns.Item("ItemName").TitleObject.Caption = "Item Description"
            oGrid.Columns.Item("Cost").TitleObject.Caption = "Estimated Cost"

            oGrid.Columns.Item("OrdQty").TitleObject.Sortable = True
            oGrid.Columns.Item("OrdQty").TitleObject.Caption = "Ordered Quantity"



            oGrid.Columns.Item("OrdCost").TitleObject.Caption = "Ordered Amount"

            oGrid.Columns.Item("RecQty").TitleObject.Caption = "Received Quantity"
            oGrid.Columns.Item("RecCost").TitleObject.Caption = "Invoiced Amount"

            oGrid.Columns.Item("RelQty").TitleObject.Caption = "Released Quantity"
            oGrid.Columns.Item("RelCost").TitleObject.Caption = "Released Amount"

            oGrid.CollapseLevel = 3
            oGrid.AutoResizeColumns()
            '  PouplateProjectHours(oGrid)
            aForm.PaneLevel = 2
            aForm.Items.Item("7").Enabled = False
            aForm.Items.Item("29").Visible = False
            aForm.Items.Item("35").Visible = False
            aForm.Items.Item("30").Visible = False
            aForm.Items.Item("34").Visible = False
        ElseIf aChoice.ToUpper = "TA" Then
            oGrid = aForm.Items.Item("7").Specific
            oGrid.DataTable.ExecuteQuery(strQuery)
            oGrid.Columns.Item(0).TitleObject.Caption = "Employee Code"
            oEditTextColumn = oGrid.Columns.Item(0)
            oEditTextColumn.LinkedObjectType = "171"
            oGrid.Columns.Item(0).TitleObject.Sortable = True
            oGrid.Columns.Item(0).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(1).TitleObject.Caption = "Employee Name"
            oGrid.Columns.Item(1).TitleObject.Sortable = True
            oGrid.Columns.Item(1).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(2).TitleObject.Caption = "Document Date"
            oGrid.Columns.Item(2).TitleObject.Sortable = True
            oGrid.Columns.Item(2).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(3).TitleObject.Caption = "Project Code"
            oGrid.Columns.Item(3).TitleObject.Sortable = True
            oGrid.Columns.Item(3).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(4).TitleObject.Caption = "Project Name"
            oGrid.Columns.Item(4).TitleObject.Sortable = True
            oGrid.Columns.Item(4).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(4).Visible = False
            oGrid.Columns.Item(5).TitleObject.Caption = "Phase"
            oGrid.Columns.Item(5).TitleObject.Sortable = True
            oGrid.Columns.Item(5).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(6).TitleObject.Caption = "Activity"
            oGrid.Columns.Item(6).Editable = False
            oGrid.Columns.Item(6).TitleObject.Sortable = True
            oGrid.Columns.Item(6).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(7).TitleObject.Sortable = True
            oGrid.Columns.Item(7).TitleObject.Caption = "Leave Type / Time sheet Type"
            oGrid.Columns.Item(7).Editable = False
            oGrid.Columns.Item(7).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(8).TitleObject.Caption = "From Date"
            oGrid.Columns.Item(8).Editable = False
            oGrid.Columns.Item(8).TitleObject.Sortable = True
            oGrid.Columns.Item(8).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(9).TitleObject.Caption = "End Date"
            oGrid.Columns.Item(9).Editable = False
            oGrid.Columns.Item(9).TitleObject.Sortable = True
            oGrid.Columns.Item(9).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oGrid.Columns.Item(10).TitleObject.Caption = "No of Days / Hours"
            oEditTextColumn = oGrid.Columns.Item(10)
            ' oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oGrid.Columns.Item(11).TitleObject.Caption = "Reason / Remarks"
            oGrid.Columns.Item(12).TitleObject.Caption = "Status"
            oGrid.Columns.Item(12).TitleObject.Sortable = True
            oGrid.Columns.Item(12).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oGrid.AutoResizeColumns()
            aForm.PaneLevel = 2
            aForm.Items.Item("7").Enabled = False
            aForm.Items.Item("29").Visible = False
            aForm.Items.Item("30").Visible = False
            aForm.Items.Item("34").Visible = False
            aForm.Items.Item("35").Visible = False
        Else
            oGrid = aForm.Items.Item("26").Specific
            oGrid.DataTable = CreateColumns(oGrid, aForm)
            DatabindSummary(oForm)
            oGrid.AutoResizeColumns()
            aForm.Items.Item("26").Enabled = False

            ' aForm.PaneLevel = 2

            oGrid = aForm.Items.Item("27").Specific
            oGrid.DataTable.ExecuteQuery(strQuery)
            oGrid.Columns.Item(0).TitleObject.Caption = "Employee Code"
            oEditTextColumn = oGrid.Columns.Item(0)
            oEditTextColumn.LinkedObjectType = "171"
            oGrid.Columns.Item(0).TitleObject.Sortable = True
            oGrid.Columns.Item(0).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(1).TitleObject.Caption = "Employee Name"
            oGrid.Columns.Item(1).TitleObject.Sortable = True
            oGrid.Columns.Item(1).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(2).TitleObject.Caption = "Document Date"
            oGrid.Columns.Item(2).TitleObject.Sortable = True
            oGrid.Columns.Item(2).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(3).TitleObject.Caption = "Project Code"
            oEditTextColumn = oGrid.Columns.Item(3)
            oEditTextColumn.LinkedObjectType = "63"
            oGrid.Columns.Item(3).TitleObject.Sortable = True
            oGrid.Columns.Item(3).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(4).TitleObject.Caption = "Project Name"


            oGrid.Columns.Item(3).TitleObject.Sortable = True
            oGrid.Columns.Item(3).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(4).TitleObject.Caption = "Project Name"



            oGrid.Columns.Item(5).TitleObject.Caption = "Phase"
            oGrid.Columns.Item(5).TitleObject.Sortable = True
            oGrid.Columns.Item(5).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(6).TitleObject.Caption = "Activity"
            oGrid.Columns.Item(6).TitleObject.Sortable = True
            oGrid.Columns.Item(6).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oGrid.Columns.Item("U_Z_TYPE").TitleObject.Caption = "Activity Type"
            oGrid.Columns.Item("U_Z_BDGQTY").TitleObject.Caption = "Budgeted  Hours"
            oGrid.Columns.Item("U_Z_QUANTITY").TitleObject.Caption = "Quantity"
            oGrid.Columns.Item("U_Z_QUANTITY").Visible = False
            oGrid.Columns.Item("U_Z_MEASURE").TitleObject.Caption = "Measure"
            ' oGrid.Columns.Item("U_Z_TYPE").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oGrid.Columns.Item("U_Z_MEASURE").Visible = False
            oGrid.Columns.Item(11).TitleObject.Caption = "Date"
            oGrid.Columns.Item(11).TitleObject.Sortable = True
            oGrid.Columns.Item(11).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(11).Visible = False
            oEditTextColumn = oGrid.Columns.Item("U_Z_HOURS")
            oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oGrid.Columns.Item("U_Z_HOURS").TitleObject.Caption = "Hours"
            oGrid.Columns.Item("U_Z_HOURS").TitleObject.Sortable = True
            oGrid.Columns.Item("U_Z_HOURS").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item("U_Z_REFCODE").Visible = False
            oGrid.Columns.Item("U_Z_REMARKS").TitleObject.Caption = "Remarks"
            oGrid.Columns.Item("STATUS").TitleObject.Caption = "Approval Status"
            oGrid.Columns.Item("STATUS").TitleObject.Sortable = True
            oGrid.Columns.Item("STATUS").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oGrid.AutoResizeColumns()
            aForm.Items.Item("27").Enabled = False
            aForm.PaneLevel = 3

        End If
        'oGrid.Columns.Item(4).TitleObject.Caption = "Module Name"
        'oGrid.Columns.Item(5).TitleObject.Caption = "Activity Name"
        'oGrid.Columns.Item(6).TitleObject.Caption = "Date"
        'oGrid.Columns.Item(7).TitleObject.Caption = "No of Hours"
        'oEditTextColumn = oGrid.Columns.Item(7)
        '' oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
        'oGrid.CollapseLevel = 2


    End Sub
#End Region

#Region "AddRow /Delete Row"
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)

        oGrid = aForm.Items.Item("6").Specific
        If oGrid.DataTable.GetValue("U_ITMSGRPCOD", oGrid.DataTable.Rows.Count - 1) <> "" Then
            oGrid.DataTable.Rows.Add()
        End If
    End Sub

    Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form)
        Dim orec As SAPbobsCOM.Recordset
        Dim strTablename As String
        strTablename = ""
        oGrid = aForm.Items.Item("6").Specific
        strTablename = "[@DABT_OITB]"

        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.Rows.IsSelected(intRow) Then
                orec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                orec.DoQuery("Update" & strTablename & " set Name = 'D' +Name where code='" & oGrid.DataTable.GetValue("Code", intRow) & "'")
                oGrid.DataTable.Rows.Remove(intRow)
                Exit For
            End If
        Next
    End Sub
#End Region

#Region "AddtoUDT"
    Private Function AddToUDT(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strItem, strCode, strWhs, strBin, strwhsdesc, strbindesc, strTo, strHeaderRef, strConditionType, strfromdate, strCardCode As String
        Dim oTempRec, otemp As SAPbobsCOM.Recordset
        Dim ousertable As SAPbobsCOM.UserTable
        Dim ocheckbox As SAPbouiCOM.CheckBoxColumn
        Dim oedittext As SAPbouiCOM.EditTextColumn
        Dim dblPercentage As Double
        Dim dtFrom, dtTo As Date
        Dim oBPGrid As SAPbouiCOM.Grid

        oBPGrid = aform.Items.Item("6").Specific
        strCardCode = oApplication.Utilities.getEdittextvalue(aform, "4")
        ousertable = oApplication.Company.UserTables.Item("DABT_OITB")
        For intLoop As Integer = 0 To oBPGrid.DataTable.Rows.Count - 1
            strCode = oBPGrid.DataTable.GetValue(0, intLoop)
            oCombobox = oBPGrid.Columns.Item("U_ITMSGRPCOD")
            strfromdate = oCombobox.GetSelectedValue(intLoop).Value
            If strfromdate <> "" Then
                otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                otemp.DoQuery("Select * from [@DABT_OITB] where U_CardCode='" & strCardCode & "' and U_ITMSGRPCOD='" & strfromdate & "'")
                If otemp.RecordCount > 0 Then
                    strCode = otemp.Fields.Item("Code").Value
                Else
                    strCode = oBPGrid.DataTable.GetValue(0, intLoop)
                End If
                If strCode <> "" Then
                    ousertable.GetByKey(strCode)
                    ousertable.Name = strCode
                    ousertable.UserFields.Fields.Item("U_CardCode").Value = strCardCode

                    oCombobox = oBPGrid.Columns.Item("U_ITMSGRPCOD")
                    ousertable.UserFields.Fields.Item("U_ITMSGRPCOD").Value = oCombobox.GetSelectedValue(intLoop).Value
                    If ousertable.Update <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                Else
                    strCode = oApplication.Utilities.getMaxCode("@DABT_OITB", "Code")
                    ousertable.Code = strCode
                    ousertable.Name = strCode
                    ousertable.UserFields.Fields.Item("U_CardCode").Value = strCardCode
                    oCombobox = oBPGrid.Columns.Item("U_ITMSGRPCOD")
                    ousertable.UserFields.Fields.Item("U_ITMSGRPCOD").Value = oCombobox.GetSelectedValue(intLoop).Value
                    If ousertable.Add <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                End If
            End If
        Next
        oRecset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecset.DoQuery("Delete from [@DABT_OITB] where name like 'D%'")
        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_None)
        Return True
    End Function

#End Region

#Region "Post Journal Entry"
    Private Function PostJournalEntry_Old(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim oJE As SAPbobsCOM.JournalEntries
            Dim strcreditAccount As String
            Dim dblCreditAmount As Double
            Dim intCount As Double = 0
            strcreditAccount = oApplication.Utilities.getEdittextvalue(aForm, "102")
            strcreditAccount = oApplication.Utilities.GetAccount(strcreditAccount, "AcctCode", "FormatCode")
            oGrid = aForm.Items.Item("7").Specific
            oJE = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
            dblCreditAmount = 0
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If intRow > 0 Then
                    oJE.Lines.Add()
                    oJE.Lines.SetCurrentLine(intRow)
                End If
                oJE.Lines.AccountCode = oApplication.Utilities.GetAccount(oGrid.DataTable.GetValue(3, intRow), "AcctCode", "FormatCode")
                oJE.Lines.ProjectCode = oGrid.DataTable.GetValue(1, intRow)
                oJE.Lines.Debit = oGrid.DataTable.GetValue(5, intRow)
                oJE.Lines.Credit = 0
                intCount = intCount + 1
                dblCreditAmount = dblCreditAmount + oGrid.DataTable.GetValue(5, intRow)
            Next
            If dblCreditAmount > 0 Then
                oJE.Lines.Add()
                oJE.Lines.SetCurrentLine(intCount)
                oJE.Lines.AccountCode = strcreditAccount
                oJE.Lines.Credit = dblCreditAmount
                If oJE.Add <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Else
                    Dim stCode As String
                    oApplication.Company.GetNewObjectCode(stCode)
                    Dim ors As SAPbobsCOM.Recordset
                    ors = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    ors.DoQuery("Update [@Z_EXP1] set U_Z_Flag='N' , U_Z_DocNum='" & stCode & "' where U_Z_Flag='Y' ")
                    oApplication.Utilities.Message("Journal Entry Created successfully : " & stCode, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                    Return True
                End If
            End If

        Catch ex As Exception

        End Try
    End Function

    Private Function PostJournalEntry(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim oJE As SAPbobsCOM.JournalEntries
            Dim strcreditAccount, strCondition, strSQL, strsql1, strEMpid, strEmpName, strDate, strcode, stremployeeName, strPostingChoice As String
            Dim dblCreditAmount As Double
            Dim dtDate, dtDocDate As Date
            Dim oTempRs, oJournalRS As SAPbobsCOM.Recordset
            Dim intCount As Double = 0
            strcreditAccount = oApplication.Utilities.getEdittextvalue(aForm, "102")
            strDate = oApplication.Utilities.getEdittextvalue(aForm, "104")
            dtDocDate = oApplication.Utilities.GetDateTimeValue(strDate)
            strcreditAccount = oApplication.Utilities.GetAccount(strcreditAccount, "AcctCode", "FormatCode")
            oCombo = aForm.Items.Item("36").Specific
            If oCombo.Selected.Value = "B" Then
                strPostingChoice = "B"
            Else
                strPostingChoice = "E"
            End If
            oGrid = aForm.Items.Item("7").Specific
            dblCreditAmount = 0
            strCondition = JournalQuery(aForm)
            If strCondition <> "" Then
                'strSQL = "SELECT  T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], convert(varchar(10),T0.[U_Z_DOCDATE],105),T1.[U_Z_EXPNAME], T1.[U_Z_PRJCODE],T1.[U_Z_PRJNAME],T3.[U_Z_ActCode],T4.[AcctName],T1.[U_Z_LocAmount],T1.[U_Z_Ref1],T1.[U_Z_Remarks]  FROM [dbo].[@Z_OEXP]  T0 inner join  [dbo].[@Z_EXP1]  T1  on T0.Code=T1.U_Z_REFCODE inner join [@Z_Expances] T3 on T3.U_Z_ExpName=T1.[U_Z_ExpName] inner join OACT T4 on T4.AcctCode=T3.[U_Z_ActCode]"
                'strSQL = strSQL & " where T1.[U_Z_APPROVED]='A' and isnull(T1.[U_Z_DocNum],'')='' and  T1.[U_Z_Flag]='Y' and " & strCondition & " Order by T0.U_Z_DocDate, T0.[U_Z_EMPCODE],T1.[U_Z_EXPNAME], T1.[U_Z_PRJCODE] "
                strSQL = "SELECT  T0.[U_Z_EMPCODE], T0.[U_Z_DOCDATE],count(*)  FROM [dbo].[@Z_OEXP]  T0 inner join  [dbo].[@Z_EXP1]  T1  on T0.Code=T1.U_Z_REFCODE inner join [@Z_Expances] T3 on T3.U_Z_ExpName=T1.[U_Z_ExpName] inner join OACT T4 on T4.FormatCode=T3.[U_Z_ActCode]"
                strSQL = strSQL & " where T1.[U_Z_APPROVED]='A' and isnull(T1.[U_Z_DocNum],'')='' and  T1.[U_Z_Flag]='Y' and " & strCondition & " group by T0.[U_Z_EMPCODE], T0.[U_Z_DOCDATE]"
                oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oTempRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTempRs.DoQuery(strSQL)
                If oApplication.Company.InTransaction Then
                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                End If
                oApplication.Company.StartTransaction()
                Dim strProjectcode As String

                For intRow As Integer = 0 To oTempRs.RecordCount - 1
                    oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    strEMpid = oTempRs.Fields.Item(0).Value
                    dtDate = oTempRs.Fields.Item(1).Value
                    strsql1 = "SELECT  T1.[Code],T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], convert(varchar(10),T0.[U_Z_DOCDATE],105),T1.[U_Z_EXPNAME], T1.[U_Z_PRJCODE],T1.[U_Z_PRJNAME],T3.[U_Z_ActCode],T4.[AcctName],T1.[U_Z_LocAmount],T1.[U_Z_Ref1],T1.[U_Z_Remarks],T1.U_Z_MODNAME 'Mod',T1.U_Z_ACTNAME 'Act' FROM [dbo].[@Z_OEXP]  T0 inner join  [dbo].[@Z_EXP1]  T1  on T0.Code=T1.U_Z_REFCODE inner join [@Z_Expances] T3 on T3.U_Z_ExpName=T1.[U_Z_ExpName] inner join OACT T4 on T4.FormatCode=T3.[U_Z_ActCode]"
                    strsql1 = strsql1 & " where T0.[U_Z_EMPCODE]='" & strEMpid & "' and U_Z_DocDate='" & dtDate.ToString("yyyy-MM-dd") & "' and  T1.[U_Z_APPROVED]='A' and isnull(T1.[U_Z_DocNum],'')='' and  T1.[U_Z_Flag]='Y' and " & strCondition & " Order by T0.U_Z_DocDate, T0.[U_Z_EMPCODE],T1.[U_Z_EXPNAME], T1.[U_Z_PRJCODE] "
                    oJournalRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oJournalRS.DoQuery(strsql1)
                    dblCreditAmount = 0
                    intCount = 0
                    stremployeeName = ""
                    oJE = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                    oJE.DueDate = dtDocDate
                    ' oJE.TaxDate = dtDocDate
                    oJE.TaxDate = dtDate
                    oJE.ReferenceDate = dtDocDate
                    strcode = ""
                    strProjectcode = ""
                    For intloop As Integer = 0 To oJournalRS.RecordCount - 1
                        If intloop > 0 Then
                            oJE.Lines.Add()
                            oJE.Lines.SetCurrentLine(intloop)
                        End If
                        stremployeeName = oJournalRS.Fields.Item("U_Z_EMPNAME").Value
                        oJE.Lines.AccountCode = oApplication.Utilities.GetAccount(oJournalRS.Fields.Item("U_Z_ActCode").Value, "AcctCode", "FormatCode")
                        oJE.Lines.ProjectCode = oJournalRS.Fields.Item("U_Z_PRJCODE").Value
                        oJE.Lines.Reference1 = oJournalRS.Fields.Item("U_Z_Ref1").Value
                        oJE.Lines.Reference2 = oJournalRS.Fields.Item("U_Z_EMPNAME").Value
                        oJE.Lines.LineMemo = oJournalRS.Fields.Item("U_Z_Remarks").Value
                        oJE.Lines.Debit = oJournalRS.Fields.Item("U_Z_LocAmount").Value
                        strProjectcode = oJournalRS.Fields.Item("U_Z_PRJCODE").Value
                        oJE.Lines.UserFields.Fields.Item("U_Z_ACTNAME").Value = oJournalRS.Fields.Item("Act").Value
                        oJE.Lines.UserFields.Fields.Item("U_Z_MDNAME").Value = oJournalRS.Fields.Item("Mod").Value
                        oJE.Lines.Credit = 0
                        intCount = intCount + 1
                        dblCreditAmount = dblCreditAmount + oJournalRS.Fields.Item("U_Z_LocAmount").Value

                        If strcode = "" Then
                            strcode = "'" & oJournalRS.Fields.Item("Code").Value & "'"
                        Else
                            strcode = strcode & ",'" & oJournalRS.Fields.Item("Code").Value & "'"
                        End If
                        oJournalRS.MoveNext()
                    Next
                    If dblCreditAmount > 0 Then
                        oJE.Lines.Add()
                        oJE.Lines.SetCurrentLine(intCount)
                        If strPostingChoice = "E" Then
                            oJE.Lines.AccountCode = strcreditAccount
                        Else
                            Dim oBPRS As SAPbobsCOM.Recordset
                            oBPRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oBPRS.DoQuery("Select isnull(U_CardCode,'') from OHEM where empID=" & strEMpid)
                            oJE.Lines.ShortName = oBPRS.Fields.Item(0).Value
                        End If
                        If strProjectcode <> "" Then
                            oJE.Lines.ProjectCode = strProjectcode
                        End If
                        oJE.Lines.Credit = dblCreditAmount
                        oJE.Lines.Reference2 = stremployeeName
                        If oJE.Add <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            If oApplication.Company.InTransaction Then
                                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                            Return False
                        Else
                            Dim stCode As String
                            oApplication.Company.GetNewObjectCode(stCode)
                            Dim ors As SAPbobsCOM.Recordset
                            ors = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'ors.DoQuery("Update [@Z_EXP1] set U_Z_Flag='N' , U_Z_DocNum='" & stCode & "' where U_Z_Flag='Y' and  U_Z_Date='" & dtDate.ToString("yyyy-MM-dd") & "' ")
                            ors.DoQuery("Update [@Z_EXP1] set U_Z_Flag='N' , U_Z_DocNum='" & stCode & "' where U_Z_Flag='Y' and  U_Z_Date='" & dtDate.ToString("yyyy-MM-dd") & "' and code in (" & strcode & ")")
                            oApplication.Utilities.Message("Journal Entry Created successfully : " & stCode, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            '  Return True
                        End If
                    End If
                    oTempRs.MoveNext()
                Next

            End If
            If oApplication.Company.InTransaction Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If

            oApplication.Utilities.Message("Operation Completed Successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            'For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            '    If intRow > 0 Then
            '        oJE.Lines.Add()
            '        oJE.Lines.SetCurrentLine(intRow)
            '    End If
            '    oJE.Lines.AccountCode = oApplication.Utilities.GetAccount(oGrid.DataTable.GetValue(3, intRow), "AcctCode", "FormatCode")
            '    oJE.Lines.ProjectCode = oGrid.DataTable.GetValue(1, intRow)
            '    oJE.Lines.Debit = oGrid.DataTable.GetValue(5, intRow)
            '    oJE.Lines.Credit = 0
            '    intCount = intCount + 1
            '    dblCreditAmount = dblCreditAmount + oGrid.DataTable.GetValue(5, intRow)
            'Next
            'If dblCreditAmount > 0 Then
            '    oJE.Lines.Add()
            '    oJE.Lines.SetCurrentLine(intCount)
            '    oJE.Lines.AccountCode = strcreditAccount
            '    oJE.Lines.Credit = dblCreditAmount
            '    If oJE.Add <> 0 Then
            '        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        Return False
            '    Else
            '        Dim stCode As String
            '        oApplication.Company.GetNewObjectCode(stCode)
            '        Dim ors As SAPbobsCOM.Recordset
            '        ors = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '        ors.DoQuery("Update [@Z_EXP1] set U_Z_Flag='N' , U_Z_DocNum='" & stCode & "' where U_Z_Flag='Y' ")
            '        oApplication.Utilities.Message("Journal Entry Created successfully : " & stCode, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            '        Return True
            '    End If
            'End If
            Return True
        Catch ex As Exception
            Return False

        End Try
    End Function
#End Region



#End Region

#End Region

#Region "Events"



#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_PrjReports
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If

                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD, mnu_FIND
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        'EnableControls(oForm)
                    End If

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
            If pVal.FormTypeEx = frm_PrjReports Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "2" Then
                                    Dim otempRs As SAPbobsCOM.Recordset
                                    otempRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "7" Then
                                    oCombo = oForm.Items.Item("7").Specific
                                    If oCombo.Selected.Value = "" Then
                                        ' oApplication.Utilities.Message("Report type is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        ' Exit Sub
                                    Else
                                        oForm.Freeze(True)
                                        oApplication.Utilities.setEdittextvalue(oForm, "9", "")
                                        oApplication.Utilities.setEdittextvalue(oForm, "11", "")
                                        oApplication.Utilities.setEdittextvalue(oForm, "13", "")
                                        oApplication.Utilities.setEdittextvalue(oForm, "15", "")
                                        oApplication.Utilities.setEdittextvalue(oForm, "17", "")
                                        If oCombo.Selected.Value = "TA" Or oCombo.Selected.Value = "RS" Then
                                            oForm.Items.Item("117").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oForm.Items.Item("8").Visible = True
                                            oForm.Items.Item("9").Visible = True
                                            oForm.Items.Item("10").Visible = True
                                            oForm.Items.Item("11").Visible = True
                                            oForm.Items.Item("12").Visible = True
                                            oForm.Items.Item("13").Visible = True
                                            oForm.Items.Item("14").Visible = True
                                            oForm.Items.Item("15").Visible = True
                                            oForm.Items.Item("13").Visible = True
                                            oForm.Items.Item("15").Visible = True
                                            oForm.Items.Item("16").Visible = False
                                            oForm.Items.Item("17").Visible = False
                                            oForm.Items.Item("9").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        Else
                                            oForm.Items.Item("117").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oForm.Items.Item("16").Visible = True
                                            oForm.Items.Item("17").Visible = True
                                            oForm.Items.Item("8").Visible = False
                                            oForm.Items.Item("9").Visible = False
                                            oForm.Items.Item("10").Visible = False
                                            oForm.Items.Item("11").Visible = False
                                            oForm.Items.Item("12").Visible = False
                                            oForm.Items.Item("13").Visible = False
                                            oForm.Items.Item("14").Visible = False
                                            oForm.Items.Item("15").Visible = False
                                            oForm.Items.Item("13").Visible = False
                                            oForm.Items.Item("15").Visible = False
                                            oForm.Items.Item("17").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        End If
                                        oForm.PaneLevel = 3
                                        oForm.Freeze(False)
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oForm.Freeze(True)
                                Select Case pVal.ItemUID
                                    Case "21"
                                        ' databind(oForm)
                                    Case "3"
                                        oForm.PaneLevel = oForm.PaneLevel - 1
                                        oForm.Items.Item("7").Enabled = True
                                    Case "4"
                                        If oForm.PaneLevel = 3 Then
                                            databind(oForm)
                                            oForm.Items.Item("117").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oForm.Items.Item("7").Enabled = False
                                            oForm.PaneLevel = oForm.PaneLevel + 1
                                        ElseIf oForm.PaneLevel = 1 Then
                                            oForm.PaneLevel = oForm.PaneLevel + 1
                                            oForm.Items.Item("117").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oForm.Items.Item("7").Enabled = True
                                        ElseIf oForm.PaneLevel = 2 Then
                                            oCombo = oForm.Items.Item("7").Specific
                                            If oCombo.Selected.Value = "" Then
                                                oApplication.Utilities.Message("Report type is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                oForm.Freeze(False)
                                                Exit Sub
                                            Else
                                                oForm.PaneLevel = oForm.PaneLevel + 1
                                                oForm.Items.Item("117").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                oForm.Items.Item("7").Enabled = True
                                                ''  oForm.Freeze(True)
                                                'If oCombo.Selected.Value = "TA" Or oCombo.Selected.Value = "RS" Then
                                                '    oForm.Items.Item("117").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                                                '    oForm.Items.Item("8").Visible = True
                                                '    oForm.Items.Item("9").Visible = True
                                                '    oForm.Items.Item("10").Visible = True
                                                '    oForm.Items.Item("11").Visible = True
                                                '    oForm.Items.Item("12").Visible = True
                                                '    oForm.Items.Item("13").Visible = True
                                                '    oForm.Items.Item("14").Visible = True
                                                '    oForm.Items.Item("15").Visible = True
                                                '    oForm.Items.Item("13").Visible = True
                                                '    oForm.Items.Item("15").Visible = True
                                                '    oForm.Items.Item("16").Visible = False
                                                '    oForm.Items.Item("17").Visible = False
                                                '    oForm.Items.Item("9").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                'Else
                                                '    oForm.Items.Item("117").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                '    oForm.Items.Item("16").Visible = True
                                                '    oForm.Items.Item("17").Visible = True
                                                '    oForm.Items.Item("8").Visible = False
                                                '    oForm.Items.Item("9").Visible = False
                                                '    oForm.Items.Item("10").Visible = False
                                                '    oForm.Items.Item("11").Visible = False
                                                '    oForm.Items.Item("12").Visible = False
                                                '    oForm.Items.Item("13").Visible = False
                                                '    oForm.Items.Item("14").Visible = False
                                                '    oForm.Items.Item("15").Visible = False
                                                '    oForm.Items.Item("13").Visible = False
                                                '    oForm.Items.Item("15").Visible = False
                                                '    oForm.Items.Item("17").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                'End If
                                                ' oForm.Freeze(False)
                                            End If
                                        End If

                                    Case "30"
                                        oGrid = oForm.Items.Item("7").Specific
                                        For introw As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                            oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            oCheckBoxColumn = oGrid.Columns.Item(13)
                                            oCheckBoxColumn.Check(introw, True)
                                        Next
                                        oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
                                    Case "34"
                                        oGrid = oForm.Items.Item("7").Specific
                                        For introw As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                            oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            oCheckBoxColumn = oGrid.Columns.Item(13)
                                            oCheckBoxColumn.Check(introw, False)
                                        Next
                                        oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
                                    Case "35"
                                        oitems = oForm.Items.Item("29")
                                        obutton = oitems.Specific
                                        obutton.Caption = "Copy to Journal"
                                        Dim otempRs As SAPbobsCOM.Recordset
                                        otempRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        otempRs.DoQuery("Update  [@Z_EXP1] set U_Z_Flag='N' where U_Z_Flag='Y'")
                                        databind(oForm)
                                        oForm.Items.Item("21").Enabled = True
                                    Case "29"
                                        oGrid = oForm.Items.Item("7").Specific
                                        Dim otempRs As SAPbobsCOM.Recordset
                                        Dim st, strAmount As String
                                        otempRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oitems = oForm.Items.Item("29")
                                        obutton = oitems.Specific
                                        If obutton.Caption <> "Post Journal" Then
                                            For introw As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                                oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                                oCheckBoxColumn = oGrid.Columns.Item(13)
                                                If oCheckBoxColumn.IsChecked(introw) Then
                                                    strAmount = oGrid.DataTable.GetValue(10, introw)
                                                    Dim dblAmount As Double
                                                    dblAmount = (oApplication.Utilities.getDocumentQuantity(strAmount))
                                                    st = "Update  [@Z_EXP1] set U_Z_Flag='Y',U_Z_LocAmount='" & dblAmount.ToString.Replace(",", ".") & "' where Code='" & oGrid.DataTable.GetValue(14, introw) & "'"
                                                    otempRs.DoQuery(st)
                                                End If
                                            Next
                                            oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
                                            otempRs.DoQuery("Select * from [@Z_EXP1] where U_Z_Flag='Y'")
                                            If otempRs.RecordCount > 0 Then
                                                oitems = oForm.Items.Item("29")
                                                obutton = oitems.Specific
                                                obutton.Caption = "Post Journal"
                                                databind(oForm)
                                                oForm.Items.Item("34").Visible = False
                                                oForm.Items.Item("30").Visible = False
                                                oForm.Items.Item("21").Enabled = False
                                            Else
                                                oApplication.Utilities.Message("No rows selected", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            End If

                                        Else
                                            oitems = oForm.Items.Item("29")
                                            obutton = oitems.Specific
                                            'If oApplication.Company.InTransaction() Then
                                            '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                            'End If
                                            'oApplication.Company.StartTransaction()
                                            If PostJournalEntry(oForm) = False Then
                                                obutton.Caption = "Copy to Journal"
                                                oForm.Items.Item("35").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                'If oApplication.Company.InTransaction() Then
                                                '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                'End If
                                            Else
                                                oApplication.Utilities.Message("Operation completed succssfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                'If oApplication.Company.InTransaction() Then
                                                '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                'End If
                                                oForm.Items.Item("35").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            End If
                                        End If
                                End Select
                                oForm.Freeze(False)
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "31" Then
                                    oForm.PaneLevel = 3

                                End If
                                If pVal.ItemUID = "32" Then
                                    oForm.PaneLevel = 4

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

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
                                        If pVal.ItemUID = "11" Or pVal.ItemUID = "9" Then
                                            val = oDataTable.GetValue("empID", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        ElseIf pVal.ItemUID = "17" Then
                                            val = oDataTable.GetValue("U_Z_PRJCODE", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                End Try

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
