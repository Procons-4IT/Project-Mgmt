Public Class clsReports
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
    private oitems as SAPbouiCOM.Item 
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
        oForm = oApplication.Utilities.LoadForm(xml_Report, frm_Report)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataSources.UserDataSources.Add("empFrom", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("empTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("dtFrom", SAPbouiCOM.BoDataType.dt_DATE)
        oForm.DataSources.UserDataSources.Add("dtTo", SAPbouiCOM.BoDataType.dt_DATE)
        oForm.DataSources.UserDataSources.Add("prjFrom", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("prjto", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        AddChooseFromList(oForm)
        oApplication.Utilities.setUserDatabind(oForm, "9", "empfrom")
        oEditText = oForm.Items.Item("9").Specific
        oEditText.ChooseFromListUID = "CFL4"
        oEditText.ChooseFromListAlias = "empID"
        oApplication.Utilities.setUserDatabind(oForm, "12", "empTo")
        oEditText = oForm.Items.Item("12").Specific
        oEditText.ChooseFromListUID = "CFL5"
        oEditText.ChooseFromListAlias = "empID"
        oApplication.Utilities.setUserDatabind(oForm, "18", "dtfrom")
        oApplication.Utilities.setUserDatabind(oForm, "20", "dtTo")
        'databind(oForm)
        oCombo = oForm.Items.Item("4").Specific
        oCombo.ValidValues.Add("", "")
        oCombo.ValidValues.Add("Exp", "Expenses")
        oCombo.ValidValues.Add("TIME", "Time Sheet")
        oCombo.ValidValues.Add("Prj", "Project Wise Time Sheet")
        oCombo.ValidValues.Add("LEV", "Leave Request")
        oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oForm.Items.Item("4").DisplayDesc = True
        oForm.Items.Item("6").DisplayDesc = True
        oForm.Items.Item("15").DisplayDesc = True
        oForm.Items.Item("7").Visible = False
        ' oForm.Items.Item("17").Visible = False
        fillCombo(oForm)
        oForm.PaneLevel = 0
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
            oCFL = oCFLs.Item("CFL_7")

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


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
        AddChooseFromList(oForm)
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
            Dim stEmpTA As String = ""
            If strFromEMP <> "" And strToEMP <> "" Then
                strEMPCondition = " U_Z_EMPCODE between '" & strFromEMP & "' and '" & strToEMP & "'"
                stEmpTA = " T0.U_Z_EMPID between '" & strFromEMP & "' and '" & strToEMP & "'"
            ElseIf strFromEMP <> "" And strToEMP = "" Then
                strEMPCondition = " U_Z_EMPCODE >= '" & strFromEMP & "'"
                stEmpTA = " T0.U_Z_EMPID >= '" & strFromEMP & "'"
            ElseIf strFromEMP = "" And strToEMP <> "" Then
                strEMPCondition = " U_Z_EMPCODE <= '" & strToEMP & "'"
                stEmpTA = " T0.U_Z_EMPID <= '" & strToEMP & "'"
            Else
                strEMPCondition = " 1=1"
                stEmpTA = " 1=1"
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
                otemp.DoQuery("Select isnull(sum(U_Z_Days),0) from [@Z_PRJ1] where  U_Z_ModName='" & strModule.Replace("'", "''") & "' and  DocEntry= (Select docEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & oRS.Fields.Item(0).Value & "') group by U_Z_ModName")
                Dim s1 As String
                s1 = "Select Sum(Isnull(U_Z_HOURS,0)/8) from [@Z_PRJ1] T0 inner Join [@Z_HPRJ] T1 on T1.DocEntry=T0.DocEntry where (" & stEmpTA & ") and  U_Z_ModName='" & strModule.Replace("'", "''") & "' and T1.U_Z_PRJCODE='" & oRS.Fields.Item(0).Value & "' group by U_Z_MODNAME"
                otemp.DoQuery(s1)
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
                    Exit Sub
                End If
            Else
                If oApplication.Utilities.getEdittextvalue(aForm, "104") = "" Then
                    oApplication.Utilities.Message("Posting date missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Exit Sub
                End If
                oCombo = aForm.Items.Item("36").Specific
                If oCombo.Selected.Value = "" Then
                    oApplication.Utilities.Message("Posting Type missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Exit Sub
                End If
                strreporttype = oApplication.Utilities.getEdittextvalue(aForm, "102")
                If strreporttype = "" And oCombo.Selected.Value = "E" Then
                    oApplication.Utilities.Message("Account is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Exit Sub
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
            Dim strExmEmpCondition As String = ""
            strEMPCondition = ""
            strDateCondition = ""
            strProjectCondition = ""
            Dim stPrjEmpCondition As String
            If strFromEMP <> "" And strToEMP <> "" Then
                strEMPCondition = " Convert(Decimal,isnull(U_Z_EMPCODE,'0')) between " & CDbl(strFromEMP) & " and " & CDbl(strToEMP)
                strExmEmpCondition = " Convert(Decimal,isnull(T1.U_Z_EMPID,'0')) between " & CDbl(strFromEMP) & " and " & CDbl(strToEMP)


            ElseIf strFromEMP <> "" And strToEMP = "" Then
                strEMPCondition = " Convert(Decimal,isnull(U_Z_EMPCODE,'0')) >= " & CDbl(strFromEMP)
                strExmEmpCondition = " Convert(Decimal,isnull(T1.U_Z_EMPID,'0')) >= " & CDbl(strFromEMP)
            ElseIf strFromEMP = "" And strToEMP <> "" Then
                strEMPCondition = " Convert(Decimal,isnull(U_Z_EMPCODE,'0')) <= " & CDbl(strToEMP)
                strExmEmpCondition = " Convert(Decimal,isnull(T1.U_Z_EMPID,'0')) <= " & CDbl(strToEMP)
            Else
                strEMPCondition = " 1=1"
                strExmEmpCondition = " 1=1"
            End If
            Dim projectMaterialcond As String
            If strFromPRJ <> "" And strToPRJ <> "" Then
                strProjectCondition = " U_Z_PRJCODE between '" & strFromPRJ & "' and '" & strToPRJ & "'"
                projectMaterialcond = " T0.U_Z_PRJCODE between '" & strFromPRJ & "' and '" & strToPRJ & "'"
            ElseIf strFromPRJ <> "" And strToPRJ = "" Then
                strProjectCondition = " U_Z_PRJCODE >= '" & strFromPRJ & "'"
                projectMaterialcond = " T0.U_Z_PRJCODE >= '" & strFromPRJ & "'"
            ElseIf strFromPRJ = "" And strToPRJ <> "" Then
                strProjectCondition = " U_Z_PRJCODE <= '" & strToPRJ & "'"
                projectMaterialcond = " T0.U_Z_PRJCODE <= '" & strToPRJ & "'"
            Else
                strProjectCondition = " 1=1"
                projectMaterialcond = " 1=1"
            End If

            strProjectCondition = strProjectCondition
            projectMaterialcond = projectMaterialcond

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

            Dim strTimeAttendance, strLeaveCondition, strCondition, strGroupcondition, st1, st2, strJEPostingCondition, strLeaveCondition1, strExpPostCondition As String
            strCondition = strEMPCondition & " and " & strDateCondition & " and " & strProjectCondition & " Order by U_Z_EMPCODE,U_Z_DocDate,U_Z_PRJCODE"
            strTimeAttendance = strEMPCondition & " and " & strDateCondition & " and " & strProjectCondition
            strLeaveCondition1 = strEMPCondition & " and " & strDateCondition

            strExpPostCondition = strExmEmpCondition & " and " & strDateCondition & " and " & strProjectCondition & " Order by U_Z_EMPCODE,U_Z_DocDate,U_Z_PRJCODE"
            strLeaveCondition = strEMPCondition & " and " & strDateCondition & " Order by U_Z_EMPCODE,U_Z_DocDate"
            strJEPostingCondition = strEMPCondition & " and " & strDateCondition & " and " & strProjectCondition
            st2 = strEMPCondition & " and " & strDateCondition & " and " & strProjectCondition
            strGroupcondition = strEMPCondition & " and " & strDateCondition & " and " & strProjectCondition & "  group by T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], T1.[U_Z_PRJCODE], T1.[U_Z_PRCNAME]  Order by U_Z_EMPCODE,U_Z_PRJCODE"
            st1 = strEMPCondition & " and " & strDateCondition & " and " & strProjectCondition & "  group by T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], T1.[U_Z_PRJCODE], T1.[U_Z_ACTNAME]  Order by U_Z_EMPCODE,U_Z_PRJCODE"
            Dim st5 As String
            Dim strEMPIDS As String = oApplication.Utilities.getEmpIDforMangers_Reports(oApplication.Company.UserName)

            If (strreporttype.ToUpper = "EXP") Then

                strString = "SELECT case T1.[U_Z_EMPID] when '' then T1.U_Z_CNTID else T1.U_Z_EMPID end,  T1.[U_Z_EMPNAME], convert(varchar(10),T0.[U_Z_DOCDATE],105), T1.[U_Z_EXPNAME], T1.[U_Z_EXPTYPE], T1.[U_Z_PRJCODE],T1.[U_Z_PRJNAME],"
                strString = strString & " T1.U_Z_MODNAME,T1.U_Z_ACTNAME, T1.[U_Z_CURRENCY], T1.[U_Z_AMOUNT], T1.[U_Z_LocAmt], Convert(Decimal,(replace(T1.[U_Z_LocAmt],substring(T1.[U_Z_LocAmt],0,4),''))),Convert(Decimal,(replace(T1.[U_Z_SysAmt],substring(T1.[U_Z_SysAmt],0,4),''))), T1.[U_Z_REFCODE],T1.[U_Z_DisRule], T1.[U_Z_REMARKS], case T1.[U_Z_APPROVED] when 'A' then 'Approved'  when 'P' then 'Pending'  else 'Declined' end ,T1.[U_Z_DocNum] FROM [dbo].[@Z_OEXP]  T0 inner join  [dbo].[@Z_EXP1]  T1  on T0.Code=T1.U_Z_REFCODE  where (T1.U_Z_CNTID<>'' ) or  T1.U_Z_EMPID in (" & strEMPIDS & ") and " & strExpPostCondition
                strString1 = ""
                strString2 = ""
            ElseIf strreporttype.ToUpper = "TIME" Then
                strString = "SELECT T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], convert(varchar(10),T0.[U_Z_DOCDATE],105), T1.[U_Z_PRJCODE],T1.[U_Z_PRJNAME], T1.[U_Z_PRCNAME], T1.[U_Z_ACTNAME],  case isnull(T1.[U_Z_TYPE],'I') when 'I' then 'Item' else 'Resource' end 'U_Z_TYPE',T1.U_Z_BDGQTY,T1.U_Z_QUANTITY,T1.U_Z_MEASURE, T1.[U_Z_DATE], T1.[U_Z_HOURS],  T1.[U_Z_REFCODE],T1.[U_Z_REMARKS], case T1.[U_Z_APPROVED] when 'A' then 'Approved' when 'P' then 'Pending' else 'Declined' end 'STATUS' FROM [dbo].[@Z_OTIM]  T0  inner join  [dbo].[@Z_TIM1]  T1 on T0.Code=T1.U_Z_REFCODE   where T0.[U_Z_EMPCODE] in (" & strEMPIDS & ") and " & strCondition
                strString1 = "SELECT T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], T1.[U_Z_PRJCODE], T1.[U_Z_PRCNAME], sum(T1.[U_Z_HOURS]) FROM [dbo].[@Z_OTIM]  T0  inner join  [dbo].[@Z_TIM1]  T1 on T0.Code=T1.U_Z_REFCODE where T0.[U_Z_EMPCODE] in (" & strEMPIDS & ") and " & strGroupcondition
                strString2 = "SELECT T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], T1.[U_Z_PRJCODE], T1.[U_Z_ACTNAME], sum(T1.[U_Z_HOURS]) FROM [dbo].[@Z_OTIM]  T0  inner join  [dbo].[@Z_TIM1]  T1 on T0.Code=T1.U_Z_REFCODE where T0.[U_Z_EMPCODE] in (" & strEMPIDS & ") and " & st1
            ElseIf strreporttype.ToUpper = "TA" Then
                strString = "SELECT  T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], convert(varchar(10),T0.[U_Z_DOCDATE],105) 'Act.Date', ' ' 'U_Z_PRJCODE' ,' ' 'U_Z_PRJNAME' ,' ' 'U_Z_PRCNAME' ,' ' 'U_Z_ACTNAME', T0.[U_Z_TYPE], T0.[U_Z_FROMDATE]  ,T0.[U_Z_TODATE],T0.[U_Z_DAYS] 'Days', T0.[U_Z_DAYS] * 8 'Hours', T0.[U_Z_REASON] 'Reason /Remarks' ,  case T0.[U_Z_APPROVED] when 'A' then 'Approved' when 'P' then 'Pending' else 'Declined' end 'STATUS' FROM [dbo].[@Z_OLEV]  T0  where T0.[U_Z_EMPCODE] in (" & strEMPIDS & ") and " & strLeaveCondition1
                strString = strString & "   union all"
                strString = strString & " SELECT T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], convert(varchar(10),T1.[U_Z_DATE],105) 'Act.Date', T1.[U_Z_PRJCODE],T1.[U_Z_PRJNAME], T1.[U_Z_PRCNAME], T1.[U_Z_ACTNAME],  case isnull(T1.[U_Z_TYPE],'I') when 'I' then 'Item' else 'Resource' end 'U_Z_TYPE',T1.[U_Z_DATE] 'U_Z_FROMDATE',T1.[U_Z_DATE] 'U_Z_TODATE'  , T1.[U_Z_HOURS] /8  'Days', T1.[U_Z_HOURS] 'Hours',  T1.[U_Z_REMARKS] 'Reason /Remarks', case T1.[U_Z_APPROVED] when 'A' then 'Approved' when 'P' then 'Pending' else 'Declined' end 'STATUS' FROM [dbo].[@Z_OTIM]  T0  inner join  [dbo].[@Z_TIM1]  T1 on T0.Code=T1.U_Z_REFCODE   where T0.[U_Z_EMPCODE] in (" & strEMPIDS & ") and " & strTimeAttendance

            ElseIf strreporttype.ToUpper = "PRJ" Then
                strString = " SELECT T0.[U_Z_PRJCODE],T0.[U_Z_PRJName], T1.[U_Z_MODNAME], T1.[U_Z_ACTNAME], T1.[U_Z_EmpID],T1.[U_Z_Position], T1.[U_Z_HOURS],T1.[U_Z_HOURS],T1.[U_Z_HOURS] ,T1.[U_Z_AMOUNT],T1.[U_Z_AMOUNT],T1.[U_Z_AMOUNT],T1.U_Z_AMOUNT 'ReceivedAmt' FROM [dbo].[@Z_HPRJ]  T0  inner join  [dbo].[@Z_PRJ1]  T1 on T0.DocEntry=T1.DocEntry where " & strProjectCondition & " and " & strExmEmpCondition & "and ( T1.U_Z_EMPID<>'')"

                strString = " SELECT T0.[U_Z_PRJCODE],T0.[U_Z_PRJName], T1.[U_Z_MODNAME], T1.[U_Z_ACTNAME], T1.[U_Z_EmpID],T1.[U_Z_Position], T1.[U_Z_HOURS],T1.[U_Z_HOURS],T1.[U_Z_HOURS] ,T1.[U_Z_AMOUNT],T1.[U_Z_AMOUNT],T1.[U_Z_AMOUNT],T1.U_Z_AMOUNT 'InvoicedAmt' FROM [dbo].[@Z_HPRJ]  T0  inner join  [dbo].[@Z_PRJ1]  T1 on T0.DocEntry=T1.DocEntry where " & strProjectCondition & " and " & strExmEmpCondition & " and (  T1.U_Z_EMPID<>'')"
                strString = " SELECT T0.[U_Z_PRJCODE],T0.[U_Z_PRJName], T1.[U_Z_MODNAME], T1.[U_Z_ACTNAME], case T1.[U_Z_EMPID] when '' then T1.U_Z_CNTID else T1.U_Z_EMPID end 'U_Z_EmpID',T1.[U_Z_Position], Sum(T1.[U_Z_HOURS]),Sum(T1.[U_Z_HOURS]),Sum(T1.[U_Z_HOURS]) ,Sum(T1.[U_Z_AMOUNT]),Sum(T1.[U_Z_AMOUNT]),Sum(T1.[U_Z_AMOUNT]),Sum(T1.[U_Z_AMOUNT]) 'InvoicedAmt',T1.U_Z_CNTID,T1.[U_Z_EmpID] 'Emp' FROM [dbo].[@Z_HPRJ]  T0  inner join  [dbo].[@Z_PRJ1]  T1 on T0.DocEntry=T1.DocEntry  where T0.U_Z_STATUS='X'  and    " & strProjectCondition & " and " & strExmEmpCondition & " and (  T1.U_Z_EMPID<>'') group by T0.[U_Z_PRJCODE],T0.[U_Z_PRJName], T1.[U_Z_MODNAME], T1.[U_Z_ACTNAME], T1.[U_Z_EmpID],T1.U_Z_CNTID,T1.[U_Z_Position]"
            ElseIf strreporttype.ToUpper = "PRJM" Then
                strString = " select T0.U_Z_PrjCode 'Project',T0.U_Z_PrjName 'PrjName',T1.U_Z_MODNAME 'Business',T1.U_Z_ACTNAME 'Activity',isnull(T2.U_Z_ItemCode,'') 'ItemCode'"
                strString = strString & " ,T2.U_Z_ItemName 'ItemName',T2.U_Z_UOM 'UOM',T2.U_Z_REQQTY 'Qty',T2.U_Z_ESTCOST 'Cost' ,T2.U_Z_REQQTY-T2.U_Z_REQQTY 'OrdQty',"
                strString = strString & "   T2.U_Z_ESTCOST-T2.U_Z_ESTCOST 'OrdCost',T2.U_Z_REQQTY-T2.U_Z_REQQTY 'RecQty',"
                strString = strString & "   T2.U_Z_ESTCOST-T2.U_Z_ESTCOST 'RecCost',T2.U_Z_REQQTY-T2.U_Z_REQQTY 'RelQty',"
                      strString = strString & "   T2.U_Z_ESTCOST-T2.U_Z_ESTCOST 'RelCost' from [@Z_PRJ1] T1 inner Join [@Z_HPRJ] T0 on T0.DOcEntry=T1.DocEntry left outer Join [@Z_PRJ2] T2 on T2.U_Z_BOQREF=T1.U_Z_BOQ where T0.U_Z_STATUS='X' and  " & projectMaterialcond
            ElseIf strreporttype.ToUpper = "PRJS" Then
                   strString = " T1.U_Z_Quantity 'Qty',0.00 'Percentage', case T1.U_Z_Status when 'I' then 'In Process' when 'C' then 'Completed' else 'Pending' end 'Status',T1.U_Z_CMPDATE 'CmpDate'"
                st5 = " Select  T0.[U_Z_PRJCODE] 'Project',T0.[U_Z_PRJName] 'ProjectName', 'Expenses' 'Business', 'Expenses' 'Activity', ' ' 'EmpID',' ' 'Position',  T0.U_Z_FROMDATE  'FromDate',T0.[U_Z_ToDate] 'EndDate',"
                st5 = st5 & " T0.[U_Z_TotalExpense] 'EstimatedHours', T0.[U_Z_TotalExpense] 'ActualHours', T0.[U_Z_TotalExpense] 'VarianceHours' ,  0 'Qty',0.00 'Percentage', ' ' 'Status',U_Z_FROMDATE 'CmpDate',0 'EstimatedCost',0 'MatrialOrdered',0 'MaterialReceived',0 'OrderPending',0 'MaterialPending', 0 'Cost Variance' ,0.0 'Invoiced Amount' FROM [dbo].[@Z_HPRJ]  T0 where " & strProjectCondition
                 strString = " SELECT T0.[U_Z_PRJCODE] 'Project',T0.[U_Z_PRJName] 'ProjectName', T1.[U_Z_MODNAME] 'Business', T1.[U_Z_ACTNAME] 'Activity', case T1.U_Z_Type when 'I' then 'Item' when 'R' then 'Resource' else 'Expenses' end 'Type',T1.[U_Z_EmpID] 'EmpID',T1.[U_Z_Position] 'Position',  T1.[U_Z_FromDate] 'FromDate',T1.[U_Z_ToDate] 'EndDate',T1.[U_Z_HOURS] 'EstimatedHours',T1.[U_Z_HOURS] 'ActualHours',T1.[U_Z_HOURS] 'VarianceHours' , " & strString & ",T1.[U_Z_AMOUNT] 'EstimatedCost',T1.[U_Z_Quantity]-T1.[U_Z_Quantity] 'MatrialOrdered',T1.[U_Z_Quantity]-T1.[U_Z_Quantity] 'MaterialReceived',T1.[U_Z_Quantity]-T1.[U_Z_Quantity] 'OrderPending',T1.[U_Z_Quantity]-T1.[U_Z_Quantity] 'MaterialPending', T1.[U_Z_Quantity]-T1.[U_Z_Quantity] 'Cost Variance',T1.U_Z_AMOUNT-T1.U_Z_AMOUNT 'ReceivedAmt' FROM [dbo].[@Z_HPRJ]  T0  inner join  [dbo].[@Z_PRJ1]  T1 on T0.DocEntry=T1.DocEntry where T0.U_Z_STATUS='X' and  T1.U_Z_EMPID<>'' and " & strProjectCondition & " and  " & strExmEmpCondition
              ElseIf strreporttype.ToUpper = "POSTING" Then
                '  strString = "SELECT T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], convert(varchar(10),T0.[U_Z_DOCDATE],105), T1.[U_Z_EXPNAME], T1.[U_Z_EXPTYPE], T1.[U_Z_PRJCODE],T1.[U_Z_PRJNAME],"
                ' strString = strString & " T1.[U_Z_CURRENCY], T1.[U_Z_AMOUNT], T1.[U_Z_LocAmt], (replace(T1.[U_Z_LocAmt],T1.[U_Z_CURRENCY],'')), T1.[U_Z_REFCODE], T1.[U_Z_REMARKS], 'N' 'Select' ,T1.Code FROM [dbo].[@Z_OEXP]  T0 inner join  [dbo].[@Z_EXP1]  T1  on T0.Code=T1.U_Z_REFCODE  where T1.[U_Z_APPROVED]='A' and isnull(T1.[U_Z_DocNum],'')='' and " & strCondition
                strString = "SELECT T1.[U_Z_EMPID] 'U_Z_EMPCODE', T1.[U_Z_EMPNAME], convert(varchar(10),T0.[U_Z_DOCDATE],105), T1.[U_Z_EXPNAME], T1.[U_Z_EXPTYPE], T1.[U_Z_PRJCODE],T1.[U_Z_PRJNAME],"
                strString = strString & " T1.[U_Z_CURRENCY], T1.[U_Z_AMOUNT], T1.[U_Z_LocAmt], (replace(T1.[U_Z_LocAmt],T1.[U_Z_CURRENCY],'')), T1.[U_Z_REFCODE],T1.[U_Z_DisRule], T1.[U_Z_REMARKS], 'N' 'Select' ,T1.Code FROM [dbo].[@Z_OEXP]  T0 inner join  [dbo].[@Z_EXP1]  T1  on T0.Code=T1.U_Z_REFCODE  where T1.[U_Z_APPROVED]='A' and isnull(T1.[U_Z_DocNum],'')='' and " & strExpPostCondition
            ElseIf strreporttype.ToUpper = "POSTJE" Then
                strString = "SELECT  T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], convert(varchar(10),T0.[U_Z_DOCDATE],105),T1.[U_Z_EXPNAME], T1.[U_Z_PRJCODE],T1.[U_Z_PRJNAME],T3.[U_Z_ActCode],T4.[AcctName],T1.[U_Z_LocAmount],T1.[U_Z_DisRule],T1.[U_Z_Ref1],T1.[U_Z_Remarks]  FROM [dbo].[@Z_OEXP]  T0 inner join  [dbo].[@Z_EXP1]  T1  on T0.Code=T1.U_Z_REFCODE inner join [@Z_Expances] T3 on T3.U_Z_ExpName=T1.[U_Z_ExpName] inner join OACT T4 on T4.FormatCode=T3.[U_Z_ActCode]"
                strString = strString & " where T1.[U_Z_APPROVED]='A' and isnull(T1.[U_Z_DocNum],'')='' and  T1.[U_Z_Flag]='Y' and " & strJEPostingCondition & " Order by T0.U_Z_DocDate, T0.[U_Z_EMPCODE],T1.[U_Z_EXPNAME], T1.[U_Z_PRJCODE] "
            ElseIf strreporttype.ToUpper = "LEV" Then
                strString = "SELECT T0.[Code], T0.[Name], T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], T0.[U_Z_DOCDATE], T0.[U_Z_SUBEMP], T0.[U_Z_TYPE], T0.[U_Z_FROMDATE],T0.[U_Z_TODATE], T0.[U_Z_DAYS], T0.[U_Z_REASON], T0.[U_Z_APPROVED], T0.[U_Z_REMARKS] FROM [dbo].[@Z_OLEV]  T0"
                strString = strString & " where " & strLeaveCondition
                'strString = "Select * from [@Z_OLEV]  where " & strLeaveCondition
            Else
                strString = ""
            End If

            If strString = "" Then
                aForm.Items.Item("7").Visible = False
                'aForm.Items.Item("17").Visible = False
                aForm.Freeze(False)
                Exit Sub
            Else
                aForm.Items.Item("7").Visible = True
                ' aForm.Items.Item("17").Visible = True
            End If
            oGrid = aForm.Items.Item("7").Specific
            oGrid.DataTable.ExecuteQuery(strString)
            FormatGrids(aForm, strreporttype, strString, strString1, strString2)
            If strreporttype.ToUpper = "PRJS" Then
                DisplayProjectwiseReport_Summary(oGrid, st2)
            ElseIf strreporttype.ToUpper = "PRJM" Then
                DisplayProjectwiseReport_Material(oGrid, st2)
            ElseIf strreporttype.ToUpper = "PRJ" Then
                DisplayProjectwiseReport(oGrid, st2)
            End If
            'DatabindSummary(aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub

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
        Dim strSql, strProject, strProcess, strActivity, strEmpID As String
        Dim dblEstimated, dblActual, dblVariance As Double
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            strProject = oGrid.DataTable.GetValue(0, intRow)
            If strProject <> "" Then
                strProcess = oGrid.DataTable.GetValue(2, intRow)
                strActivity = oGrid.DataTable.GetValue(3, intRow)
                ' dblEstimated = oGrid.DataTable.GetValue(4, intRow)
                '   strEmpID = oGrid.DataTable.GetValue("U_Z_EmpID", intRow)
                strEmpID = oGrid.DataTable.GetValue("Emp", intRow)
                dblEstimated = oGrid.DataTable.GetValue(6, intRow)
                Dim strCNTID As String = oGrid.DataTable.GetValue("U_Z_CNTID", intRow)
                ' strSql = "SELECT T0.[U_Z_PRJCODE],T0.[U_Z_PRJNAME], T0.[U_Z_PRJNAME], T1.[U_Z_MODNAME], T1.[U_Z_ACTNAME], T1.[U_Z_DAYS], T1.[U_Z_HOURS] FROM [dbo].[@Z_HPRJ]  T0  inner join  [dbo].[@Z_PRJ1]  T1 on T0.DocEntry=T1.DocEntry"
                'strSql = strSql & " where T0.[U_Z_PRJCODE]='" & strProject & "' and T1.[U_Z_MODNAME]='" & strProcess & "' and T1.[U_Z_ACTNAME]='" & strActivity & "'"
                '  strSql = "SELECT  T1.[U_Z_PRJCODE],T1.[U_Z_PRJCODE], T1.[U_Z_PRCNAME], T1.[U_Z_ACTNAME],Sum(T1.U_Z_HOURS)  FROM [dbo].[@Z_OTIM]  T0  inner join  [dbo].[@Z_TIM1]  T1 on T0.Code=T1.U_Z_REFCODE and T1.[U_Z_APPROVED]='A'   where " & strCondition & " and T1.[U_Z_PRJCODE]='" & strProject.Replace("'", "''") & "' and T1.[U_Z_PRCNAME]='" & strProcess.Replace("'", "''") & "' and T1.[U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' group by T1.U_Z_PRJCODE,T1.U_Z_PRCNAME,T1.U_Z_ACTNAME "
                If strCNTID <> "" Then
                    strSql = "SELECT  T1.[U_Z_PRJCODE],T1.[U_Z_PRJCODE], T1.[U_Z_PRCNAME], T1.[U_Z_ACTNAME],Sum(T1.U_Z_HOURS)  FROM [dbo].[@Z_OTIM]  T0  inner join  [dbo].[@Z_TIM1]  T1 on T0.Code=T1.U_Z_REFCODE and T1.[U_Z_APPROVED]='A'   where T0.U_Z_EMPCODE='" & strEmpID & "' and  " & strCondition & " and T1.[U_Z_PRJCODE]='" & strProject.Replace("'", "''") & "' and T1.[U_Z_PRCNAME]='" & strProcess.Replace("'", "''") & "' and T1.[U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' group by T1.U_Z_PRJCODE,T1.U_Z_PRCNAME,T1.U_Z_ACTNAME "
                Else
                    strSql = "SELECT  T1.[U_Z_PRJCODE],T1.[U_Z_PRJCODE], T1.[U_Z_PRCNAME], T1.[U_Z_ACTNAME],Sum(T1.U_Z_HOURS)  FROM [dbo].[@Z_OTIM]  T0  inner join  [dbo].[@Z_TIM1]  T1 on T0.Code=T1.U_Z_REFCODE and T1.[U_Z_APPROVED]='A'   where T0.U_Z_EMPCODE='" & strEmpID & "' and  " & strCondition & " and T1.[U_Z_PRJCODE]='" & strProject.Replace("'", "''") & "' and T1.[U_Z_PRCNAME]='" & strProcess.Replace("'", "''") & "' and T1.[U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' group by T1.U_Z_PRJCODE,T1.U_Z_PRCNAME,T1.U_Z_ACTNAME "

                End If
                 oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp.DoQuery(strSql)
                If oTemp.RecordCount > 0 Then
                    ' dblActual = oTemp.Fields.Item(4).Value
                    dblActual = oTemp.Fields.Item(4).Value
                    dblVariance = dblEstimated - dblActual
                Else
                    dblActual = 0
                    dblVariance = dblEstimated - dblActual
                End If

                oGrid.DataTable.SetValue(7, intRow, (dblActual))
                oGrid.DataTable.SetValue(8, intRow, (dblVariance))




                Dim stEmpID As String
                Dim oTemprs As SAPbobsCOM.Recordset
                Dim dblAcutalcost, dblVarianceCost, dblestimatedcost As Double
                oTemprs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strSql = "SELECT T0.[MainCurncy] FROM OADM T0"
                oTemprs.DoQuery(strSql)
                Dim strLocalCurrency As String = oTemprs.Fields.Item(0).Value

                Dim dblestimatedExpense, dblactualexpenses, dbleVarianceExpenses As Double
                If 1 = 1 Then 'Actual Cost Time sheet

                    strSql = "select sum(isnull(T3.U_HR_RATE,0) * isnull(T1.U_Z_HOURS,0) /1) from [@Z_OTIM] T0 inner join [@Z_TIM1] T1 on T1.U_Z_REFCODE=T0.Code inner Join OHEM T3 on T3.empID=T0.U_Z_EMPCODE "
                    strSql = strSql & "  where T0.U_Z_EMPCODE='" & strEmpID & "' and  T1.U_Z_TYPE='R' and T1.U_Z_PRJCODE='" & strProject & "'  and T1.U_Z_PRCNAME='" & strProcess & "' and T1.U_Z_ACTNAME='" & strActivity & "'"
                    oTemprs.DoQuery(strSql)
                    Dim dblTotalProjectCost, dblActualProjectCost, dblVarianceProjectCost As Double
                    dblAcutalcost = oTemprs.Fields.Item(0).Value
                    dblActualProjectCost = dblAcutalcost
                    If strCNTID <> "" Then
                        strSql = "select isnull(Sum(Convert(Decimal,REPLACE(U_Z_LOCAMT,'" & strLocalCurrency & "',''))),0) from [@Z_EXP1]  where U_Z_CNTID='" & strCNTID & "' and U_Z_MODNAME='" & strProcess & "' and U_Z_ACTNAME='" & strActivity & "' and  U_Z_Approved='A' and  U_Z_PRJCODE ='" & strProject & "'"
                    Else
                        strSql = "select isnull(Sum(Convert(Decimal,REPLACE(U_Z_LOCAMT,'" & strLocalCurrency & "',''))),0) from [@Z_EXP1]  where U_Z_EMPID='" & strEmpID & "' and U_Z_MODNAME='" & strProcess & "' and U_Z_ACTNAME='" & strActivity & "' and  U_Z_Approved='A' and  U_Z_PRJCODE ='" & strProject & "'"

                    End If
                    oTemprs.DoQuery(strSql)
                    If oTemprs.RecordCount > 0 Then
                        dblactualexpenses = oTemprs.Fields.Item(0).Value
                    Else
                        dblactualexpenses = 0
                    End If
                    dblAcutalcost = dblAcutalcost + dblactualexpenses
                    dblEstimated = oGrid.DataTable.GetValue(9, intRow)
                    dblVariance = dblEstimated - dblAcutalcost
                    oGrid.DataTable.SetValue(10, intRow, (dblAcutalcost))
                    oGrid.DataTable.SetValue(11, intRow, (dblVariance))
                End If
                'End Acutal Cost Time sheet



                ' strSql = "SELECT   U_Z_ACTNAME,SUm(LineTOtal),Project from PCH1 where  U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' group by PROJECT,U_Z_ACTNAME "


                'strSql = "select sum(LineTotal) from ("
                'strSql = strSql & " select LineTotal from PCH1 where  baseType<>20 and     U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                'strSql = strSql & " select LineTotal 'LineTotal' from PDN1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                'strSql = strSql & " select LineTotal *-1 'LineTotal' from RPD1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                ''strSql = strSql & " select Debit-Credit 'LineTotal' from JDT1 T0 inner Join OJDT T1 on T1.TransId=T0.TransId where T1.TransType=30 and T0.U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  T0.[PROJECT]='" & strProject.Replace("'", "''") & "' and  T0.[U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                'strSql = strSql & " select LineTotal *-1 'LineTotal' from RPC1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' ) as x"


                'Dim dblJDT As Double
                'oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oTemp.DoQuery(strSql)

                ''dblEstimated = oGrid.DataTable.GetValue(7, intRow)
                'dblEstimated = oGrid.DataTable.GetValue(9, intRow)

                'If oTemp.RecordCount > 0 Then
                '    ' dblActual = oTemp.Fields.Item(1).Value
                '    dblActual = oTemp.Fields.Item(0).Value
                '    '   dblVariance = dblEstimated - dblActual
                '    ' oGrid.DataTable.SetValue(1, intRow, oTemp.Fields.Item(1).Value)
                'Else
                '    dblActual = 0
                '    '  dblVariance = dblEstimated - dblActual
                'End If

                'dblActual = dblActual + oGrid.DataTable.GetValue(10, intRow)



                'strSql = " SELECT U_Z_MDNAME, U_Z_ACTNAME , SUm(Debit-Credit),Project from JDT1 where  U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' group by PROJECT,U_Z_MDNAME,U_Z_ACTNAME "

                'oTemp.DoQuery(strSql)
                'If oTemp.RecordCount > 0 Then
                '    dblActual = dblActual + oTemp.Fields.Item(2).Value
                'Else
                '    dblActual = dblActual
                'End If
                'dblVariance = dblEstimated - dblActual
                ''oGrid.DataTable.SetValue(8, intRow, (dblActual))
                '' oGrid.DataTable.SetValue(9, intRow, (dblVariance))
                'oGrid.DataTable.SetValue(10, intRow, (dblActual))
                'oGrid.DataTable.SetValue(11, intRow, (dblVariance))


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
        Dim strSql, strProject, strProcess, strActivity, strType, strEMPID As String
        Dim dblEstimated, dblActual, dblVariance As Double
        Dim dblMaterialOrdered, dblMaterialReceived, dblOrderPending, dblMaterialPending, dblCostVariance, dblestimatedcost As Double
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            ' strProject = oGrid.DataTable.GetValue(0, intRow)
            strProject = oGrid.DataTable.GetValue("Project", intRow)
            If strProject <> "" Then
                strProcess = oGrid.DataTable.GetValue("Business", intRow)
                strActivity = oGrid.DataTable.GetValue("Activity", intRow)
                strType = oGrid.DataTable.GetValue("Type", intRow)
                strEMPID = oGrid.DataTable.GetValue("EmpID", intRow)
                ' dblEstimated = oGrid.DataTable.GetValue(4, intRow)
                dblEstimated = oGrid.DataTable.GetValue("EstimatedHours", intRow)
                dblestimatedcost = oGrid.DataTable.GetValue("EstimatedCost", intRow)
                Dim stEmpID As String
                Dim oTemprs As SAPbobsCOM.Recordset
                Dim dblAcutalcost, dblVarianceCost As Double
                oTemprs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


                If strType = "Resource" Then 'Actual Cost Time sheet
                    strSql = "select sum(isnull(T3.U_HR_RATE,0) * isnull(T1.U_Z_HOURS,0) /1) from [@Z_OTIM] T0 inner join [@Z_TIM1] T1 on T1.U_Z_REFCODE=T0.Code inner Join OHEM T3 on T3.empID=T0.U_Z_EMPCODE "
                    strSql = strSql & "  where T0.U_Z_EMPCODE='" & strEMPID & "' and  T1.U_Z_TYPE='R' and T1.U_Z_PRJCODE='" & strProject & "'  and T1.U_Z_PRCNAME='" & strProcess & "' and T1.U_Z_ACTNAME='" & strActivity & "'"
                    oTemprs.DoQuery(strSql)
                    Dim dblTotalProjectCost, dblActualProjectCost, dblVarianceProjectCost As Double
                    dblAcutalcost = oTemprs.Fields.Item(0).Value
                    dblActualProjectCost = dblAcutalcost
                    dblestimatedcost = oGrid.DataTable.GetValue("EstimatedCost", intRow)
                    dblTotalProjectCost = dblestimatedcost
                    dblVarianceCost = dblestimatedcost - dblAcutalcost
                    aGrid.DataTable.SetValue("MatrialOrdered", intRow, oTemprs.Fields.Item(0).Value)
                    aGrid.DataTable.SetValue("MaterialReceived", intRow, dblVarianceCost)
                Else
                    Dim dblestimatedExpense, dblactualexpenses, dbleVarianceExpenses As Double
                    'strSql = " select T0.U_Z_PRJCODE, (U_Z_AMOUNT) 'Count' from  [@Z_HPRJ]  T0 inner Join [@Z_PRJ1] T1 on T1.DocEntry=T0.DocEntry where  U_Z_TYPE='E' and T1.U_Z_PRJCODE='" & strProject & "'  and T1.U_Z_PRCNAME='" & strProcess & "' and T1.U_Z_ACTNAME='" & strActivity & "'"
                    'oTemprs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'oTemprs.DoQuery(strSql)
                    'If oTemprs.RecordCount > 0 Then
                    '    dblestimatedcost = oTemprs.Fields.Item(1).Value
                    'Else
                    '    dblestimatedcost = 0
                    'End If
                   
                    dblestimatedcost = dblestimatedcost
                    strSql = "SELECT T0.[MainCurncy] FROM OADM T0"
                    oTemprs.DoQuery(strSql)
                    Dim strLocalCurrency As String = oTemprs.Fields.Item(0).Value

                    strSql = "select isnull(Sum(Convert(Decimal,REPLACE(U_Z_LOCAMT,'" & strLocalCurrency & "',''))),0) from [@Z_EXP1]  where U_Z_EMPID='" & strEMPID & "' and U_Z_MODNAME='" & strProcess & "' and U_Z_ACTNAME='" & strActivity & "' and  U_Z_Approved='A' and  U_Z_PRJCODE ='" & strProject & "'"
                    oTemprs.DoQuery(strSql)
                    If oTemprs.RecordCount > 0 Then
                        dblactualexpenses = oTemprs.Fields.Item(0).Value
                    Else
                        dblactualexpenses = 0
                    End If
                    'dblTotalProjectCost = dblTotalProjectCost + dblestimatedcost
                    'dblActualProjectCost = dblActualProjectCost + dblactualexpenses
                    dblVarianceCost = dblestimatedcost - dblactualexpenses
                    aGrid.DataTable.SetValue("MatrialOrdered", intRow, dblactualexpenses)
                    aGrid.DataTable.SetValue("MaterialReceived", intRow, dblVarianceCost)
                End If
                'End Acutal Cost Time sheet

                If strType <> "Expenses" Then
                    stEmpID = oGrid.DataTable.GetValue("EmpID", intRow)
                    strSql = "SELECT  T1.[U_Z_PRJCODE],T1.[U_Z_PRJCODE], T1.[U_Z_PRCNAME], T1.[U_Z_ACTNAME],Sum(T1.U_Z_HOURS),sum(T1.U_Z_Quantity)  FROM [dbo].[@Z_OTIM]  T0  inner join  [dbo].[@Z_TIM1]  T1 on T0.Code=T1.U_Z_REFCODE and T1.[U_Z_APPROVED]='A'   where " & strCondition & " and T0.[U_Z_EMPCODe]='" & stEmpID & "' and  T1.[U_Z_PRJCODE]='" & strProject.Replace("'", "''") & "' and T1.[U_Z_PRCNAME]='" & strProcess.Replace("'", "''") & "' and T1.[U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' group by T1.U_Z_PRJCODE,T1.U_Z_PRCNAME,T1.U_Z_ACTNAME "
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

                If strType = "Expenses" Then
                    dblActual = 0

                    aGrid.DataTable.SetValue("ActualHours", intRow, dblActual)
                    aGrid.DataTable.SetValue("VarianceHours", intRow, dblActual)
                End If

                'strSql = "SELECT   U_Z_ACTNAME,SUm(LineTOtal),Project from PCH1 where  U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' group by PROJECT,U_Z_ACTNAME "
                'oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oTemp.DoQuery(strSql)
                'Dim strBOQRef As String
                'strSql = "Select isnull(T0.U_Z_BOQ,''),* from [@Z_PRJ1] T0 inner Join [@Z_HPRJ] T1 on T0.DocEntry=T1.DocEntry where T1.U_Z_PrjCode='" & strProject.Replace("'", "''") & "' and T0.U_Z_MODNAME='" & strProcess.Replace("'", "''") & "' and U_Z_Actname='" & strActivity.Replace("'", "''") & "'"
                'oTemp.DoQuery(strSql)
                'If oTemp.RecordCount > 0 Then
                '    strBOQRef = oTemp.Fields.Item(0).Value
                'Else
                '    strBOQRef = ""
                'End If
                ''dblEstimated = oGrid.DataTable.GetValue(7, intRow)
                'dblEstimated = oGrid.DataTable.GetValue("EstimatedCost", intRow)

                '' strBOQRef = ""
                'If strBOQRef <> "" Then
                '    strSql = "Select sum(LineTotal) from POR1 T0 inner Join OPOR T1 on T1.DocEntry=T0.DocEntry and T1.CANCELED <>'Y' where T0.U_Z_BOQREF='" & strBOQRef & "'"
                '    'strSql = "Select sum(LineTotal) from POR1 T0 inner Join OPOR T1 on T1.DocEntry=T0.DocEntry and T1.CANCELED <>'Y' where T0.U_Z_BOQREF='" & strBOQRef & "'"
                '    oTemp.DoQuery(strSql)
                '    If oTemp.RecordCount > 0 Then
                '        dblMaterialOrdered = oTemp.Fields.Item(0).Value
                '    Else
                '        dblMaterialOrdered = 0
                '        ' dblVariance = dblEstimated - dblActual
                '    End If
                '    oGrid.DataTable.SetValue("MatrialOrdered", intRow, dblMaterialOrdered)
                '    strSql = "select sum(LineTotal) from ("
                '    strSql = strSql & " select LineTotal from PCH1 where  baseType<>20 and  isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                '    strSql = strSql & " select LineTotal 'LineTotal' from PDN1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                '    strSql = strSql & " select LineTotal *-1 'LineTotal' from RPD1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                '    strSql = strSql & " select Debit-Credit 'LineTotal' from JDT1 T0 inner Join OJDT T1 on T1.TransId=T0.TransId where T1.TransType=30 and isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                '    strSql = strSql & " select LineTotal *-1 'LineTotal' from RPC1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' ) as x"
                '    oTemp.DoQuery(strSql)

                '    If oTemp.RecordCount > 0 Then
                '        dblMaterialReceived = oTemp.Fields.Item(0).Value
                '    Else
                '        dblMaterialReceived = 0
                '        ' dblVariance = dblEstimated - dblActual
                '    End If
                '    oGrid.DataTable.SetValue("MaterialReceived", intRow, dblMaterialReceived)



                '    'strSql = "Select sum(LineTotal) from POR1 T0 inner Join OPOR T1 on T1.DocEntry=T0.DocEntry and T1.CANCELED <>'Y' where T0.U_Z_BOQREF='" & strBOQRef & "'"
                '    'oTemp.DoQuery(strSql)
                '    'If oTemp.RecordCount > 0 Then
                '    '    dblOrderPending = oTemp.Fields.Item(0).Value
                '    'Else
                '    '    dblOrderPending = 0
                '    '    ' dblVariance = dblEstimated - dblActual
                '    'End If
                '    dblOrderPending = dblestimatedcost - dblMaterialOrdered
                '    oGrid.DataTable.SetValue("OrderPending", intRow, dblOrderPending)
                '    strSql = "select sum(LineTotal) from ("
                '    strSql = strSql & " select LineTotal from POR1 T0 inner Join OPOR T1  on T1.DocEntry=T0.DocEntry and T1.DocStatus='O' where isnull(U_Z_BOQREF,'')='" & strBOQRef & "'"
                '    'strSql = strSql & " select LineTotal from PDN1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                '    'strSql = strSql & " select LineTotal *-1 from RPD1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                '    'strSql = strSql & " select LineTotal *-1 from RPC1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "'
                '    strSql = strSql & " ) as x"
                '    oTemp.DoQuery(strSql)

                '    If oTemp.RecordCount > 0 Then
                '        dblMaterialPending = oTemp.Fields.Item(0).Value
                '    Else
                '        dblMaterialPending = 0
                '        ' dblVariance = dblEstimated - dblActual
                '    End If
                '    oGrid.DataTable.SetValue("MaterialPending", intRow, dblMaterialPending)
                '    dblCostVariance = dblestimatedcost - dblMaterialReceived
                '    oGrid.DataTable.SetValue("Cost Variance", intRow, dblCostVariance)


                '    strSql = "select sum(LineTotal) from ("
                '    strSql = strSql & " select LineTotal from inv1 where    isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                '    '   strSql = strSql & " select LineTotal 'LineTotal' from PDN1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                '    '  strSql = strSql & " select LineTotal *-1 'LineTotal' from RPD1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                '    '   strSql = strSql & " select Debit-Credit 'LineTotal' from JDT1 T0 inner Join OJDT T1 on T1.TransId=T0.TransId where T1.TransType=30 and isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                '    strSql = strSql & " select LineTotal *-1 'LineTotal' from RIN1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' ) as x"
                '    oTemp.DoQuery(strSql)

                '    If oTemp.RecordCount > 0 Then
                '        dblMaterialReceived = oTemp.Fields.Item(0).Value
                '    Else
                '        dblMaterialReceived = 0
                '        ' dblVariance = dblEstimated - dblActual
                '    End If
                '    oGrid.DataTable.SetValue("ReceivedAmt", intRow, dblMaterialReceived)

                'Else
                '    strSql = "Select sum(LineTotal) from POR1 T0 inner Join OPOR T1 on T1.DocEntry=T0.DocEntry and T1.CANCELED <>'Y' where    U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  T0.Project='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' group by T0.PROJECT,U_Z_ACTNAME "
                '    oTemp.DoQuery(strSql)
                '    If oTemp.RecordCount > 0 Then
                '        dblMaterialOrdered = oTemp.Fields.Item(0).Value
                '    Else
                '        dblMaterialOrdered = 0
                '        ' dblVariance = dblEstimated - dblActual
                '    End If
                '    oGrid.DataTable.SetValue("MatrialOrdered", intRow, dblMaterialOrdered)
                '    strSql = "select sum(LineTotal) from ("
                '    strSql = strSql & " select LineTotal from PCH1 where  baseType<>20 and     U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                '    strSql = strSql & " select LineTotal 'LineTotal' from PDN1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                '    strSql = strSql & " select LineTotal *-1 'LineTotal' from RPD1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                '    strSql = strSql & " select Debit-Credit 'LineTotal' from JDT1 T0 inner Join OJDT T1 on T1.TransId=T0.TransId where T1.TransType=30 and T0.U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  T0.[PROJECT]='" & strProject.Replace("'", "''") & "' and  T0.[U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' union all"
                '    strSql = strSql & " select LineTotal *-1 'LineTotal' from RPC1 where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  [PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "' ) as x"
                '    oTemp.DoQuery(strSql)

                '    If oTemp.RecordCount > 0 Then
                '        dblMaterialReceived = oTemp.Fields.Item(0).Value
                '    Else
                '        dblMaterialReceived = 0
                '        ' dblVariance = dblEstimated - dblActual
                '    End If
                '    oGrid.DataTable.SetValue("MaterialReceived", intRow, dblMaterialReceived)



                '    'strSql = "Select sum(LineTotal) from POR1 T0 inner Join OPOR T1 on T1.DocEntry=T0.DocEntry and T1.CANCELED <>'Y' where T0.U_Z_BOQREF='" & strBOQRef & "'"
                '    'oTemp.DoQuery(strSql)
                '    'If oTemp.RecordCount > 0 Then
                '    '    dblOrderPending = oTemp.Fields.Item(0).Value
                '    'Else
                '    '    dblOrderPending = 0
                '    '    ' dblVariance = dblEstimated - dblActual
                '    'End If
                '    dblOrderPending = dblestimatedcost - dblMaterialOrdered
                '    oGrid.DataTable.SetValue("OrderPending", intRow, dblOrderPending)
                '    strSql = "select sum(LineTotal) from ("
                '    strSql = strSql & " select LineTotal from POR1 T0 inner Join OPOR T1  on T1.DocEntry=T0.DocEntry and T1.DocStatus='O' where U_Z_MDNAME='" & strProcess.Replace("'", "''") & "' and  T0.[PROJECT]='" & strProject.Replace("'", "''") & "' and  [U_Z_ACTNAME]='" & strActivity.Replace("'", "''") & "'"
                '    'strSql = strSql & " select LineTotal from PDN1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                '    'strSql = strSql & " select LineTotal *-1 from RPD1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "' union all"
                '    'strSql = strSql & " select LineTotal *-1 from RPC1 where isnull(U_Z_BOQREF,'')='" & strBOQRef & "'
                '    strSql = strSql & " ) as x"
                '    oTemp.DoQuery(strSql)

                '    If oTemp.RecordCount > 0 Then
                '        dblMaterialPending = oTemp.Fields.Item(0).Value
                '    Else
                '        dblMaterialPending = 0
                '        ' dblVariance = dblEstimated - dblActual
                '    End If
                '    oGrid.DataTable.SetValue("MaterialPending", intRow, dblMaterialPending)
                '    dblCostVariance = dblestimatedcost - dblMaterialReceived
                '    oGrid.DataTable.SetValue("Cost Variance", intRow, dblCostVariance)

                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
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
            oGrid.Columns.Item(0).TitleObject.Caption = "Employee Code / Contract ID"
            oEditTextColumn = oGrid.Columns.Item(0)
            oEditTextColumn.LinkedObjectType = "171"
            oGrid.Columns.Item(0).TitleObject.Sortable = True
            oGrid.Columns.Item(0).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(1).TitleObject.Caption = "Employee Name / Contractor Name"
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
            oGrid.Columns.Item(11).TitleObject.Caption = "Amount in Local Currency (" & LocalCurrency & ")"
            oGrid.Columns.Item(12).TitleObject.Caption = "Amount in Local Currency (" & LocalCurrency & ")"
            oEditTextColumn = oGrid.Columns.Item(12)
            oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oGrid.Columns.Item(11).Visible = False
            oGrid.Columns.Item(14).Visible = False
            oGrid.Columns.Item(13).TitleObject.Caption = "Amount in System Currency (" & systemcurrency & ")"
            oEditTextColumn = oGrid.Columns.Item(13)
            oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            oGrid.Columns.Item(16).TitleObject.Caption = "Remarks"
            oGrid.Columns.Item(17).TitleObject.Caption = "Select"
            oGrid.Columns.Item(18).TitleObject.Caption = "Approval Status"
            oGrid.Columns.Item(15).TitleObject.Sortable = True
            '   oGrid.Columns.Item(14).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item("U_Z_DisRule").TitleObject.Caption = "Distribution Rule"
            oGrid.Columns.Item("U_Z_DisRule").Editable = False
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

            oGrid.Columns.Item("U_Z_DisRule").TitleObject.Caption = "Distribution Rule"
            oGrid.Columns.Item("U_Z_DisRule").Editable = False

            oGrid.Columns.Item(13).TitleObject.Caption = "Remarks"
            oGrid.Columns.Item(13).Editable = False
            'oGrid.Columns.Item(13).TitleObject.Caption = "Approval Status"
            'oGrid.Columns.Item(13).TitleObject.Sortable = True
            'oGrid.Columns.Item(13).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(14).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item(14).Editable = True
            oGrid.Columns.Item(15).TitleObject.Caption = "Code"
            oGrid.Columns.Item(15).Editable = False
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
            oGrid.Columns.Item("U_Z_DisRule").TitleObject.Caption = "Distribution Rule"
            oGrid.Columns.Item("U_Z_DisRule").Editable = False
            oGrid.Columns.Item(10).TitleObject.Caption = "Ref1"
            oGrid.Columns.Item(10).Editable = False
            oGrid.Columns.Item(11).TitleObject.Caption = "Remakrs"
            oGrid.Columns.Item(11).Editable = False
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
            oGrid.Columns.Item(5).TitleObject.Caption = "Employee Name / Contractor Name"

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
            oGrid.Columns.Item("U_Z_CNTID").TitleObject.Caption = "Contract ID"
            oGrid.Columns.Item("U_Z_CNTID").Editable = False
            oGrid.Columns.Item("Emp").TitleObject.Caption = "Employee ID"
            oGrid.Columns.Item("Emp").Editable = False


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
            oGrid.Columns.Item("EmpID").Visible = True
            oGrid.Columns.Item("Position").TitleObject.Caption = "Employee Name"
            oGrid.Columns.Item("Position").Visible = True



            oGrid.Columns.Item("FromDate").TitleObject.Caption = "Start Date"
            oGrid.Columns.Item("FromDate").Visible = False
            oGrid.Columns.Item("EndDate").TitleObject.Caption = "End Date"
            oGrid.Columns.Item("EndDate").Visible = False

            oGrid.Columns.Item("EstimatedHours").TitleObject.Caption = "Estimated Hours "
            oGrid.Columns.Item("EstimatedHours").TitleObject.Sortable = True
            oGrid.Columns.Item("EstimatedHours").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)


            oGrid.Columns.Item("ActualHours").TitleObject.Caption = "Actual Hours "
            oGrid.Columns.Item("ActualHours").TitleObject.Sortable = True
            oGrid.Columns.Item("ActualHours").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item("ActualHours").Visible = True
            oGrid.Columns.Item("VarianceHours").TitleObject.Caption = "Variance Hours "
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
            oGrid.Columns.Item("MatrialOrdered").TitleObject.Caption = "Actual Cost"
            oGrid.Columns.Item("MatrialOrdered").TitleObject.Sortable = True
            oGrid.Columns.Item("MatrialOrdered").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)



            oGrid.Columns.Item("MaterialReceived").TitleObject.Caption = "Variance In Cost"
            oGrid.Columns.Item("MaterialReceived").TitleObject.Sortable = True
            oGrid.Columns.Item("MaterialReceived").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oGrid.Columns.Item("OrderPending").TitleObject.Caption = "Order Pending"
            oGrid.Columns.Item("OrderPending").Visible = False
            ' oGrid.Columns.Item("OrderPending").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oGrid.Columns.Item("MaterialPending").TitleObject.Caption = "Material Pending"
            oGrid.Columns.Item("MaterialPending").Visible = False
            ' oGrid.Columns.Item("MaterialPending").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oGrid.Columns.Item("Cost Variance").TitleObject.Caption = "Cost Variance"
            oGrid.Columns.Item("Cost Variance").Visible = False
            '  oGrid.Columns.Item("Cost Variance").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

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
            oGrid.Columns.Item(7).TitleObject.Caption = "Type"
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

            oGrid.Columns.Item(10).TitleObject.Caption = "Days "
            oEditTextColumn = oGrid.Columns.Item(10)

            oEditTextColumn = oGrid.Columns.Item(10)
            oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oGrid.Columns.Item(11).TitleObject.Caption = " Hours"
            oEditTextColumn = oGrid.Columns.Item(11)

            oEditTextColumn = oGrid.Columns.Item(11)
            oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            ' oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oGrid.Columns.Item(12).TitleObject.Caption = "Reason / Remarks"
            oGrid.Columns.Item(13).TitleObject.Caption = "Status"
            oGrid.Columns.Item(13).TitleObject.Sortable = True
            oGrid.Columns.Item(13).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

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

            oGrid.Columns.Item("U_Z_TYPE").TitleObject.Caption = "Type"
            oGrid.Columns.Item("U_Z_BDGQTY").TitleObject.Caption = "Budgeted  Hours"
            oGrid.Columns.Item("U_Z_BDGQTY").Visible = False
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
            oGrid.Columns.Item("STATUS").TitleObject.Caption = "Status"
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
                strSQL = "SELECT  T0.[U_Z_EMPCODE], T0.[U_Z_DOCDATE],T1.[U_Z_DisRule],count(*)  FROM [dbo].[@Z_OEXP]  T0 inner join  [dbo].[@Z_EXP1]  T1  on T0.Code=T1.U_Z_REFCODE inner join [@Z_Expances] T3 on T3.U_Z_ExpName=T1.[U_Z_ExpName] inner join OACT T4 on T4.FormatCode=T3.[U_Z_ActCode]"
                strSQL = strSQL & " where T1.[U_Z_APPROVED]='A' and isnull(T1.[U_Z_DocNum],'')='' and  T1.[U_Z_Flag]='Y' and " & strCondition & " group by T0.[U_Z_EMPCODE], T0.[U_Z_DOCDATE],T1.[U_Z_DisRule]"
                oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oTempRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTempRs.DoQuery(strSQL)
                If oApplication.Company.InTransaction Then
                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                End If
                oApplication.Company.StartTransaction()
                Dim strProjectcode, strDisRule As String

                For intRow As Integer = 0 To oTempRs.RecordCount - 1
                    oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    strEMpid = oTempRs.Fields.Item(0).Value
                    dtDate = oTempRs.Fields.Item(1).Value
                    strDisRule = oTempRs.Fields.Item(2).Value
                    strsql1 = "SELECT  T1.[Code],T0.[U_Z_EMPCODE], T0.[U_Z_EMPNAME], convert(varchar(10),T0.[U_Z_DOCDATE],105),T1.[U_Z_EXPNAME], T1.[U_Z_PRJCODE],T1.[U_Z_PRJNAME],T3.[U_Z_ActCode],T4.[AcctName],T1.[U_Z_LocAmount],T1.[U_Z_Ref1],T1.[U_Z_Remarks],T1.U_Z_MODNAME 'Mod',T1.U_Z_ACTNAME 'Act' FROM [dbo].[@Z_OEXP]  T0 inner join  [dbo].[@Z_EXP1]  T1  on T0.Code=T1.U_Z_REFCODE inner join [@Z_Expances] T3 on T3.U_Z_ExpName=T1.[U_Z_ExpName] inner join OACT T4 on T4.FormatCode=T3.[U_Z_ActCode]"
                    strsql1 = strsql1 & " where T0.[U_Z_EMPCODE]='" & strEMpid & "' and isnull(T1.U_Z_DisRule,'')='" & strDisRule & "' and  U_Z_DocDate='" & dtDate.ToString("yyyy-MM-dd") & "' and  T1.[U_Z_APPROVED]='A' and isnull(T1.[U_Z_DocNum],'')='' and  T1.[U_Z_Flag]='Y' and " & strCondition & " Order by T0.U_Z_DocDate, T0.[U_Z_EMPCODE],T1.[U_Z_EXPNAME], T1.[U_Z_PRJCODE] "
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
                        If strDisRule <> "" Then
                            Dim stString As String()
                            stString = strDisRule.Split(";")
                            Try
                                oJE.Lines.CostingCode = stString(0)
                            Catch ex As Exception

                            End Try
                            Try
                                oJE.Lines.CostingCode2 = stString(1)
                            Catch ex As Exception

                            End Try
                            Try
                                oJE.Lines.CostingCode3 = stString(2)
                            Catch ex As Exception

                            End Try
                            Try
                                oJE.Lines.CostingCode4 = stString(3)
                            Catch ex As Exception

                            End Try
                            Try
                                oJE.Lines.CostingCode5 = stString(4)
                            Catch ex As Exception

                            End Try

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

                        If strDisRule <> "" Then
                            Dim stString As String()
                            stString = strDisRule.Split(";")
                            Try
                                oJE.Lines.CostingCode = stString(0)
                            Catch ex As Exception

                            End Try
                            Try
                                oJE.Lines.CostingCode2 = stString(1)
                            Catch ex As Exception

                            End Try
                            Try
                                oJE.Lines.CostingCode3 = stString(2)
                            Catch ex As Exception

                            End Try
                            Try
                                oJE.Lines.CostingCode4 = stString(3)
                            Catch ex As Exception

                            End Try
                            Try
                                oJE.Lines.CostingCode5 = stString(4)
                            Catch ex As Exception

                            End Try

                        End If
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
                Case mnu_report
                    If pVal.BeforeAction = False Then
                        'LoadForm()
                        Dim oTe As New clsLogin
                        oTe.LoadForm("Reports")
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
            If pVal.FormTypeEx = frm_Report Then
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
                                    otempRs.DoQuery("Update  [@Z_EXP1] set U_Z_Flag='N' where U_Z_Flag='Y'")
                              
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                '    oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oForm.Freeze(True)

                                Select Case pVal.ItemUID

                                    Case "21"
                                        databind(oForm)
                                    Case "31"
                                        oForm.PaneLevel = 3
                                    Case "32"
                                        oForm.PaneLevel = 4
                                    Case "30"
                                        oGrid = oForm.Items.Item("7").Specific
                                        For introw As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                            oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            oCheckBoxColumn = oGrid.Columns.Item("Select")
                                            oCheckBoxColumn.Check(introw, True)
                                        Next
                                        oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
                                    Case "34"
                                        oGrid = oForm.Items.Item("7").Specific
                                        For introw As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                            oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            oCheckBoxColumn = oGrid.Columns.Item("Select")
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
                                                oCheckBoxColumn = oGrid.Columns.Item("Select")
                                                If oCheckBoxColumn.IsChecked(introw) Then
                                                    strAmount = oGrid.DataTable.GetValue(10, introw)
                                                    Dim dblAmount As Double
                                                    dblAmount = (oApplication.Utilities.getDocumentQuantity(strAmount))
                                                    st = "Update  [@Z_EXP1] set U_Z_Flag='Y',U_Z_LocAmount='" & dblAmount.ToString.Replace(",", ".") & "' where Code='" & oGrid.DataTable.GetValue("Code", introw) & "'"
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
                                        If pVal.ItemUID = "12" Or pVal.ItemUID = "9" Then
                                            val = oDataTable.GetValue("empID", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        ElseIf pVal.ItemUID = "102" Then
                                            val = oDataTable.GetValue("FormatCode", 0)
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
