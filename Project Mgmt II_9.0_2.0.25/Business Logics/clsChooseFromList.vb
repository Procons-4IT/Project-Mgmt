Public Class clsChooseFromList
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
    Private oedit As SAPbouiCOM.EditText
    '   Private oForm As SAPbouiCOM.Form
    Private objSoureceForm As SAPbouiCOM.Form
    Private objform As SAPbouiCOM.Form
    Private oMatrix As SAPbouiCOM.Grid
    Private osourcegrid As SAPbouiCOM.Matrix
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
            objform.Freeze(True)
            objform.DataSources.DataTables.Add("dtLevel3")
            Ouserdatasource = objform.DataSources.UserDataSources.Add("dbFind", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 250)
            oedit = objform.Items.Item("etFind").Specific
            oedit.DataBind.SetBound(True, "", "dbFind")
            objGrid = objform.Items.Item("mtchoose").Specific
            dtTemp = objform.DataSources.DataTables.Item("dtLevel3")
            Dim strEMPIDS As String = oApplication.Utilities.getManagerEmPID(oApplication.Company.UserName)
            If BinDescrUID <> "" Then
                strEMPIDS = BinDescrUID
            End If

            If choice = "MODULE" Then
                objform.Title = "Phase - Selection"
                If ItemCode <> "" And ItemCode <> "" Then
                    If Documentchoice = "" Then
                        Dim stModQury As String = "Select U_Z_MODNAME from [@Z_MODULE] where U_Z_STATUS='N'"
                        ' strSQL = "Select U_Z_MODNAME,U_Z_HOURS from " & CFLChoice & " where DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')"
                        '  strSQL = "Select U_Z_MODNAME,sum(U_Z_HOURS) from " & CFLChoice & " where  U_Z_MODNAME In (" & stModQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')  and U_Z_EMPID  in (" & strEMPIDS & ") group by U_Z_MODNAME"
                        strSQL = "Select U_Z_MODNAME,sum(U_Z_HOURS) from " & CFLChoice & " where  U_Z_MODNAME In (" & stModQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')   group by U_Z_MODNAME"
                        dtTemp.ExecuteQuery(strSQL)
                        objGrid.DataTable = dtTemp
                    ElseIf Documentchoice = "Document" Then
                        Dim stModQury1 As String = "Select U_Z_MODNAME from [@Z_MODULE] where U_Z_STATUS='N'"
                        strSQL = "Select U_Z_MODNAME,sum(U_Z_HOURS) from " & CFLChoice & " where  U_Z_MODNAME In (" & stModQury1 & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')   group by U_Z_MODNAME"
                        dtTemp.ExecuteQuery(strSQL)
                        objGrid.DataTable = dtTemp
                    End If
                    'strSQL = "Select U_BinCode,U_Descrip,U_OnHand from [@DABT_OITBL] where U_ItemCode='" & ItemCode & "'and  U_WhsCode='" & CFLChoice & "' Order by U_BinCode "
                    objGrid.Columns.Item(0).TitleObject.Caption = "Phase Name"
                    objGrid.Columns.Item(1).TitleObject.Caption = "Hours"
                End If
            ElseIf choice = "ACTIVITY" Or choice = "ACTIVITY1" Or choice = "ACTIVITY2" Or choice = "ACTIVITY3" Then
                objform.Title = "Activity - Selection"
                If ItemCode <> "" And ItemCode <> "" Then
                    Dim stActQury As String = "Select U_Z_ACTNAME from [@Z_ACTIVITY] where U_Z_STATUS='N' and U_Z_MODNAME IN (Select U_Z_MODNAME From [@Z_MODULE] where U_Z_STATUS='N')"
                    Dim stModQury As String = "Select U_Z_MODNAME from [@Z_MODULE] where U_Z_STATUS='N'"
                    If choice = "ACTIVITY1" Or choice = "ACTIVITY3" Then
                        If Documentchoice = "" Then
                            If choice = "ACTIVITY3" Then
                                If ContractID <> "" Then
                                    '  strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours ,U_Z_CNTID,U_Z_POSITION,U_Z_BOQ,U_Z_CUSTCNTID,U_Z_EMPID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')  and U_Z_CNTID  in (" & ContractID & ") and U_Z_TYPE='E' "
                                    strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours ,U_Z_CNTID,U_Z_POSITION,U_Z_BOQ,U_Z_CUSTCNTID,U_Z_EMPID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')   and U_Z_TYPE='E' "
                                Else
                                    '  strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours ,U_Z_CNTID,U_Z_POSITION,U_Z_BOQ,U_Z_CUSTCNTID,U_Z_EMPID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')  and U_Z_EMPID  in (" & strEMPIDS & ") and U_Z_TYPE='E' "
                                    strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours ,U_Z_CNTID,U_Z_POSITION,U_Z_BOQ,U_Z_CUSTCNTID,U_Z_EMPID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')  and  U_Z_TYPE='E' "

                                End If
                            Else
                                '   strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours ,U_Z_CNTID,U_Z_POSITION,U_Z_BOQ,U_Z_CUSTCNTID,U_Z_EMPID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')  and U_Z_EMPID  in (" & strEMPIDS & ")"
                                strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours ,U_Z_CNTID,U_Z_POSITION,U_Z_BOQ,U_Z_CUSTCNTID,U_Z_EMPID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')" '  and U_Z_EMPID  in (" & strEMPIDS & ")"

                            End If
                            dtTemp.ExecuteQuery(strSQL)
                            objGrid.DataTable = dtTemp
                        ElseIf Documentchoice = "Document" Then
                            strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours ,U_Z_CNTID,U_Z_POSITION,U_Z_BOQ,U_Z_CUSTCNTID,U_Z_EMPID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')" '  and U_Z_EMPID  in (" & strEMPIDS & ")"
                            dtTemp.ExecuteQuery(strSQL)
                            objGrid.DataTable = dtTemp
                        Else
                            'strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Amount from " & CFLChoice & " where DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')and U_Z_MODNAME='" & Documentchoice & "'"
                            '  strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours,U_Z_CNTID,U_Z_POSITION ,U_Z_BOQ,U_Z_CUSTCNTID,U_Z_EMPID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')and U_Z_MODNAME='" & Documentchoice & "'  and U_Z_EMPID  in (" & strEMPIDS & ") "
                            strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours,U_Z_CNTID,U_Z_POSITION ,U_Z_BOQ,U_Z_CUSTCNTID,U_Z_EMPID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')and U_Z_MODNAME='" & Documentchoice & "'" '  and U_Z_EMPID  in (" & strEMPIDS & ") "
                            dtTemp.ExecuteQuery(strSQL)
                            objGrid.DataTable = dtTemp
                        End If
                    Else

                        If Documentchoice = "" Then
                            ' strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours ,U_Z_CNTID,U_Z_POSITION,U_Z_BOQ,U_Z_CUSTCNTID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')  and U_Z_EMPID  in (" & strEMPIDS & ")"
                            strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours ,U_Z_CNTID,U_Z_POSITION,U_Z_BOQ,U_Z_CUSTCNTID,U_Z_EMPID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')" '  and U_Z_EMPID  in (" & strEMPIDS & ")"
                            dtTemp.ExecuteQuery(strSQL)
                            objGrid.DataTable = dtTemp
                        Else
                            'strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Amount from " & CFLChoice & " where DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')and U_Z_MODNAME='" & Documentchoice & "'"
                            '  strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours,U_Z_CNTID,U_Z_POSITION ,U_Z_BOQ,U_Z_CUSTCNTID from " & CFLChoice & " where (U_Z_STATUS='I') and  U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')and U_Z_MODNAME='" & Documentchoice & "'  and U_Z_EMPID  in (" & strEMPIDS & ") "
                            strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours,U_Z_CNTID,U_Z_POSITION ,U_Z_BOQ,U_Z_CUSTCNTID,U_Z_EMPID from " & CFLChoice & " where (U_Z_STATUS='I') and  U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')and U_Z_MODNAME='" & Documentchoice & "'" '  and U_Z_EMPID  in (" & strEMPIDS & ") "
                            dtTemp.ExecuteQuery(strSQL)
                            objGrid.DataTable = dtTemp
                        End If
                    End If
                    'strSQL = "Select U_BinCode,U_Descrip,U_OnHand from [@DABT_OITBL] where U_ItemCode='" & ItemCode & "'and  U_WhsCode='" & CFLChoice & "' Order by U_BinCode "
                    objGrid.Columns.Item(0).TitleObject.Caption = "Activity Name"
                    objGrid.Columns.Item(1).TitleObject.Caption = "Phase Name"
                    'objGrid.Columns.Item(2).TitleObject.Caption = "Eastimate Amount"
                    objGrid.Columns.Item(2).TitleObject.Caption = "Eastimate Hours"
                    objGrid.Columns.Item("U_Z_CNTID").TitleObject.Caption = "Sub.Contract Number"
                    objGrid.Columns.Item("U_Z_POSITION").TitleObject.Caption = "Emp Name / Contractor Name"
                    objGrid.Columns.Item("U_Z_BOQ").TitleObject.Caption = "BoQ Reference"
                    objGrid.Columns.Item("U_Z_EMPID").TitleObject.Caption = "Employee ID"
                End If
            Else
                objform.Title = "Activity - Selection"
                strSQL = "Select U_Z_ACTNAME,U_Z_ACTNAME from " & CFLChoice & " where DocEntry= (Select DocEntry from [@Z_HPRJ] where  U_Z_PRJCODE='" & ItemCode & "') and U_Z_MODNAME='" & Documentchoice & "'"
                dtTemp.ExecuteQuery(strSQL)
                objGrid.DataTable = dtTemp
                objGrid.Columns.Item(0).TitleObject.Caption = "Activity Name"
                objGrid.Columns.Item(1).TitleObject.Caption = "Activity Name"
            End If
            objGrid.AutoResizeColumns()
            objGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            If objGrid.Rows.Count > 0 Then
                objGrid.Rows.SelectedRows.Add(0)
            End If
            objform.Freeze(False)
            objform.Update()
            sSearchList = " "
            Dim i As Integer = 0
            While i <= objGrid.DataTable.Rows.Count - 1
                sSearchList += Convert.ToString(objGrid.DataTable.GetValue(0, i)) + SEPRATOR + i.ToString + " "
                System.Math.Min(System.Threading.Interlocked.Increment(i), i - 1)
            End While
        Catch ex As Exception
            oApplication.SBO_Application.MessageBox(ex.Message)
            oApplication.SBO_Application.MessageBox(ex.Message)
        Finally
        End Try
    End Sub
    'Public Sub databound(ByVal objform As SAPbouiCOM.Form)
    '    Try
    '        Dim strSQL As String = ""
    '        Dim ObjSegRecSet As SAPbobsCOM.Recordset
    '        ObjSegRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        objform.Freeze(True)
    '        objform.DataSources.DataTables.Add("dtLevel3")
    '        Ouserdatasource = objform.DataSources.UserDataSources.Add("dbFind", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 250)
    '        oedit = objform.Items.Item("etFind").Specific
    '        oedit.DataBind.SetBound(True, "", "dbFind")
    '        objGrid = objform.Items.Item("mtchoose").Specific
    '        dtTemp = objform.DataSources.DataTables.Item("dtLevel3")
    '        Dim strEMPIDS As String = oApplication.Utilities.getManagerEmPID(oApplication.Company.UserName)
    '        If BinDescrUID <> "" Then
    '            strEMPIDS = BinDescrUID
    '        End If

    '        If choice = "MODULE" Then
    '            objform.Title = "Phase - Selection"
    '            If ItemCode <> "" And ItemCode <> "" Then
    '                If Documentchoice = "" Then
    '                    Dim stModQury As String = "Select U_Z_MODNAME from [@Z_MODULE] where U_Z_STATUS='N'"
    '                    ' strSQL = "Select U_Z_MODNAME,U_Z_HOURS from " & CFLChoice & " where DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')"
    '                    strSQL = "Select U_Z_MODNAME,sum(U_Z_HOURS) from " & CFLChoice & " where  U_Z_MODNAME In (" & stModQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')  and U_Z_EMPID  in (" & strEMPIDS & ") group by U_Z_MODNAME"
    '                    dtTemp.ExecuteQuery(strSQL)
    '                    objGrid.DataTable = dtTemp
    '                ElseIf Documentchoice = "Document" Then
    '                    Dim stModQury1 As String = "Select U_Z_MODNAME from [@Z_MODULE] where U_Z_STATUS='N'"
    '                    strSQL = "Select U_Z_MODNAME,sum(U_Z_HOURS) from " & CFLChoice & " where  U_Z_MODNAME In (" & stModQury1 & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')   group by U_Z_MODNAME"
    '                    dtTemp.ExecuteQuery(strSQL)
    '                    objGrid.DataTable = dtTemp
    '                End If
    '                'strSQL = "Select U_BinCode,U_Descrip,U_OnHand from [@DABT_OITBL] where U_ItemCode='" & ItemCode & "'and  U_WhsCode='" & CFLChoice & "' Order by U_BinCode "
    '                objGrid.Columns.Item(0).TitleObject.Caption = "Phase Name"
    '                objGrid.Columns.Item(1).TitleObject.Caption = "Hours"
    '            End If
    '        ElseIf choice = "ACTIVITY" Or choice = "ACTIVITY1" Or choice = "ACTIVITY2" Or choice = "ACTIVITY3" Then
    '            objform.Title = "Activity - Selection"
    '            If ItemCode <> "" And ItemCode <> "" Then
    '                Dim stActQury As String = "Select U_Z_ACTNAME from [@Z_ACTIVITY] where U_Z_STATUS='N' and U_Z_MODNAME IN (Select U_Z_MODNAME From [@Z_MODULE] where U_Z_STATUS='N')"
    '                Dim stModQury As String = "Select U_Z_MODNAME from [@Z_MODULE] where U_Z_STATUS='N'"
    '                If choice = "ACTIVITY1" Or choice = "ACTIVITY3" Then
    '                    If Documentchoice = "" Then
    '                        If choice = "ACTIVITY3" Then
    '                            If ContractID <> "" Then
    '                                strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours ,U_Z_CNTID,U_Z_POSITION,U_Z_BOQ,U_Z_CUSTCNTID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')  and U_Z_CNTID  in (" & ContractID & ") and U_Z_TYPE='E' "
    '                            Else
    '                                strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours ,U_Z_CNTID,U_Z_POSITION,U_Z_BOQ,U_Z_CUSTCNTID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')  and U_Z_EMPID  in (" & strEMPIDS & ") and U_Z_TYPE='E' "

    '                            End If
    '                        Else
    '                            strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours ,U_Z_CNTID,U_Z_POSITION,U_Z_BOQ,U_Z_CUSTCNTID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')  and U_Z_EMPID  in (" & strEMPIDS & ")"

    '                        End If
    '                         dtTemp.ExecuteQuery(strSQL)
    '                        objGrid.DataTable = dtTemp
    '                    ElseIf Documentchoice = "Document" Then
    '                        strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours ,U_Z_CNTID,U_Z_POSITION,U_Z_BOQ,U_Z_CUSTCNTID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')" '  and U_Z_EMPID  in (" & strEMPIDS & ")"
    '                        dtTemp.ExecuteQuery(strSQL)
    '                        objGrid.DataTable = dtTemp
    '                    Else
    '                        'strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Amount from " & CFLChoice & " where DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')and U_Z_MODNAME='" & Documentchoice & "'"
    '                        strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours,U_Z_CNTID,U_Z_POSITION ,U_Z_BOQ,U_Z_CUSTCNTID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')and U_Z_MODNAME='" & Documentchoice & "'  and U_Z_EMPID  in (" & strEMPIDS & ") "
    '                        dtTemp.ExecuteQuery(strSQL)
    '                        objGrid.DataTable = dtTemp
    '                    End If
    '                Else

    '                    If Documentchoice = "" Then
    '                        strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours ,U_Z_CNTID,U_Z_POSITION,U_Z_BOQ,U_Z_CUSTCNTID from " & CFLChoice & " where U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')  and U_Z_EMPID  in (" & strEMPIDS & ")"
    '                        dtTemp.ExecuteQuery(strSQL)
    '                        objGrid.DataTable = dtTemp
    '                    Else
    '                        'strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Amount from " & CFLChoice & " where DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')and U_Z_MODNAME='" & Documentchoice & "'"
    '                        strSQL = "Select U_Z_ActName,U_Z_MODNAME,U_Z_Hours,U_Z_CNTID,U_Z_POSITION ,U_Z_BOQ,U_Z_CUSTCNTID from " & CFLChoice & " where (U_Z_STATUS='I') and  U_Z_MODNAME In (" & stModQury & ") and  U_Z_ACTNAME IN (" & stActQury & ") and  DocEntry= (Select DocEntry from [@Z_HPRJ] where U_Z_PRJCODE='" & ItemCode & "')and U_Z_MODNAME='" & Documentchoice & "'  and U_Z_EMPID  in (" & strEMPIDS & ") "
    '                        dtTemp.ExecuteQuery(strSQL)
    '                        objGrid.DataTable = dtTemp
    '                    End If
    '                End If
    '                'strSQL = "Select U_BinCode,U_Descrip,U_OnHand from [@DABT_OITBL] where U_ItemCode='" & ItemCode & "'and  U_WhsCode='" & CFLChoice & "' Order by U_BinCode "
    '                objGrid.Columns.Item(0).TitleObject.Caption = "Activity Name"
    '                objGrid.Columns.Item(1).TitleObject.Caption = "Phase Name"
    '                'objGrid.Columns.Item(2).TitleObject.Caption = "Eastimate Amount"
    '                objGrid.Columns.Item(2).TitleObject.Caption = "Eastimate Hours"
    '                objGrid.Columns.Item("U_Z_CNTID").TitleObject.Caption = "Sub.Contract Number"
    '                objGrid.Columns.Item("U_Z_POSITION").TitleObject.Caption = "Emp Name / Contractor Name"
    '                objGrid.Columns.Item("U_Z_BOQ").TitleObject.Caption = "BoQ Reference"
    '            End If
    '        Else
    '            objform.Title = "Activity - Selection"
    '            strSQL = "Select U_Z_ACTNAME,U_Z_ACTNAME from " & CFLChoice & " where DocEntry= (Select DocEntry from [@Z_HPRJ] where  U_Z_PRJCODE='" & ItemCode & "') and U_Z_MODNAME='" & Documentchoice & "'"
    '            dtTemp.ExecuteQuery(strSQL)
    '            objGrid.DataTable = dtTemp
    '            objGrid.Columns.Item(0).TitleObject.Caption = "Activity Name"
    '            objGrid.Columns.Item(1).TitleObject.Caption = "Activity Name"
    '        End If
    '        objGrid.AutoResizeColumns()
    '        objGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    '        If objGrid.Rows.Count > 0 Then
    '            objGrid.Rows.SelectedRows.Add(0)
    '        End If
    '        objform.Freeze(False)
    '        objform.Update()
    '        sSearchList = " "
    '        Dim i As Integer = 0
    '        While i <= objGrid.DataTable.Rows.Count - 1
    '            sSearchList += Convert.ToString(objGrid.DataTable.GetValue(0, i)) + SEPRATOR + i.ToString + " "
    '            System.Math.Min(System.Threading.Interlocked.Increment(i), i - 1)
    '        End While
    '    Catch ex As Exception
    '        oApplication.SBO_Application.MessageBox(ex.Message)
    '        oApplication.SBO_Application.MessageBox(ex.Message)
    '    Finally
    '    End Try
    'End Sub
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

#Region "FormDataEvent"


#End Region

#Region "Class Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

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

#Region "Item Event"

    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        BubbleEvent = True
        If pVal.FormTypeEx = frm_ChoosefromList Then
            If pVal.Before_Action = True Then
                If pVal.ItemUID = "mtchoose" Then
                    Try
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row <> -1 Then
                            oForm = GetForm(pVal.FormUID)
                            oItem = CType(oForm.Items.Item(pVal.ItemUID), SAPbouiCOM.Item)
                            oMatrix = CType(oItem.Specific, SAPbouiCOM.Grid)
                            oMatrix.Rows.SelectedRows.Add(pVal.Row)
                        End If
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK And pVal.Row <> -1 Then
                            oForm = GetForm(pVal.FormUID)
                            oItem = oForm.Items.Item("mtchoose")
                            oMatrix = CType(oItem.Specific, SAPbouiCOM.Grid)
                            Dim inti As Integer
                            inti = 0
                            inti = 0
                            While inti <= oMatrix.DataTable.Rows.Count - 1
                                If oMatrix.Rows.IsSelected(inti) = True Then
                                    intRowId = inti
                                End If
                                System.Math.Min(System.Threading.Interlocked.Increment(inti), inti - 1)
                            End While
                            If CFLChoice <> "" Then
                                If intRowId = 0 Then
                                    strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                                    strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(1, intRowId))
                                Else
                                    strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                                    strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(1, intRowId))
                                End If
                                If choice = "ACTIVITY1" Then
                                    strContractID = Convert.ToString(oMatrix.DataTable.GetValue("U_Z_CNTID", intRowId))
                                    strContractName = Convert.ToString(oMatrix.DataTable.GetValue("U_Z_POSITION", intRowId))
                                    strCustCntID = Convert.ToString(oMatrix.DataTable.GetValue("U_Z_CUSTCNTID", intRowId))
                                    If strContractID = "" Then
                                        strContractName = ""
                                    End If
                                End If
                                oForm.Close()
                                oForm = GetForm(SourceFormUID)
                                If choice <> "" Then
                                    If choice = "ACTIVITY2" Then
                                        oApplication.Utilities.setEdittextvalue(oForm, "9", strSelectedItem2)
                                        Try
                                            oApplication.Utilities.setEdittextvalue(oForm, "11", strSelectedItem1)

                                        Catch ex As Exception

                                        End Try
                                    ElseIf choice = "ACTIVITY3" Then
                                        osourcegrid = oForm.Items.Item(ItemUID).Specific
                                        oApplication.Utilities.SetMatrixValues(osourcegrid, "V_21", sourcerowId, strSelectedItem1)
                                        oApplication.Utilities.SetMatrixValues(osourcegrid, "V_20", sourcerowId, strSelectedItem2)
                                    Else
                                        osourcegrid = oForm.Items.Item(ItemUID).Specific
                                        If choice = "ACTIVITY1" Then
                                            Try
                                                oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_ACTNAME", sourcerowId, strSelectedItem1)
                                                oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_MDNAME", sourcerowId, strSelectedItem2)
                                                oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_CNTID", sourcerowId, strContractID)
                                                oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_CNTNAME", sourcerowId, strContractName)
                                                oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_CUSTCNTID", sourcerowId, strCustCntID)
                                                oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_BOQREF", sourcerowId, getBOQReference(sourceItemCode, ItemCode, strSelectedItem2, strSelectedItem1))
                                            Catch ex As Exception
                                                oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_BOQREF", sourcerowId, getBOQReference(sourceItemCode, ItemCode, strSelectedItem2, strSelectedItem1))
                                            End Try
                                        Else
                                            oApplication.Utilities.SetMatrixValues(osourcegrid, sourceColumID, sourcerowId, strSelectedItem1)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        oApplication.SBO_Application.MessageBox(ex.Message)
                    End Try
                End If
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Then
                    Try
                        If pVal.ItemUID = "mtchoose" Then
                            oForm = GetForm(pVal.FormUID)
                            oItem = CType(oForm.Items.Item("mtchoose"), SAPbouiCOM.Item)
                            oMatrix = CType(oItem.Specific, SAPbouiCOM.Grid)
                            intRowId = pVal.Row - 1
                        End If
                        Dim inti As Integer
                        If pVal.CharPressed = 13 Then
                            inti = 0
                            inti = 0
                            oForm = GetForm(pVal.FormUID)
                            oItem = CType(oForm.Items.Item("mtchoose"), SAPbouiCOM.Item)

                            oMatrix = CType(oItem.Specific, SAPbouiCOM.Grid)
                            While inti <= oMatrix.DataTable.Rows.Count - 1
                                If oMatrix.Rows.IsSelected(inti) = True Then
                                    intRowId = inti
                                End If
                                System.Math.Min(System.Threading.Interlocked.Increment(inti), inti - 1)
                            End While
                            If CFLChoice <> "" Then
                                If intRowId = 0 Then
                                    strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                                    strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(1, intRowId))
                                Else
                                    strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                                    strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(1, intRowId))
                                End If
                                If choice = "ACTIVITY1" Then
                                    strContractID = Convert.ToString(oMatrix.DataTable.GetValue("U_Z_CNTID", intRowId))
                                    strContractName = Convert.ToString(oMatrix.DataTable.GetValue("U_Z_POSITION", intRowId))
                                    strCustCntID = Convert.ToString(oMatrix.DataTable.GetValue("U_Z_CUSTCNTID", intRowId))
                                    If strContractID = "" Then
                                        strContractName = ""
                                    End If
                                End If
                                oForm.Close()
                                oForm = GetForm(SourceFormUID)
                                If choice <> "" Then
                                    If choice = "ACTIVITY2" Then
                                        oApplication.Utilities.setEdittextvalue(oForm, "9", strSelectedItem2)
                                        Try
                                            oApplication.Utilities.setEdittextvalue(oForm, "11", strSelectedItem1)
                                        Catch ex As Exception
                                        End Try
                                    ElseIf choice = "ACTIVITY3" Then
                                        osourcegrid = oForm.Items.Item(ItemUID).Specific
                                        oApplication.Utilities.SetMatrixValues(osourcegrid, "V_21", sourcerowId, strSelectedItem1)
                                        oApplication.Utilities.SetMatrixValues(osourcegrid, "V_20", sourcerowId, strSelectedItem2)
                                    Else
                                        osourcegrid = oForm.Items.Item(ItemUID).Specific
                                        If choice = "ACTIVITY1" Then
                                            Try
                                                oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_ACTNAME", sourcerowId, strSelectedItem1)
                                                oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_MDNAME", sourcerowId, strSelectedItem2)
                                                oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_CNTID", sourcerowId, strContractID)
                                                oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_CNTNAME", sourcerowId, strContractName)
                                                oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_CUSTCNTID", sourcerowId, strCustCntID)
                                                oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_BOQREF", sourcerowId, getBOQReference(sourceItemCode, ItemCode, strSelectedItem2, strSelectedItem1))
                                            Catch ex As Exception
                                                oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_BOQREF", sourcerowId, getBOQReference(sourceItemCode, ItemCode, strSelectedItem2, strSelectedItem1))
                                            End Try
                                        Else
                                            oApplication.Utilities.SetMatrixValues(osourcegrid, sourceColumID, sourcerowId, strSelectedItem1)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        oApplication.SBO_Application.MessageBox(ex.Message)
                    End Try
                End If


                If pVal.ItemUID = "btnChoose" AndAlso pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    oForm = GetForm(pVal.FormUID)
                    oItem = oForm.Items.Item("mtchoose")
                    oMatrix = CType(oItem.Specific, SAPbouiCOM.Grid)
                    Dim inti As Integer
                    inti = 0
                    inti = 0
                    While inti <= oMatrix.DataTable.Rows.Count - 1
                        If oMatrix.Rows.IsSelected(inti) = True Then
                            intRowId = inti
                        End If
                        System.Math.Min(System.Threading.Interlocked.Increment(inti), inti - 1)
                    End While
                    If CFLChoice <> "" Then
                        If intRowId = 0 Then
                            strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                            strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(1, intRowId))
                        Else
                            strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                            strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(1, intRowId))
                        End If
                        If choice = "ACTIVITY1" Then
                            strContractID = Convert.ToString(oMatrix.DataTable.GetValue("U_Z_CNTID", intRowId))
                            strContractName = Convert.ToString(oMatrix.DataTable.GetValue("U_Z_POSITION", intRowId))
                            strCustCntID = Convert.ToString(oMatrix.DataTable.GetValue("U_Z_CUSTCNTID", intRowId))
                            If strContractID = "" Then
                                strContractName = ""
                            End If
                        End If
                        oForm.Close()
                        oForm = GetForm(SourceFormUID)
                        If choice <> "" Then
                            osourcegrid = oForm.Items.Item(ItemUID).Specific
                            ' osourcegrid.DataTable.SetValue(sourceColumID, sourcerowId, strSelectedItem1)
                            If choice = "ACTIVITY2" Then
                                oApplication.Utilities.setEdittextvalue(oForm, "9", strSelectedItem2)
                                Try
                                    oApplication.Utilities.setEdittextvalue(oForm, "11", strSelectedItem1)

                                Catch ex As Exception

                                End Try
                            ElseIf choice = "ACTIVITY3" Then
                                osourcegrid = oForm.Items.Item(ItemUID).Specific
                                oApplication.Utilities.SetMatrixValues(osourcegrid, "V_21", sourcerowId, strSelectedItem1)
                                oApplication.Utilities.SetMatrixValues(osourcegrid, "V_20", sourcerowId, strSelectedItem2)
                            Else
                                osourcegrid = oForm.Items.Item(ItemUID).Specific
                                If choice = "ACTIVITY1" Then
                                    oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_ACTNAME", sourcerowId, strSelectedItem1)
                                    oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_MDNAME", sourcerowId, strSelectedItem2)
                                    oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_CNTID", sourcerowId, strContractID)
                                    oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_CNTNAME", sourcerowId, strContractName)
                                    oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_CUSTCNTID", sourcerowId, strCustCntID)

                                    Try
                                        oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_BOQREF", sourcerowId, getBOQReference(sourceItemCode, ItemCode, strSelectedItem2, strSelectedItem1))

                                    Catch ex As Exception
                                      End Try
                                    ' oApplication.Utilities.SetMatrixValues(osourcegrid, "U_Z_BOQREF", sourcerowId, getBOQReference(sourceItemCode, ItemCode, strSelectedItem2, strSelectedItem1))


                                Else
                                    oApplication.Utilities.SetMatrixValues(osourcegrid, sourceColumID, sourcerowId, strSelectedItem1)
                                End If


                            End If
                        End If
                    End If
                End If
            Else
                If pVal.BeforeAction = False Then
                    If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK) Then
                        BubbleEvent = False
                        Dim objGrid As SAPbouiCOM.Grid
                        Dim oedit As SAPbouiCOM.EditText
                        If pVal.ItemUID = "etFind" And pVal.CharPressed <> "13" Then
                            Dim i, j As Integer
                            Dim strItem As String
                            oForm = oApplication.SBO_Application.Forms.ActiveForm() 'oApplication.SBO_Application.Forms.GetForm("TWBS_FA_CFL", pVal.FormTypeCount)
                            objGrid = oForm.Items.Item("mtchoose").Specific
                            oedit = oForm.Items.Item("etFind").Specific
                            For i = 0 To objGrid.DataTable.Rows.Count - 1
                                strItem = ""
                                strItem = objGrid.DataTable.GetValue(0, i)
                                For j = 1 To oedit.String.Length
                                    If oedit.String.Length <= strItem.Length Then
                                        If strItem.Substring(0, j).ToUpper = oedit.String.ToUpper Then
                                            objGrid.Rows.SelectedRows.Add(i)
                                            Exit Sub
                                        End If
                                    End If
                                Next
                            Next
                        End If
                    End If
                End If
            End If
        End If
    End Sub
#End Region

End Class
