Public Module modVariables

    Public oApplication As clsListener
    Public strSQL As String
    Public cfl_Text As String
    Public cfl_Btn As String

    Public CompanyDecimalSeprator As String
    Public CompanyThousandSeprator As String
    Public LocalCurrency As String
    Public systemcurrency As String
    Public blnMasterExport As Boolean = False
    Public blnFEExport As Boolean = False
    Public blnDocumentItem As Boolean = False
    Public frm_SPBP, frmSourceForm, aSourceForm, frm_ScaleDiscountForm, frm_FilterTableForm, rm_SourceForm As SAPbouiCOM.Form
    Public frmSourceMatrix As SAPbouiCOM.Matrix
    Public frmSourceGrid As SAPbouiCOM.Grid
    Public strItemSelectionQuery As String = ""
    Public rowtoDelete1 As Integer

    Public intSelectedMatrixrow As Integer = 0
    Public strSPBPCode As String = ""
    Public strSPItemCode As String = ""
    Public ErrorLogFile As String
    Public sSearchList As String = ""
    Public loginEmployeeID As String = ""
    Public EntryChoice As String = ""
    Public strApprovalType As String = ""
    Public blnSourceForm As Boolean = False
    Public strSourceformEmpID As String = ""
    Public strMdbFilePath As String
    Dim strFileName As String
    Public strSelectedFilepath, sPath, strSelectedFolderPath, strFilepath As String
    Public strHeaderQry, strDetailQry As String


    Public Enum ValidationResult As Integer
        CANCEL = 0
        OK = 1
    End Enum

    Public Enum DocumentType As Integer
        RENTAL_QUOTATION = 1
        RENTAL_ORDER
        RENTAL_RETURN
    End Enum

    Public Const frm_ContractCFL As String = "frm_ContractCFL"

    Public Const mnu_ContrctTimeSheet As String = "Z_mnu_117"
    Public Const frm_ContrctTimeSheet As String = "frm_ContractTime"
    Public Const xml_ContrctTimeSheet As String = "xml_Contract_Timesheet.xml"

    Public Const frm_DisRule As String = "frm_DisRule"
    Public Const xml_DisRule As String = "frm_DisRule.xml"
   
    Public Const frm_BOMWizard As String = "frm_BOMWizard"
    Public Const xml_BOMWizard As String = "frm_BOMCopy.xml"

    Public Const mnu_PrjReports As String = "Z_mnu_PrjRpt"
    Public Const frm_PrjReports As String = "frm_PrjReports"
    Public Const xml_PrjReports As String = "frm_PrjReports.xml"

    'phase II Modules
    Public Const frm_ProjectMaster As String = "711"
    Public Const mnu_ChangePassword As String = "Z_mnu_chPwd"
    Public Const frm_ChangePassword As String = "frm_ChngPwd"

    Public Const frm_MyTasks As String = "frm_MyTasks"
    Public Const mnu_MyTasks As String = "Z_mnu_MyTasks"
    Public Const xml_MyTasks As String = "frm_MyTasks.xml"



    Public Const frm_Contract As String = "frm_Contract"
    Public Const xml_Contract As String = "xml_Contract.xml"
    Public Const mnu_Contract As String = "Z_mnu_305"

    Public Const frm_BudgetContract As String = "frm_CntBudget"
    Public Const xml_BudgetContract As String = "frm_BudgetDetails_Contract.xml"
    Public Const frm_PORequest As String = "1470000200"

    'End Phase II
    Public Const frm_BOQDeatils As String = "frm_BOQDeatils"
    Public Const mnu_BOQDetails As String = "Z_mnu_22"
    Public Const xml_BOQDetails As String = "frm_BOQDefine.xml"

    Public Const frm_Itemgroup As String = "frm_ItemGroup"
    Public Const frm_WAREHOUSES As Integer = 62

    Public Const frm_SalesQuotation As Integer = 149
    Public Const frm_SalesOrder As Integer = 139
    Public Const frm_Delivery As Integer = 140
    Public Const frm_Return As Integer = 180

    Public Const frm_Downpaymentrequest As Integer = 65308
    Public Const frm_DownpaymentInvoice As Integer = 65300
    Public Const frm_InvoicePayment As Integer = 60090
    Public Const frm_reverseInvoice As Integer = 60091
    Public Const frm_EmployeePostion As String = "frm_Position"
    Public Const frm_PO As String = "142"
    Public Const frm_GRPO As String = "143"
    Public Const frm_APReturn As String = "182"
    Public Const frm_APInvoice As String = "141"
    Public Const frm_APCR As String = "181"
    Public Const frm_APDOWN As String = "65309"
    Public Const frm_APDOWNREquest As String = "65301"
    Public Const frm_Reservce As String = "60092"

    Public Const frm_GoodsIssue As String = "720"
    Public Const frm_GoodsReceipt As String = "721"



    Public Const frm_ARInvoice As Integer = 133
    Public Const frm_ARCreditNote As Integer = 179

    Public Const mnu_ItemGroup As String = "Z_Itms01"
  
    Public Const frm_ItemMaster As String = "150"
    Public Const frm_BPMaster As String = "134"

    Public Const frm_Activity As String = "frm_Activity"
    Public Const frm_Module As String = "frm_Module"
    Public Const frm_EmpTime As String = "frm_EmpTime"
    Public Const frm_ProjectTime As String = "frm_PrjTime"
    Public Const frm_Budget As String = "frm_Budget"
    Public Const frm_BudgetEntry As String = "frm_BudgetEntry"
    Public Const frm_BudgetTemplate As String = "frm_BudgetTemplate"
    Public Const frm_Report As String = "frm_Report"
    Public Const frm_Expances As String = "frm_Expances"
    Public Const frm_ExpEntry As String = "frm_ExpEntry"
    Public Const frm_ACCEntry As String = "frm_ACCEntry"
    Public Const frm_Timesheet As String = "frm_Entry"
    Public Const frm_Login As String = "frm_Login"
    Public Const frm_Approal As String = "frm_Details"
    Public Const frm_LoginSetup As String = "frm_LogSetup"
    Public Const frm_LeaveType As String = "frm_LeaveType"
    Public Const frm_Leaverequest As String = "frm_LeaveRequest"


    Public Const mnu_FIND As String = "1281"
    Public Const mnu_ADD As String = "1282"
    Public Const mnu_Delete As String = "1283"
    Public Const mnu_CLOSE As String = "1286"
    Public Const mnu_NEXT As String = "1288"
    Public Const mnu_PREVIOUS As String = "1289"
    Public Const mnu_FIRST As String = "1290"
    Public Const mnu_LAST As String = "1291"
    Public Const mnu_ADD_ROW As String = "1292"
    Public Const mnu_DELETE_ROW As String = "1293"
    Public Const mnu_Duplicate_Row As String = "1294"
    Public Const mnu_TAX_GROUP_SETUP As String = "8458"
    Public Const mnu_DEFINE_ALTERNATIVE_ITEMS As String = "11531"
    Public Const mnu_Activity As String = "Z_mnu_02"
    Public Const mnu_Moudule As String = "Z_mnu_03"
    Public Const mnu_Budget As String = "Z_mnu_05"
    Public Const mnu_BudgetEntry As String = "Z_mnu_105"
    Public Const mnu_BudgetTemplate As String = "Z_mnu_205"
    Public Const mnu_EmpTime As String = "Z_mnu_06"
    Public Const mnu_PrjTime As String = "Z_mnu_07"
    Public Const mnu_report As String = "Z_mnu_08"
    Public Const mnu_Expances As String = "Z_mnu_09"
    Public Const mnu_ExpEntry As String = "Z_mnu_10"
    Public Const mnu_ExpApproval As String = "Z_mnu_11"
    Public Const mnu_Loginsetup As String = "Z_mnu_15"
    Public Const mnu_EmpPosition As String = "Z_mnu_16"
    Public Const mnu_ExpClaim As String = "Z_mnu_17"

    Public Const mnu_ACCClaim As String = "Z_mnu_218"
    Public Const mnu_LeaveType As String = "Z_mnu_18"
    Public Const mnu_Leaverequest As String = "Z_mnu_20"
    Public Const mnu_LeaveApproval As String = "Z_mnu_19"


    Public Const frm_ChoosefromList As String = "frm_CFL"
    Public Const frm_ChoosefromList1 As String = "frm_CFL1"

    Public Const xml_MENU As String = "Menu.xml"
    Public Const xml_MENU_REMOVE As String = "RemoveMenus.xml"
    Public Const xml_ItemGroup As String = "xml_ItemGroup.xml"
    Public Const xml_Activity As String = "frm_Activity.xml"
    Public Const xml_Module As String = "frm_Module.xml"
    Public Const xml_EmpTime As String = "frm_TimeSheetEntry.xml"
    Public Const xml_Budget As String = "frm_Budget.xml"
    Public Const xml_BudgetEntry As String = "frm_BudgetEntry.xml"
    Public Const xml_BudgetTemplate As String = "frm_BudgetTemplate.xml"
    Public Const xml_Prjtime As String = "frm_ProjectTime.xml"
    Public Const xml_Report As String = "frm_report.xml"
    Public Const xml_Expances As String = "frm_Expances.xml"
    Public Const xml_ExpEntry As String = "frm_ExpanceEntry1.xml"

    Public Const xml_AcctEntry As String = "frm_ACCExpanceEntry1.xml"
    Public Const xml_Login As String = "frm_Login.xml"
    Public Const xml_LoginSetup As String = "frm_LoginSetup.xml"
    Public Const xml_Details As String = "frm_ApprovalDetails.xml"
    Public Const xml_EmpoyeePostion As String = "frm_Position.xml"
    Public Const xml_LeaveType As String = "frm_Leave.xml"
    Public Const xml_LeaveReqest As String = "frm_LeaveRequest.xml"





End Module
