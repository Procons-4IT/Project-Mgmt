Public Class clsStart
    
    Shared Sub Main()
        Dim oRead As System.IO.StreamReader
        Dim LineIn, strUsr, strPwd As String
        Dim i As Integer
        Try
            Try
                oApplication = New clsListener
                oApplication.Utilities.Connect()
                oApplication.SetFilter()

                'Dim oCompanyService As SAPbobsCOM.CompanyService
                'Dim oChildren As SAPbobsCOM.GeneralDataCollection
                'oCompanyService = oApplication.Company.GetCompanyService()

                With oApplication.Company.GetCompanyService
                    CompanyDecimalSeprator = .GetAdminInfo.DecimalSeparator
                    CompanyThousandSeprator = .GetAdminInfo.ThousandsSeparator
                    LocalCurrency = .GetAdminInfo.LocalCurrency
                    systemcurrency = .GetAdminInfo.SystemCurrency
                End With
            Catch ex As Exception
                '     MessageBox.Show(ex.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                'End
                'Exit Sub
            End Try
            oApplication.Utilities.CreateTables()
            oApplication.Utilities.createHRMainAuthorization()
            oApplication.Utilities.AuthorizationCreation()
            oApplication.Utilities.AddRemoveMenus("Menu.xml")
            Dim oMenuItem As SAPbouiCOM.MenuItem
            oMenuItem = oApplication.SBO_Application.Menus.Item("Z_mnu_04")
            oMenuItem.Image = Application.StartupPath & "\Inv.bmp"
            oApplication.Utilities.Message("Project Management Addon Connected successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oApplication.Utilities.NotifyAlert()
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            System.Windows.Forms.Application.Exit()
        End Try

    End Sub

End Class
