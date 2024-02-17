' Copyright 2021 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SapPPMExcelAddIn

    Private Sub SapPPMExcelAddIn_Startup() Handles Me.Startup
        log4net.Config.XmlConfigurator.Configure()
    End Sub

    Private Sub SapPPMExcelAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
