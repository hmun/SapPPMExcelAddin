Partial Class SapPPMRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.SapPPM = Me.Factory.CreateRibbonTab
        Me.SapPPMActuals = Me.Factory.CreateRibbonGroup
        Me.ButtonPPMGetActuals = Me.Factory.CreateRibbonButton
        Me.SapPPMTS = Me.Factory.CreateRibbonGroup
        Me.ButtonPPM_TS_TargetGet = Me.Factory.CreateRibbonButton
        Me.ButtonPPM_TS_Post = Me.Factory.CreateRibbonButton
        Me.ButtonPPM_TS_PostedGet = Me.Factory.CreateRibbonButton
        Me.SapPPMParam = Me.Factory.CreateRibbonGroup
        Me.ButtonPPM_ParamGet = Me.Factory.CreateRibbonButton
        Me.SapPPMLogon = Me.Factory.CreateRibbonGroup
        Me.ButtonLogon = Me.Factory.CreateRibbonButton
        Me.ButtonLogoff = Me.Factory.CreateRibbonButton
        Me.SapPPM.SuspendLayout()
        Me.SapPPMActuals.SuspendLayout()
        Me.SapPPMTS.SuspendLayout()
        Me.SapPPMParam.SuspendLayout()
        Me.SapPPMLogon.SuspendLayout()
        Me.SuspendLayout()
        '
        'SapPPM
        '
        Me.SapPPM.Groups.Add(Me.SapPPMActuals)
        Me.SapPPM.Groups.Add(Me.SapPPMTS)
        Me.SapPPM.Groups.Add(Me.SapPPMParam)
        Me.SapPPM.Groups.Add(Me.SapPPMLogon)
        Me.SapPPM.Label = "SAP PPM"
        Me.SapPPM.Name = "SapPPM"
        '
        'SapPPMActuals
        '
        Me.SapPPMActuals.Items.Add(Me.ButtonPPMGetActuals)
        Me.SapPPMActuals.Label = "PPM Actuals"
        Me.SapPPMActuals.Name = "SapPPMActuals"
        '
        'ButtonPPMGetActuals
        '
        Me.ButtonPPMGetActuals.Label = "Get Actuals"
        Me.ButtonPPMGetActuals.Name = "ButtonPPMGetActuals"
        '
        'SapPPMTS
        '
        Me.SapPPMTS.Items.Add(Me.ButtonPPM_TS_TargetGet)
        Me.SapPPMTS.Items.Add(Me.ButtonPPM_TS_Post)
        Me.SapPPMTS.Items.Add(Me.ButtonPPM_TS_PostedGet)
        Me.SapPPMTS.Label = "PPM Timesheet"
        Me.SapPPMTS.Name = "SapPPMTS"
        '
        'ButtonPPM_TS_TargetGet
        '
        Me.ButtonPPM_TS_TargetGet.Label = "Get Target hours"
        Me.ButtonPPM_TS_TargetGet.Name = "ButtonPPM_TS_TargetGet"
        '
        'ButtonPPM_TS_Post
        '
        Me.ButtonPPM_TS_Post.Label = "Post Timesheet"
        Me.ButtonPPM_TS_Post.Name = "ButtonPPM_TS_Post"
        '
        'ButtonPPM_TS_PostedGet
        '
        Me.ButtonPPM_TS_PostedGet.Label = "Get Posted TS IDs"
        Me.ButtonPPM_TS_PostedGet.Name = "ButtonPPM_TS_PostedGet"
        '
        'SapPPMParam
        '
        Me.SapPPMParam.Items.Add(Me.ButtonPPM_ParamGet)
        Me.SapPPMParam.Label = "PPM Parameter"
        Me.SapPPMParam.Name = "SapPPMParam"
        '
        'ButtonPPM_ParamGet
        '
        Me.ButtonPPM_ParamGet.Label = "Get PPM Parameters"
        Me.ButtonPPM_ParamGet.Name = "ButtonPPM_ParamGet"
        '
        'SapPPMLogon
        '
        Me.SapPPMLogon.Items.Add(Me.ButtonLogon)
        Me.SapPPMLogon.Items.Add(Me.ButtonLogoff)
        Me.SapPPMLogon.Label = "Logon"
        Me.SapPPMLogon.Name = "SapPPMLogon"
        '
        'ButtonLogon
        '
        Me.ButtonLogon.Label = "SAP Logon"
        Me.ButtonLogon.Name = "ButtonLogon"
        '
        'ButtonLogoff
        '
        Me.ButtonLogoff.Label = "SAP Logoff"
        Me.ButtonLogoff.Name = "ButtonLogoff"
        '
        'SapPPMRibbon
        '
        Me.Name = "SapPPMRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SapPPM)
        Me.SapPPM.ResumeLayout(False)
        Me.SapPPM.PerformLayout()
        Me.SapPPMActuals.ResumeLayout(False)
        Me.SapPPMActuals.PerformLayout()
        Me.SapPPMTS.ResumeLayout(False)
        Me.SapPPMTS.PerformLayout()
        Me.SapPPMParam.ResumeLayout(False)
        Me.SapPPMParam.PerformLayout()
        Me.SapPPMLogon.ResumeLayout(False)
        Me.SapPPMLogon.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SapPPM As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents SapPPMActuals As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonPPMGetActuals As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SapPPMTS As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonPPM_TS_TargetGet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPPM_TS_Post As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPPM_TS_PostedGet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SapPPMParam As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonPPM_ParamGet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SapPPMLogon As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonLogon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLogoff As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As SapPPMRibbon
        Get
            Return Me.GetRibbon(Of SapPPMRibbon)()
        End Get
    End Property
End Class
