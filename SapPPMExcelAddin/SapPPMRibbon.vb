' Copyright 2021 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports Microsoft.Office.Tools.Ribbon

Public Class SapPPMRibbon
    Private aSapCon
    Private aSapGeneral
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private Sub SapPPMRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        aSapGeneral = New SapGeneral
    End Sub

    Private Function checkCon() As Integer
        Dim aSapConRet As Integer
        Dim aSapVersionRet As Integer
        checkCon = False
        log.Debug("checkCon - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            Exit Function
        End If
        log.Debug("checkCon - " & "checking Connection")
        aSapConRet = 0
        If aSapCon Is Nothing Then
            Try
                aSapCon = New SapCon()
            Catch ex As SystemException
                log.Warn("checkCon-New SapCon - )" & ex.ToString)
            End Try
        End If
        Try
            aSapConRet = aSapCon.checkCon()
        Catch ex As SystemException
            log.Warn("checkCon-aSapCon.checkCon - )" & ex.ToString)
        End Try
        If aSapConRet = 0 Then
            log.Debug("checkCon - " & "checking version in SAP")
            Try
                aSapVersionRet = aSapGeneral.checkVersionInSAP(aSapCon)
            Catch ex As SystemException
                log.Warn("checkCon - )" & ex.ToString)
            End Try
            log.Debug("checkCon - " & "aSapVersionRet=" & CStr(aSapVersionRet))
            If aSapVersionRet = True Then
                log.Debug("checkCon - " & "checkCon = True")
                checkCon = True
            Else
                log.Debug("checkCon - " & "connection check failed")
            End If
        End If
    End Function

    Private Sub ButtonLogon_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogon.Click
        Dim aConRet As Integer

        log.Debug("ButtonLogon_Click - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            log.Debug("ButtonLogon_Click - " & "Version check failed")
            Exit Sub
        End If
        log.Debug("ButtonLogon_Click - " & "creating SapCon")
        If aSapCon Is Nothing Then
            aSapCon = New SapCon()
        End If
        log.Debug("ButtonLogon_Click - " & "calling SapCon.checkCon()")
        aConRet = aSapCon.checkCon()
        If aConRet = 0 Then
            log.Debug("ButtonLogon_Click - " & "connection successfull")
            MsgBox("SAP-Logon successful! ", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap PPM")
        Else
            log.Debug("ButtonLogon_Click - " & "connection failed")
            aSapCon = Nothing
        End If
    End Sub

    Private Sub ButtonLogoff_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogoff.Click
        log.Debug("ButtonLogoff_Click - " & "starting logoff")
        If Not aSapCon Is Nothing Then
            log.Debug("ButtonLogoff_Click - " & "calling aSapCon.SAPlogoff()")
            aSapCon.SAPlogoff()
            aSapCon = Nothing
        End If
        log.Debug("ButtonLogoff_Click - " & "exit")
    End Sub

    Private Sub ButtonPPMGetActuals_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPPMGetActuals.Click
        Dim aSapPPMRibbonActuals As New SapPPMRibbonActuals
        If checkCon() = True Then
            aSapPPMRibbonActuals.getActuals(pSapCon:=aSapCon)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonPPMGetActuals_Click")
        End If
    End Sub

    Private Sub ButtonPPM_TS_TargetGet_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPPM_TS_TargetGet.Click
        Dim aSapPPMRibbonTimesheet As New SapPPMRibbonTimesheet
        If checkCon() = True Then
            aSapPPMRibbonTimesheet.getTimeAttendance(pSapCon:=aSapCon)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonPPM_TS_TargetGet_Click")
        End If
    End Sub

    Private Sub ButtonPPM_TS_PostedGet_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPPM_TS_PostedGet.Click
        Dim aSapPPMRibbonTimesheet As New SapPPMRibbonTimesheet
        If checkCon() = True Then
            aSapPPMRibbonTimesheet.getTsIds(pSapCon:=aSapCon)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonPPM_TS_PostedGet_Click")
        End If
    End Sub

    Private Sub ButtonPPM_TS_Post_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPPM_TS_Post.Click
        Dim aSapPPMRibbonTimesheet As New SapPPMRibbonTimesheet
        If checkCon() = True Then
            aSapPPMRibbonTimesheet.postTS(pSapCon:=aSapCon)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonPPM_TS_Post_Click")
        End If
    End Sub

    Private Sub ButtonPPM_ParamGet_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPPM_ParamGet.Click
        Dim aSapPPMRibbonParam As New SapPPMRibbonParam
        If checkCon() = True Then
            aSapPPMRibbonParam.getParameter(pSapCon:=aSapCon)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonPPM_ParamGet_Click")
        End If
    End Sub
End Class
