' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAP_PPMTimesheet
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon
    Private aIntPar As SAPCommon.TStr

    Sub New(aSapCon As SapCon, ByRef pIntPar As SAPCommon.TStr)
        aIntPar = pIntPar
        Try
            sapcon = aSapCon
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_PPMTimesheet")
        End Try
    End Sub

    Public Function getTimeAttendance(pData As TSAP_GetTimeAttendanceData, ByRef pET_TIME_ATTENDANCE As Object, Optional pOKMsg As String = "OK") As String
        getTimeAttendance = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("Z_PPM_TIME_ATTENDANCE")
            RfcSessionManager.BeginContext(destination)
            Dim oET_TIME_ATTENDANCE As IRfcTable = oRfcFunction.GetTable("ET_TIME_ATTENDANCE")
            Dim oIT_EMAIL As IRfcTable = oRfcFunction.GetTable("IT_EMAIL")
            oET_TIME_ATTENDANCE.Clear()
            oIT_EMAIL.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the data values
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            Dim oIT_EMAILAppended As Boolean = False
            For Each aKvP In pData.aData.aTDataDic
                oIT_EMAILAppended = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "IT_EMAIL"
                            If Not oIT_EMAILAppended Then
                                oIT_EMAIL.Append()
                                oIT_EMAILAppended = True
                            End If
                            oIT_EMAIL.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            '            For i As Integer = 0 To oET_RETURN.Count - 1
            '            getTimeAttendance = getTimeAttendance & ";" & oET_RETURN(i).GetValue("MESSAGE")
            '           If oET_RETURN(i).GetValue("TYPE") <> "S" And oET_RETURN(i).GetValue("TYPE") <> "I" And oET_RETURN(i).GetValue("TYPE") <> "W" Then
            '           aErr = True
            '           End If
            '            Next i
            getTimeAttendance = If(getTimeAttendance = "", pOKMsg, If(aErr = False, pOKMsg & getTimeAttendance, "Error" & getTimeAttendance))
            pET_TIME_ATTENDANCE = oET_TIME_ATTENDANCE
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_PPMTimesheet")
            getTimeAttendance = "Error: Exception in getTimeAttendance"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function getTsIds(pData As TSAP_GetTsIdsData, ByRef pET_ZPPM_S_TS_IDS As Object, ByRef pET_RETURN As Object, Optional pOKMsg As String = "OK") As String
        getTsIds = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("Z_PPM_TS_GET_IDS")
            RfcSessionManager.BeginContext(destination)
            Dim oET_ZPPM_S_TS_IDS As IRfcTable = oRfcFunction.GetTable("ET_ZPPM_S_TS_IDS")
            Dim oET_RETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            Dim oIT_ZPPM_S_TS_IDS As IRfcTable = oRfcFunction.GetTable("IT_ZPPM_S_TS_IDS")
            oET_ZPPM_S_TS_IDS.Clear()
            oET_RETURN.Clear()
            oIT_ZPPM_S_TS_IDS.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the data values
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            Dim oIT_ZPPM_S_TS_IDSAppended As Boolean = False
            For Each aKvP In pData.aData.aTDataDic
                oIT_ZPPM_S_TS_IDSAppended = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "IT_ZPPM_S_TS_IDS"
                            If Not oIT_ZPPM_S_TS_IDSAppended Then
                                oIT_ZPPM_S_TS_IDS.Append()
                                oIT_ZPPM_S_TS_IDSAppended = True
                            End If
                            oIT_ZPPM_S_TS_IDS.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            For i As Integer = 0 To oET_RETURN.Count - 1
                getTsIds = getTsIds & ";" & oET_RETURN(i).GetValue("MESSAGE")
                If oET_RETURN(i).GetValue("TYPE") <> "S" And oET_RETURN(i).GetValue("TYPE") <> "I" And oET_RETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            getTsIds = If(getTsIds = "", pOKMsg, If(aErr = False, pOKMsg & getTsIds, "Error" & getTsIds))
            pET_ZPPM_S_TS_IDS = oET_ZPPM_S_TS_IDS
            pET_RETURN = oET_RETURN
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_PPMTimesheet")
            getTsIds = "Error: Exception in getTsIds"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function postTS(pData As TSAP_PostTSData, ByRef pET_CATS As Object, ByRef pET_RETURN As Object, Optional pOKMsg As String = "OK") As String
        postTS = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("Z_PPM_TIMESHEET_UPDATE")
            RfcSessionManager.BeginContext(destination)
            Dim oET_CATS As IRfcTable = oRfcFunction.GetTable("ET_CATS")
            Dim oET_RETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            Dim oIT_CATS As IRfcTable = oRfcFunction.GetTable("IT_CATS")
            oET_CATS.Clear()
            oET_RETURN.Clear()
            oIT_CATS.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the data values
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            Dim oIT_CATSAppended As Boolean = False
            For Each aKvP In pData.aData.aTDataDic
                oIT_CATSAppended = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "IT_CATS"
                            If Not oIT_CATSAppended Then
                                oIT_CATS.Append()
                                oIT_CATSAppended = True
                            End If
                            oIT_CATS.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            For i As Integer = 0 To oET_RETURN.Count - 1
                postTS = postTS & ";" & oET_RETURN(i).GetValue("MESSAGE")
                If oET_RETURN(i).GetValue("TYPE") <> "S" And oET_RETURN(i).GetValue("TYPE") <> "I" And oET_RETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            postTS = If(postTS = "", pOKMsg, If(aErr = False, pOKMsg & postTS, "Error" & postTS))
            pET_CATS = oET_CATS
            pET_RETURN = oET_RETURN
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_PPMTimesheet")
            postTS = "Error: Exception in postTS"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class