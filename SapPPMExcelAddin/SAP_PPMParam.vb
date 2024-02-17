' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAP_PPMParam
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
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_PPMParam")
        End Try
    End Sub

    Public Function getParam(pData As TSAP_GetParamData, ByRef pET_ZPPM_PARAM As Object, ByRef pET_RETURN As Object, Optional pOKMsg As String = "OK") As String
        getParam = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("ZPPM_PARAM_GET")
            RfcSessionManager.BeginContext(destination)
            Dim oET_ZPPM_PARAM As IRfcTable = oRfcFunction.GetTable("ET_ZPPM_PARAM")
            Dim oET_RETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            oET_ZPPM_PARAM.Clear()
            oET_RETURN.Clear()

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
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            For i As Integer = 0 To oET_RETURN.Count - 1
                getParam = getParam & ";" & oET_RETURN(i).GetValue("MESSAGE")
                If oET_RETURN(i).GetValue("TYPE") <> "S" And oET_RETURN(i).GetValue("TYPE") <> "I" And oET_RETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            getParam = If(getParam = "", pOKMsg, If(aErr = False, pOKMsg & getParam, "Error" & getParam))
            pET_ZPPM_PARAM = oET_ZPPM_PARAM
            pET_RETURN = oET_RETURN
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_PPMParam")
            getParam = "Error: Exception in getParam"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class