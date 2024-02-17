' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAP_PPMActuals
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
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_PPMActuals")
        End Try
    End Sub

    Public Function getActuals(pData As TSAP_GetActualsData, ByRef pET_ACTUALS As Object, ByRef pET_RETURN As Object, Optional pOKMsg As String = "OK") As String
        getActuals = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("Z_PPM_GET_ACTUALS_NEW")
            RfcSessionManager.BeginContext(destination)
            Dim oET_RETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            Dim oET_ACTUALS As IRfcTable = oRfcFunction.GetTable("ET_ACTUALS")
            Dim oIT_WBS As IRfcTable = oRfcFunction.GetTable("IT_PSP")
            Dim oIT_CER As IRfcTable = oRfcFunction.GetTable("IT_CER")
            Dim oIT_ORD As IRfcTable = Nothing
            Try
                oIT_ORD = oRfcFunction.GetTable("IT_ORD")
            Catch Ex As System.Exception
                oIT_ORD = Nothing
            End Try
            oET_RETURN.Clear()
                oET_ACTUALS.Clear()
                oIT_WBS.Clear()
                oIT_CER.Clear()
                If Not oIT_ORD Is Nothing Then
                    oIT_ORD.Clear()
                End If

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
                Dim oIT_WBSAppended As Boolean = False
                Dim oIT_CERAppended As Boolean = False
                Dim oIT_ORDAppended As Boolean = False
                For Each aKvP In pData.aData.aTDataDic
                    oIT_WBSAppended = False
                    oIT_CERAppended = False
                    oIT_ORDAppended = False
                    aTDataRec = aKvP.Value
                    For Each aTStrRec In aTDataRec.aTDataRecCol
                        Select Case aTStrRec.Strucname
                            Case "IT_PSP"
                                If Not oIT_WBSAppended Then
                                    oIT_WBS.Append()
                                    oIT_WBSAppended = True
                                End If
                                oIT_WBS.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                            Case "IT_CER"
                                If Not oIT_CERAppended Then
                                    oIT_CER.Append()
                                    oIT_CERAppended = True
                                End If
                                oIT_CER.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                            Case "IT_ORD"
                                If Not oIT_ORD Is Nothing Then
                                    If Not oIT_ORDAppended Then
                                        oIT_ORD.Append()
                                        oIT_ORDAppended = True
                                    End If
                                    oIT_ORD.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                                End If
                        End Select
                    Next
                Next
                ' call the BAPI
                oRfcFunction.Invoke(destination)
                Dim aErr As Boolean = False
                For i As Integer = 0 To oET_RETURN.Count - 1
                    getActuals = getActuals & ";" & oET_RETURN(i).GetValue("MESSAGE")
                    If oET_RETURN(i).GetValue("TYPE") <> "S" And oET_RETURN(i).GetValue("TYPE") <> "I" And oET_RETURN(i).GetValue("TYPE") <> "W" Then
                        aErr = True
                    End If
                Next i
                getActuals = If(getActuals = "", pOKMsg, If(aErr = False, pOKMsg & getActuals, "Error" & getActuals))
                pET_ACTUALS = oET_ACTUALS
                pET_RETURN = oET_RETURN
            Catch Ex As System.Exception
                MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_PPMActuals")
                getActuals = "Error: Exception in getActuals"
            Finally
                RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class