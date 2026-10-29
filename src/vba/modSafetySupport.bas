Attribute VB_Name = "modSafetySupport"
Option Explicit

' Общие вспомогательные процедуры отмены и обработки ошибок для макросов Excel и Word.
Public Function IsUserCancellation(ByVal errorNumber As Long) As Boolean
    IsUserCancellation = (errorNumber = 18)
End Function

Public Sub RestoreExcelApplicationState(ByVal screenUpdating As Boolean, _
                                        ByVal enableEvents As Boolean, _
                                        ByVal displayAlerts As Boolean, _
                                        ByVal calculationMode As XlCalculation, _
                                        ByVal oldStatusBar As Variant, _
                                        ByVal oldCancelKey As XlEnableCancelKey)
    On Error Resume Next
    Application.ScreenUpdating = screenUpdating
    Application.EnableEvents = enableEvents
    Application.DisplayAlerts = displayAlerts
    Application.Calculation = calculationMode
    Application.StatusBar = oldStatusBar
    Application.EnableCancelKey = oldCancelKey
    On Error GoTo 0
End Sub
