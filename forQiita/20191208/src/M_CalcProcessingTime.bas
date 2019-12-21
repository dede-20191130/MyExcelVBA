Attribute VB_Name = "M_CalcProcessingTime"

Option Explicit
 
Private Declare PtrSafe Function QueryPerformanceCounter Lib "Kernel32" _
                           (X As Double) As Boolean
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "Kernel32" _
                           (X As Double) As Boolean
Dim Freq As Double
Dim Overhead  As Double
Dim Ctr1 As Double, Ctr2 As Double, Result As Double
 
'É~Éäïbà»â∫ÇÃçÇê∏ìxÇ≈èàóùéûä‘åvë™
Public Sub SWStart()
    If QueryPerformanceCounter(Ctr1) Then
        QueryPerformanceCounter Ctr2
        QueryPerformanceFrequency Freq
        Overhead = Ctr2 - Ctr1
    Else
        Err.Raise 513, "StopwatchError", "High-resolution counter not supported."
    End If
    QueryPerformanceCounter Ctr1
End Sub
 
Public Sub SWStop()
    QueryPerformanceCounter Ctr2
    Result = (Ctr2 - Ctr1 - Overhead) / Freq * 1000
End Sub
 
Public Function SWShow(Optional Caption As String) As Double
    SWShow = Result
End Function
