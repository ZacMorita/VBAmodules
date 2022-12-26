Attribute VB_Name = "CustomFunctions"
Option Explicit

Public Function DecToBinStr(ByRef inputDecimal As Long, Optional ByRef MaxBits As Long = 16) As String
    Dim hBinStr As String
    Dim i As Long
    Dim BinRemain As Long
    
    hBinStr = ""
    BinRemain = inputDecimal
    
    For i = 1 To MaxBits
        hBinStr = CStr(BinRemain Mod 2) & hBinStr
        BinRemain = Int(BinRemain / 2)
    Next i
    
    DecToBinStr = hBinStr

End Function
