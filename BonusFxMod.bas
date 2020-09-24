Attribute VB_Name = "modBonusFx"

'   A BONUS SUB-ROUTINE USING LINEAR INTERPOLATION :)


Sub LinearInterpolation(ByVal FromX As Long, ByVal FromY As Long, ByVal ToX As Long, ByVal ToY As Long, ByVal t As Single, ByRef ReturnX As Long, ByRef ReturnY As Long)
    'this subroutine calculates the interpolated value
    't always ranges from 0 to 1
    '-------------------------------------------------
    Dim DiffX As Long, DiffY As Long
    DiffX = ToX - FromX
    DiffY = ToY - FromY
    ReturnX = FromX + (DiffX * t)
    ReturnY = FromY + (DiffY * t)
End Sub
Sub FxUsage()
    Dim i As Integer
    Dim dt As Single
    Dim StartX As Long, StartY As Long, EndX As Long, EndY As Long
    Dim PlotX As Long, PlotY As Long
    StartX = 10: StartY = 100
    EndX = 100: EndY = 50
    For i = 0 To 100 Step 5
        dt = i / 100
        LinearInterpolation StartX, StartY, EndX, EndY, dt, PlotX, PlotY
        frmMain.pic.PSet (PlotX, PlotY), vbWhite
    Next i
End Sub
    
