Attribute VB_Name = "MHitandExplode"
Option Explicit

'------------------------------------------------------------------------
' Weapon Hit
'------------------------------------------------------------------------
' this to add create Hit
Global Const HitCount As Byte = 50
' Hit Properties
Type HitProperties
    Active    As Boolean    ' Active Ship
    '-----------------------------------------------------------
    x         As Integer    ' Position x, y
    y         As Integer    '
    Time      As Byte       ' Show Squence picture
    '-----------------------------------------------------------
End Type
Global Hit(HitCount) As HitProperties

'------------------------------------------------------------------------
' Explode
' Ide : If explode outer screen try no create because direct draw not draw
'       if picture out screen
'------------------------------------------------------------------------
' this to add create Hit
Public Const ExplodeCount As Byte = 20

' Explode Properties
Type ExplodeProperties
    Active    As Boolean    ' Active Ship
    '-----------------------------------------------------------
    x         As Single     ' Position x, y
    y         As Single     '
    TimeDelay As Byte       ' Delay Show Squence picture
    Time      As Byte       ' Show Squence picture
    '-----------------------------------------------------------
End Type
Public Explode(ExplodeCount) As ExplodeProperties

Public NumberExplodeCount As Byte

Sub CreateExplode(x As Single, y As Single)
    Dim i As Byte
    For i = 0 To ExplodeCount
        If Explode(i).Active = False Then
            Explode(i).Active = True
            Explode(i).x = x
            Explode(i).y = y
            Explode(i).Time = 0
            Explode(i).TimeDelay = 0
            Exit Sub
        End If
    Next i
End Sub

Sub DisplayExplode()
    Dim i As Byte
    Dim getRECT As RECT
    Dim xMed As Single
    Dim yMed As Single
    Dim ExplodeDelay As Byte
    
    For i = 0 To ExplodeCount
        If Explode(i).Active = True Then
            xMed = Explode(i).x
            yMed = -Explode(i).y
            
            ' get RECT all Explode Image Squence
            With getRECT
                .Top = 61
                .Left = 0
                .Right = 749
                .Bottom = 111
            End With
            
            DrawSquence PicBuffer, getRECT, 15, Explode(i).Time, xMed, yMed

            Explode(i).TimeDelay = Explode(i).TimeDelay + 1
            
            ExplodeDelay = 3
            
            If Explode(i).TimeDelay = ExplodeDelay Then
                Explode(i).TimeDelay = 0
                Explode(i).Time = Explode(i).Time + 1
                If Explode(i).Time > 14 Then Explode(i).Active = False
            End If
        NumberExplodeCount = i
        End If
        
    Next i
End Sub

Sub CreateHit(x As Integer, y As Integer)
    Dim i As Byte
    For i = 0 To HitCount
        If Hit(i).Active = False Then
            Hit(i).Active = True
            Hit(i).x = x
            Hit(i).y = y
            Hit(i).Time = 0
            Exit Sub
        End If
    Next i
End Sub

Sub DisplayHit()
    Dim i As Byte
    Dim getRECT As RECT
    
    For i = 0 To UnitCount
        If Hit(i).Active = True Then
            With getRECT
                .Top = 40
                .Left = (Hit(i).Time) * 16
                .Right = (Hit(i).Time + 1) * 16
                .Bottom = 55
            End With
            BackBuffer.BltFast Hit(i).x, Hit(i).y, PicBuffer, getRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY  'DDBLTFAST_WAIT
                
            Hit(i).Time = Hit(i).Time + 1
            If Hit(i).Time > 6 Then Hit(i).Active = False
            
        End If
    Next i
End Sub

Sub DrawSquence(surface As DirectDrawSurface4, picRECT As RECT, _
    ImageMany As Byte, AnimNumber As Byte, x As Single, y As Single)
    'Optional transparent As Boolean = True, Optional Clip As Boolean = True)

    Dim PicSize As Integer
    Dim PicHeight As Integer
    
    PicSize = CInt((picRECT.Right - picRECT.Left + 1) / ImageMany)
    
    Dim squenceRECT As RECT
    With squenceRECT
        .Left = picRECT.Left + PicSize * AnimNumber
        .Top = picRECT.Top
        .Right = picRECT.Left + PicSize * (AnimNumber + 1)
        .Bottom = picRECT.Bottom
        
        ' Check if image out of screen
        If y < 0 Then
            .Top = picRECT.Top - y
            y = 0
        End If
        If x < 0 Then
            .Left = .Left - x
            x = 0
        End If
        If x + .Right > 800 + .Left Then
            .Right = 800 - x + .Left
        End If
        If y + .Bottom > 600 + .Top Then
            .Bottom = 600 - y + .Top
        End If
    End With
    
    Call BackBuffer.BltFast(x, y, surface, squenceRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub
