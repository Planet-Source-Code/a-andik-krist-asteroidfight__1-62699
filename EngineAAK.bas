Attribute VB_Name = "EngineAAK"
' [923342960150006  S T T D T T S  600051069243329]
' [===============================================]
' [               EngineAAK ver 1.1               ]
' [         Modify from my First Program:         ]
' [             ".. HomingMissile .."             ]
' [ For : Arcade/RTS/RPG/Race.., in test is work. ]
' [-----------------------------------------------]
' [  Perbaikan dari program pertama saya bernama  ]
' [             ".. HomingMissile .."             ]
' [       Untuk : Arcade/RTS/RPG/Race......       ]
' [-----------------------------------------------]
' [                                               ]
' [              By: A. Andik Krist.              ]
' [              -------------------              ]
' [              JAKARTA - INDONESIA              ]
' [                                               ]
' [-----------------------------------------------]
' [ Not for sale or commercial without permission ]
' [-----------------------------------------------]
' [ If you might use this for your project please ]
' [         give me credits if you do so.         ]
' [-----------------------------------------------]
' [                                               ]
' [        for Comments, Suggestions & Ideas      ]
' [          E-mails me: aakchat@yahoo.com        ]
' [                                               ]
' [===============91923=29873=30006===============]
'                    _______          ______
'             _______\.kKk..\        /.KKk./
' + ENGINE + /.aaaa...\.kKk..\      /.KKk./
'           /.aAAAAa...\.kKkk.\    /.KKk./
'          /.aAa.___ .....kKKk.\  /.KKk./
'         /.aAa./   \ .....kKKkkk/.KKk./
'        /.aAa./     \_____.kKkk..KKKk..\_
'       /.aAa./_____  ____ \.kKk...__.Kk..\_
'      /.aAAAaaaa../ /.aa.\ \.kKk..\ \_.Kk..\_
'     /.aAAAaaaa../ /.aAAa.\ \.kKk..\  \_.Kk..\_
'    /.aAAa._____/ /.aAaaAa.\ \.kKk..\   \_.Kk..\_
'   /.aAAa./      /.aA./\.Aa.\ \.kKk..\    \_.Kk..\_
'  /.aaaa./      /.aA./__\.Aa.\ \.kKk..\     \_.Kk..\_
' /______/      /.aA.______.Aa.\ \______\      \______\
'              /.aA./      \.Aa.\
'             /____/        \____\ + Ver 1.1 +
'
' -----------------------2982005-----------------------

Option Explicit

Public Const Pi As Single = 3.14159265358979

' Engine Properties to store new result value : X, Y and Angle
Public Type EngineProperties
    x     As Single
    y     As Single
    Angle As Single
End Type
Public EngineResult As EngineProperties

Public Enum TrigonometriResultType
    RESULT_DEGREES = 0  ' result angle in Degrees
    RESULT_RADIAN = 1   ' result for Radian
    RESULT_RADIUS = 2   ' result for Radius
End Enum

'--------------------------------------------------------------------
' Engine:
' Angle        = Position Angle (in Degrees) Unit (Nose Direction)
' x, y         = Position unit in x, y
' xDist, yDist = Position direction unit x, y
' MoveSpeed    = Speed unit
' TurnSpeed    = Turn (Spin) unit
'--------------------------------------------------------------------
Public Sub Engine(ByVal Angle As Single, ByVal x As Single, ByVal y As Single, _
                  ByVal xDist As Single, ByVal yDist As Single, _
                  ByVal MoveSpeed As Single, ByVal TurnSpeed As Single)
    
    Dim Degree    As Single
    Dim Radian    As Single
    Dim xNew      As Single
    Dim yNew      As Single
    
    Radian = -DegreeToRadian(Angle)
    
    xNew = MoveSpeed * Cos(Radian)
    yNew = MoveSpeed * Sin(Radian) * Deg_DirOffset(Angle, True)
    
    Degree = Trigonometri(x, y, xDist, yDist)
    
    '[==================== B E G I N =====================]
    '[ make object not fibrator/(bergetar)                ]
    '[----------------------------------------------------]
    If Angle > Degree - (TurnSpeed - 1) Then
        If Angle < Degree + (TurnSpeed - 1) Then
            Angle = Degree
            If Angle = 360 Then
                Angle = 359
                Degree = 359
            End If
        End If
    End If

    '[====================== E N D =======================]
    
    '[==================== B E G I N =====================]
    '[ Calculation Turning or Spin (0-359 Degree) Object ]
    '[----------------------------------------------------]
    If Angle = Degree Then
    '[ Let line empty for object to not fibrator too      ]
    Else
        Angle = Angle + DirectTurn(Degree, Angle, TurnSpeed)
        If Angle > 359 Then
            Angle = 1
        Else
            If Angle < 0 Then Angle = 359
        End If
    End If
    '[====================== E N D =======================]

    '[ Store new Result value x, y and angle              ]
    EngineResult.x = x + xNew
    EngineResult.y = y + yNew
    EngineResult.Angle = Angle
  
End Sub

'--------------------------------------------------------------------
' Trigonometri:
' Search Angle (Degrees) / Radian / Radius, from linier line with
' 2 point, point1 (x1,y1) and point2 (x2,y2)
' TypeResult = Degrees / Radian / Radius (Range)
'--------------------------------------------------------------------
Public Function Trigonometri(ByVal x1 As Single, ByVal y1 As Single, _
                             ByVal x2 As Single, ByVal y2 As Single, _
                             Optional TypeResult As TrigonometriResultType = RESULT_DEGREES) As Single
    
    Dim xAbsTrig             As Single
    Dim yAbsTrig             As Single
    Dim RadianTrig           As Single
    Dim RadiusTrig           As Single
    Dim DeggresTrig          As Single
    Dim SinTrig              As Single
    Dim PosDeggresTrig       As Single
    
    xAbsTrig = (x1 - x2)
    yAbsTrig = (y1 - y2)
    
    If Abs(yAbsTrig) <> 0 And Abs(xAbsTrig) <> 0 Then
        RadianTrig = Atn(Abs(yAbsTrig) / Abs(xAbsTrig))
        SinTrig = Sin(RadianTrig)
        RadiusTrig = Abs(yAbsTrig) / SinTrig
    Else
        If Abs(yAbsTrig) <> 0 Then
            RadiusTrig = Abs(yAbsTrig)
        Else
            RadiusTrig = Abs(xAbsTrig)
        End If
    End If

    DeggresTrig = (RadianTrig * 180 / Pi)   'hanya untuk mengetahui derajat tidak bisa
                                            'dihitung dengan Sin/Cos/Tan langsung
    
    If xAbsTrig < 0 And yAbsTrig <= 0 Then
        ' Kwadran 1 / Direction Right
        If yAbsTrig = 0 Then
            PosDeggresTrig = 0
        Else
            PosDeggresTrig = DeggresTrig
        End If
    Else
        If xAbsTrig >= 0 And yAbsTrig < 0 Then
            ' Kwadran 2 / Direction Down
            If xAbsTrig = 0 Then
                PosDeggresTrig = 90
            Else
                PosDeggresTrig = 180 - DeggresTrig
            End If
        Else
            If xAbsTrig >= 0 And yAbsTrig >= 0 Then
                ' Kwadran 3 / Direction Left
                If xAbsTrig = 0 Then
                    PosDeggresTrig = 270
                Else
                    PosDeggresTrig = 180 + DeggresTrig
                End If
            Else
                ' Kwadran 4 / Direction Up
                PosDeggresTrig = 270 + (90 - DeggresTrig)
            End If
        End If
    End If
    
    ' Result of type RADIAN
    If TypeResult = 1 Then
        Trigonometri = RadianTrig
    Else
       ' Result of type RADIUS
        If TypeResult = 2 Then
            Trigonometri = RadiusTrig
        Else
            ' Result of type DEGREES
            Trigonometri = PosDeggresTrig
        End If
    End If
        
End Function

'--------------------------------------------------------------------
' Function direction Object (Unit) to make turn Left or Right result
' value=TurnSpeed, but only add minus (-) or not
'--------------------------------------------------------------------
' AngleFirstObject  : Angel (in degrees) first Object/unit
' AngleSecondObject : Angel (in degrees) second Object/unit
'--------------------------------------------------------------------
Public Function DirectTurn(ByVal AngleFirstObject As Single, _
                           ByVal AngleSecondObject As Single, _
                           ByVal TurnSpeed As Single) As Single
                           
    Dim DeggresValueLeft   As Integer
    Dim DeggresValueRight  As Integer

    DeggresValueLeft = Abs(AngleFirstObject - AngleSecondObject)
    DeggresValueRight = Abs(359 - Abs(AngleFirstObject - AngleSecondObject))

    If AngleFirstObject < AngleSecondObject Then
        If DeggresValueLeft < DeggresValueRight Then
            ' Putar Kiri 1  / Spin left
            DirectTurn = -TurnSpeed  '-
        Else
            ' Putar Kanan 1 / Spin Right
            DirectTurn = TurnSpeed   '+
        End If
    Else
        If DeggresValueLeft < DeggresValueRight Then
            ' Putar Kanan 2 / Spin Right
            DirectTurn = TurnSpeed   '+
        Else
            ' Putar Kiri 2  / Spin Left
            DirectTurn = -TurnSpeed  '-
        End If
    End If
End Function

Private Function Deg_DirOffset(ByVal DegrresOrginal As Single, Optional Result As Boolean = False) As Integer
    If DegrresOrginal > 0 Or DegrresOrginal < 90 Then
        If Result = False Then
            Deg_DirOffset = (90 - DegrresOrginal)
        Else
            Deg_DirOffset = -1
        End If
    Else
        Deg_DirOffset = 1
    End If
End Function

' Convertion Angle (in Degree) to Radian
Function DegreeToRadian(Angle As Single) As Single
    DegreeToRadian = (Angle * Pi) / 180
End Function

