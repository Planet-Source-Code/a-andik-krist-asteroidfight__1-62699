Attribute VB_Name = "MWeapon"
Option Explicit

'----------------------------------------------------------------
' this to add create more weapon
Global Const WeaponCount    As Byte = 200 '

' Unit Properties for weapon
Type WeaponProperties
    Active      As Boolean    ' Same with UnitProperties
    FoF         As Byte       '
    Type        As Byte       '
    Power       As Byte       ' Power Weapon
    x           As Single     '
    y           As Single     '
    Angle       As Single     '
    Speed       As Single     '
    Turn        As Single     '
End Type
Public Weapon(WeaponCount)       As WeaponProperties
Public WeaponListMax As Byte

Sub WeaponCreate(FoF As Byte, x As Single, y As Single, Angle As Single, Speed As Single, Turn As Single, TypeObject As Byte, Power As Byte)
    Dim i As Byte
    
    For i = 0 To WeaponCount
        If Weapon(i).Active = False Then
            Weapon(i).Active = True
            Weapon(i).FoF = FoF
            Weapon(i).Type = TypeObject
            Weapon(i).Power = Power
            Weapon(i).x = x
            Weapon(i).y = y
            Weapon(i).Angle = Angle
            Weapon(i).Speed = Speed
            Weapon(i).Turn = Turn
            
            ' Add Visual Direct 3D Object
            WeaponFrame(i).AddVisual WeaponObject(TypeObject)
            Exit Sub
        End If
    Next i
End Sub

Sub Weapon2Fire(j As Byte, Angle As Single, Pos As Single, AngleCreate As Integer, WeaponDmg As Byte, WeaponObj As Byte, SpeedWeapon As Single, TurnWeapon As Single, Optional SingleWeapon As Boolean)
    Dim xsudut As Single
    Dim ysudut As Single

    xsudut = -Pos * Sin(DegreeToRadian(Angle))
    ysudut = -Pos * Cos(DegreeToRadian(Angle)) * -1
    WeaponCreate Unit(j).FoF, Unit(j).x + xsudut, Unit(j).y + ysudut, AddLessDegree(Angle, AngleCreate), SpeedWeapon, TurnWeapon, WeaponObj, WeaponDmg
    If SingleWeapon <> True Then
        xsudut = Pos * Sin(DegreeToRadian(Angle))
        ysudut = Pos * Cos(DegreeToRadian(Angle)) * -1
        WeaponCreate Unit(j).FoF, Unit(j).x + xsudut, Unit(j).y + ysudut, AddLessDegree(Angle, -AngleCreate), SpeedWeapon, TurnWeapon, WeaponObj, WeaponDmg
    End If
End Sub

' Calculation direction turn Add or less degree,
' like fire i want make 10 degree without this function if degree 0 or 360
' direction add 10 degree will not good
Function AddLessDegree(DegreeOrg As Single, ValueDegree As Integer) As Integer
    Dim i As Integer
    If InStr(1, Str(ValueDegree), "-") <> 0 Then
        i = DegreeOrg - Abs(ValueDegree)
        If i < 0 Then i = 360 + i
    Else
        i = DegreeOrg + ValueDegree
        If i > 360 Then i = i - 360
    End If
    AddLessDegree = i
End Function

Sub MoveWeapon()
    Dim i           As Byte
    Dim GetMouseX   As Single
    Dim GetMouseY   As Single
    Dim GetRange    As Single
    Dim GetPicNumber As Integer
    
    For i = 0 To WeaponCount
        If Weapon(i).Active = True Then
            
            GetMouseX = Weapon(0).x
            GetMouseY = Weapon(0).y
         
         '[-------------------------------------------------]
         '[ ENGINEAAK : Calculation moving unit             ]
         '[-------------------------------------------------]
            Engine Weapon(i).Angle, Weapon(i).x, Weapon(i).y, GetMouseX, GetMouseY, Weapon(i).Speed, Weapon(i).Turn
         '[-------------------------------------------------]
         '[ ENGINEAAK : Don't forget replace with new value ]
         '[-------------------------------------------------]
            Weapon(i).x = EngineResult.x
            Weapon(i).y = EngineResult.y
            Weapon(i).Angle = EngineResult.Angle
         '[-------------------------------------------------]
            
            ' Check Collision
            CollisionWeaponWtUnit i
            
            ' Set Rotation, Zoom (Scale) and Position Direct 3D
            SetPosRotObj WeaponFrame, i, Weapon(i).Angle, 0, Weapon(i).x, Weapon(i).y, 0, 1
            
            ' Check only type as weapon, missile
            If Weapon(i).x < 0 Or Weapon(i).x > 800 Or Weapon(i).y > 0 Or Weapon(i).y < -600 Then
                ' Kill Waepon (Delete Visual)
                Weapon(i).Active = False
                WeaponFrame(i).DeleteVisual WeaponObject(Weapon(i).Type)
            End If

            ' Only Text, show info
            'BackBuffer.DrawText Unit(i).x - 50, -Unit(i).y + 20, "Direct3D Object", False
            'BackBuffer.DrawText Unit(i).x - 50, -Unit(i).y + 35, "Speed:" & Unit(i).Speed & " Turn:" & Unit(i).Turn, False
            'BackBuffer.DrawText Unit(i).x - 50, -Unit(i).y + 55, "HP:" & Unit(i).HP, False
            
            WeaponListMax = i
        End If
    Next i
End Sub

Sub CollisionWeaponWtUnit(i As Byte)
    Dim j        As Byte
    Dim Range    As Single
    Dim xsudut As Single
    Dim ysudut As Single
        
    For j = 0 To UnitCount
        If Unit(j).Active = True Then
            ' Check FoF (Friend or Foe) unit and missile
            If Unit(j).FoF <> Weapon(i).FoF Then
            ' Test: Unit not small asteroid
                Range = Trigonometri(Unit(j).x, Unit(j).y, Weapon(i).x, Weapon(i).y, RESULT_RADIUS)
                If Range < Unit(j).Size Then
                    ' Calculation make hit mid weapon
                    xsudut = -12 * Sin(DegreeToRadian(45))
                    ysudut = -12 * Cos(DegreeToRadian(45)) * -1
                    
                    ' Create Hit Image Seq
                    CreateHit Int(Weapon(i).x) + xsudut, -Int(Weapon(i).y) - ysudut
                    
                    ' Alwasy Kill Weapon after hit object
                    Weapon(i).Active = False
                    WeaponFrame(i).DeleteVisual WeaponObject(Weapon(i).Type)
                    WeaponFrame(i).SetPosition Nothing, 0, 0, 0
                                                            
                    ' Decrease HP Unit / Object
                    Unit(j).HP = Unit(j).HP - Weapon(i).Power
                                        
                    ' If Unit/Object HP < 1 then create explode
                    If Unit(j).HP < 1 Then
                        KillUnit j, True
                    End If
                End If
            End If
        End If
    Next j
End Sub

