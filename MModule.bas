Attribute VB_Name = "MModule"
Option Explicit

'----------------------------------------------------------------
' this to add create more module
Global Const ModuleCount    As Byte = 5 '

' Unit Properties for Module
Type ModuleProperties
    Active      As Boolean
    Type        As Byte
    x           As Single
    y           As Single
    Angle       As Single
    AngleTurn   As Single
End Type
Public Module(ModuleCount)  As ModuleProperties

Sub ModuleCreate(x As Single, y As Single, Angle As Single, TypeModule As Byte)
    Dim i As Byte
    
    For i = 0 To ModuleCount
        If Module(i).Active = False Then
            Module(i).Active = True
            Module(i).Type = TypeModule
            Module(i).x = x
            Module(i).y = y
            Module(i).Angle = Angle
            
            ' First Create Module Base
            ModuleFrame(i).AddVisual ModuleObject(0)
            
            ' Then Create Module Color
            ModuleFrame(i).AddVisual ModuleObject(TypeModule)
            Exit Sub
        End If
    Next i
End Sub

Sub ModuleMove()
    Dim i As Byte
    
    For i = 0 To ModuleCount
        If Module(i).Active = True Then
        
         '[-------------------------------------------------]
         '[ ENGINEAAK : Calculation moving module           ]
         '[-------------------------------------------------]
            Engine Module(i).Angle, Module(i).x, Module(i).y, Unit(0).x, Unit(0).y, 1, 0.25
         '[-------------------------------------------------]
         '[ ENGINEAAK : Don't forget replace with new value ]
         '[-------------------------------------------------]
            Module(i).x = EngineResult.x
            Module(i).y = EngineResult.y
            Module(i).Angle = EngineResult.Angle
         '[-------------------------------------------------]
            Module(i).AngleTurn = Module(i).AngleTurn + 1
        
            SetPosRotObj ModuleFrame, i, Module(i).AngleTurn, Module(i).AngleTurn, Module(i).x, Module(i).y, 0, 3
            
            ' Check Module Out of screen and kill
            If Module(i).x < 0 Or Module(i).x > ScrWidth Or Module(i).y > 0 Or Module(i).y < -ScrHeight Then
                ' Kill module Base
                Module(i).Active = False
                ModuleFrame(i).DeleteVisual ModuleObject(0)
                
                ' Kill module Base
                Module(i).Active = False
                ModuleFrame(i).DeleteVisual ModuleObject(Module(i).Type)
            End If
            
            ' Check collision module with player
            If Trigonometri(Module(i).x, Module(i).y, Unit(0).x, Unit(0).y, RESULT_RADIUS) < Unit(0).Size Then
                ' Kill module Base
                Module(i).Active = False
                ModuleFrame(i).DeleteVisual ModuleObject(0)
                                
                ' Check Module Type (Base Module not include)
                If Module(i).Type = 1 Then
                    PlayerWeapon = PlayerWeapon + 1     ' Add More Fire
                    PlayerWeaponDelay = 5 + PlayerWeapon
                    If PlayerWeapon > 3 Then
                        PlayerWeapon = 3
                        PlayerWeaponDelay = 5 + PlayerWeapon
                        
                        ' Max weapon player, don't worry i give you Bonus Fire: "Around Fire Destroy EveryThing if Hit"
                        Dim j As Byte
                        For j = 1 To 36
                            Weapon2Fire 0, Module(i).Angle, 0, 10 * j, 200, 0, 10, 0, True
                        Next j
                    End If
                Else
                    If Module(i).Type = 2 Then
                        Unit(0).HP = Unit(0).HP + 50
                        If Unit(0).HP > PlayerPower Then Unit(0).HP = PlayerPower
                    End If
                End If
                
                ' Kill Module Color
                Module(i).Active = False
                ModuleFrame(i).DeleteVisual ModuleObject(Module(i).Type)
                
            End If
        End If
    Next i
End Sub

