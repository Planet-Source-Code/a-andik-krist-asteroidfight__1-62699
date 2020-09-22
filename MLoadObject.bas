Attribute VB_Name = "MLoadObject"
Option Explicit

'----------------------------------------------------------------
' Direct 3D Frame/Object
'----------------------------------------------------------------
' Frames
Public UnitFrame(UnitCount)     As Direct3DRMFrame3

' Meshes (loaded 3D objects from a *.x file)
'
'                 +-> Use () because we don't how many
'                 |   Direct 3D Object want loading
'                 |
Public UnitObject()             As Direct3DRMMeshBuilder3

'----------------------------------------------------------------
' Direct 3D Weapon
'----------------------------------------------------------------
' Frames
Public WeaponFrame(WeaponCount) As Direct3DRMFrame3
' Object
Public WeaponObject()           As Direct3DRMMeshBuilder3

'----------------------------------------------------------------
' Direct 3D Map
'----------------------------------------------------------------
' Frames
Public MapFrame                 As Direct3DRMFrame3
' Object
Public MapObject                As Direct3DRMMeshBuilder3

'----------------------------------------------------------------
' Direct 3D Module for player
'----------------------------------------------------------------
' Frames
Public ModuleFrame(ModuleCount) As Direct3DRMFrame3
' Object
Public ModuleObject()           As Direct3DRMMeshBuilder3

Sub loadDirect3DUnit()
    Dim i As Byte
    
    ReDim UnitObject(3)     ' 0 = Fighter
                            ' 1 = Enemy Fighter
                            ' 2 = Asteroid
    Set UnitObject(0) = D3D.CreateMeshBuilder()
    With UnitObject(0)
        .LoadFromFile App.Path & "\Object\PlayerFighter.x", 0, 0, Nothing, Nothing
        .ScaleMesh 1, 1, 1
    End With
    ' Load Direct 3D Object Enemy Fighter
    Set UnitObject(1) = D3D.CreateMeshBuilder()
    With UnitObject(1)
        .LoadFromFile App.Path & "\Object\Enemy08Ace.x", 0, 0, Nothing, Nothing
        .ScaleMesh 1, 1, 1
    End With
    ' Load Direct 3D Object Asteroid 1
    Set UnitObject(2) = D3D.CreateMeshBuilder()
    With UnitObject(2)
        .LoadFromFile App.Path & "\Object\Asteroid1.x", 0, 0, Nothing, Nothing
        .ScaleMesh 2, 2, 2
    End With
    ' Load Direct 3D Object Asteroid 2
    Set UnitObject(3) = D3D.CreateMeshBuilder()
    With UnitObject(3)
        .LoadFromFile App.Path & "\Object\Asteroid2.x", 0, 0, Nothing, Nothing
        .ScaleMesh 2, 2, 2
    End With
    '--------------------------------------------------------------
    ' Set Frame for Direct 3D Object
    For i = 0 To UnitCount
        Set UnitFrame(i) = D3D.CreateFrame(FrameRoot)
    Next i

End Sub

Sub loadDirect3DWeapon()
    Dim i As Byte
    
    ReDim WeaponObject(1)
    
    ' Load Direct 3D Object Weapon
    Set WeaponObject(0) = D3D.CreateMeshBuilder()
    With WeaponObject(0)
        .LoadFromFile App.Path & "\Object\Weapon.x", 0, 0, Nothing, Nothing
        .ScaleMesh 1, 1, 1
        '.SetColorRGB 0, 255, 0  ' Original color is red, but i want SetColor with green
    End With
    ' Load Direct 3D Object Missile
    Set WeaponObject(1) = D3D.CreateMeshBuilder()
    With WeaponObject(1)
        .LoadFromFile App.Path & "\Object\WeaponEnemyMiss01.x", 0, 0, Nothing, Nothing
        .ScaleMesh 2, 2, 2
    End With
    '--------------------------------------------------------------
    ' Set Frame for Direct 3D Object
    For i = 0 To WeaponCount
        Set WeaponFrame(i) = D3D.CreateFrame(FrameRoot)
    Next i
End Sub

Sub loadDirect3DModule()
    Dim i As Byte
    
    ReDim ModuleObject(2)
    
    ' Load Direct 3D Object Module Base
    Set ModuleObject(0) = D3D.CreateMeshBuilder()
    With ModuleObject(0)
        .LoadFromFile App.Path & "\Object\PlayerModBase.x", 0, 0, Nothing, Nothing
        .ScaleMesh 1, 1, 1
    End With
    ' Load Direct 3D Object Module Color (use .SetColorRGB you can change color this object)
    Set ModuleObject(1) = D3D.CreateMeshBuilder()
    With ModuleObject(1)
        .LoadFromFile App.Path & "\Object\PlayerModColor.x", 0, 0, Nothing, Nothing
        .ScaleMesh 1, 1, 1
        .SetColorRGB 255, 0, 0
    End With
    ' Load Direct 3D Object Module HP Restore
    Set ModuleObject(2) = D3D.CreateMeshBuilder()
    With ModuleObject(2)
        .LoadFromFile App.Path & "\Object\PlayerModPowerHP.x", 0, 0, Nothing, Nothing
        .ScaleMesh 1, 1, 1
    End With
    '--------------------------------------------------------------
    ' Set Frame for Direct 3D Object
    For i = 0 To ModuleCount
        Set ModuleFrame(i) = D3D.CreateFrame(FrameRoot)
    Next i
End Sub

Sub loadDirect3DMap()
    Set MapObject = D3D.CreateMeshBuilder()
    With MapObject
        .LoadFromFile App.Path & "\Object\Map.x", 0, 0, Nothing, Nothing
        .ScaleMesh 1, 1, 1
    End With
    Set MapFrame = D3D.CreateFrame(FrameRoot)
    
    MapFrame.AddVisual MapObject
End Sub

Sub Delete_All()
    Dim i As Byte
    
    Set DX = Nothing
    Set DD = Nothing
    Set D3D = Nothing
    Set Primary = Nothing
    Set BackBuffer = Nothing
    Set D3D_Device = Nothing
    Set D3D_ViewPort = Nothing
    
    For i = 0 To UnitCount
        Set UnitFrame(i) = Nothing
    Next i
    For i = 0 To 3      ' See upper how Object Unit use
        Set UnitObject(i) = Nothing
    Next i
    
    For i = 0 To WeaponCount
        Set WeaponFrame(i) = Nothing
    Next i
    For i = 0 To 1      ' See upper how Object Weapon use
        Set WeaponObject(i) = Nothing
    Next i
        
    For i = 0 To ModuleCount
        Set ModuleFrame(i) = Nothing
    Next i
    For i = 0 To 2      ' See upper how Object Module use
        Set ModuleObject(i) = Nothing
    Next i
    
    Set MapFrame = Nothing
    Set MapObject = Nothing
       
End Sub

