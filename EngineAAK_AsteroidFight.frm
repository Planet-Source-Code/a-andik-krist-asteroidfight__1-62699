VERSION 5.00
Begin VB.Form EngineAAK_AsteroidFight 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "EngineAAK_AsteroidFight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' [923342960150006  S T T D T T S  600051069243329]
' [===============================================]
' [        Introduction: EngineAAK Ver 1.1        ]
' [        -------------------------------        ]
' [           Title Game: AsteroidFight           ]
' [===============================================]
' [  Player (blue ship) will fight with 1 / more  ]
' [   enemy Fighter (red ship) in Asteroid Land   ]
' [  use asteroid to avoid them (enemy will shot  ]
' [           if player ship in front).           ]
' [  Shoot Big Asteroid to make them drop Module  ]
' [  for Player (Like FirePower and HP Restore),  ]
' [  but not alwasy work and only create smalles  ]
' [                   asteroid.                   ]
' [-----------------------------------------------]
' [ Gameplay:                                     ]
' [ Left Arrow Key  = TURN LEFT                   ]
' [ Right Arrow Key = TURN RIGHT                  ]
' [ Ctrl Key        = FIRE                        ]
' [ Esc             = QUIT GAME                   ]
' [-----------------------------------------------]
' [ Not for sale or commercial without permission ]
' [-----------------------------------------------]
' [              By: A. Andik Krist.              ]
' [              -------------------              ]
' [              JAKARTA - INDONESIA              ]
' [-----------------------------------------------]
' [                                               ]
' [        for Comments, Suggestions & Ideas      ]
' [          E-mails me: aakchat@yahoo.com        ]
' [               Date: 17-Sep-2005               ]
' [                                               ]
' [===============91923=29873=30006===============]

Option Explicit


Private Sub Form_Load()
    Randomize Time
    
    '--------------------------------------------------------------
    ScrWidth = 800   ' 640 ' 800
    ScrHeight = 600  ' 480 ' 600
    '--------------------------------------------------------------
    ' Init Direct3D
    D3DInit EngineAAK_AsteroidFight, ScrWidth, ScrHeight, 16
    '--------------------------------------------------------------
    ' Initialize Frame Direct3D like Root, Camera, Light
    FrameD3DInit ScrWidth, ScrHeight
    '--------------------------------------------------------------
    ' Load Direct 3D Object
    loadDirect3DUnit
    ' Load Direct 3D Weapon
    loadDirect3DWeapon
    ' Load Direct 3D Map
    loadDirect3DMap
    ' Load Direct 3D Module
    loadDirect3DModule
    ' Sound Init
    SoundInit EngineAAK_AsteroidFight
    '--------------------------------------------------------------
    LoadBMPandSurface PicBuffer, App.Path & "\DataImage.bmp", PicBufferRECT, 0
    '--------------------------------------------------------------
    Init_Games
    
    GameLoops
    
End Sub

