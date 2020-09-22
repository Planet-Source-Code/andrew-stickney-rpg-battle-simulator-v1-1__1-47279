VERSION 5.00
Begin VB.Form Main 
   Caption         =   "RPG Battle Simulator V1.1 - Andrew (Booda) Stickney"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8160
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Main.frx":0CCA
   ScaleHeight     =   296
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   544
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox EWinText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   5160
      TabIndex        =   18
      Text            =   "0"
      Top             =   3840
      Width           =   720
   End
   Begin VB.TextBox PWinText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Left            =   4200
      TabIndex        =   17
      Text            =   "0"
      Top             =   3840
      Width           =   720
   End
   Begin VB.TextBox BattleCountText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   3240
      TabIndex        =   16
      Text            =   "0"
      Top             =   3840
      Width           =   720
   End
   Begin VB.PictureBox EAttackButton 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   4920
      Picture         =   "Main.frx":76C0C
      ScaleHeight     =   600
      ScaleWidth      =   1080
      TabIndex        =   15
      Top             =   720
      Width           =   1080
   End
   Begin VB.PictureBox FightButton 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   3240
      Picture         =   "Main.frx":78E0E
      ScaleHeight     =   600
      ScaleWidth      =   1680
      TabIndex        =   14
      Top             =   720
      Width           =   1680
   End
   Begin VB.PictureBox PAttackButton 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   2160
      Picture         =   "Main.frx":7C2D0
      ScaleHeight     =   600
      ScaleWidth      =   1080
      TabIndex        =   13
      Top             =   720
      Width           =   1080
   End
   Begin VB.TextBox EStatText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   5
      Left            =   7200
      TabIndex        =   12
      Top             =   3840
      Width           =   720
   End
   Begin VB.TextBox EStatText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   4
      Left            =   7200
      TabIndex        =   11
      Top             =   3240
      Width           =   720
   End
   Begin VB.TextBox EStatText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   3
      Left            =   7200
      TabIndex        =   10
      Top             =   2640
      Width           =   720
   End
   Begin VB.TextBox EStatText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   2
      Left            =   7200
      TabIndex        =   9
      Top             =   2040
      Width           =   720
   End
   Begin VB.TextBox EStatText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   1
      Left            =   7200
      TabIndex        =   8
      Top             =   1440
      Width           =   720
   End
   Begin VB.TextBox EStatText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   0
      Left            =   7200
      TabIndex        =   7
      Top             =   840
      Width           =   720
   End
   Begin VB.TextBox HStatText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Index           =   5
      Left            =   1200
      TabIndex        =   6
      Top             =   3840
      Width           =   720
   End
   Begin VB.TextBox HStatText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Index           =   4
      Left            =   1200
      TabIndex        =   5
      Top             =   3240
      Width           =   720
   End
   Begin VB.TextBox HStatText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Index           =   3
      Left            =   1200
      TabIndex        =   4
      Top             =   2640
      Width           =   720
   End
   Begin VB.TextBox HStatText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   720
   End
   Begin VB.TextBox HStatText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   1440
      Width           =   720
   End
   Begin VB.TextBox HStatText 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   360
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   720
   End
   Begin VB.Image ResetButton 
      Height          =   600
      Left            =   2160
      Top             =   3720
      Width           =   960
   End
   Begin VB.Label BattleText 
      BackStyle       =   0  'Transparent
      Caption         =   "Press Fight Button to begin the battle."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1680
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   3120
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' **********************************************************************
' * RPG Simulator v1.1                                                 *
' * Andrew (Booda) Stickney                                            *
' * July 30, 2003                                                      *
' *                                                                    *
' * This program is basically just a battle engine,                    *
' * you can enter your own stats and fight mok battles,                *
' * i use this program to debug my own rpgs, and to make               *
' * sure the enemies are fair and not overly hard.                     *
' * Feel free to chop this program up anyway you want,                 *
' * but if you add something or make it better, which should be easy   *
' * because this program is as basic as it gets, just email me, please *
' *                                                                    *
' * New : V1.1                                                         *
' * I have added a new stat, Damage, now there is more random damage   *
' * Also, i have added more damage types, such as Normal, Double,      *
' *     Critical Hits, this adds more to the battle...a little suspense*
' * I added a counter at the bottom, it has number of turns, and       *
' *     battles won.                                                   *
' * I fixed the battles so there is more randomness to every aspect    *
' *     A beginner to a rpg veteran, like myself should enjoy...       *
' *                                                                    *
' * And lets admit the earlier version sucked....                      *
' *     I should have my head examined for not testing my own progs    *
' *                                                                    *
' * Only known problem, and really isn't a severe one, enemies with    *
' *     only slightly higher stats will really wail on you.            *
' * If somebody can come up with some sort of inhibitor, that will not *
' *     effect the randomness...I would like to know, thank you        *
' **********************************************************************

'variables for player stats
Dim PHp As Single, PStr As Single, PDef As Single, PAtc As Single, PDdg As Single, PDmg As Single
'variables for enemy stats
Dim EHp As Single, EStr As Single, EDef As Single, EAtc As Single, EDdg As Single, EDmg As Single
'variables for battle calculations
Dim PHitRoll As Single, EHitRoll As Single, PDamageRoll As Single, EDamageRoll As Single
Dim PDodgeRoll As Single, EDodgeRoll As Single
Dim PDamageAdd As Single, EDamageAdd As Single
Dim DPHp As Single, DEHp As Single
'these variables are for damage variation
Dim NormalHit As Single, DoubleHit As Single, ExcellentHit As Single
Dim PFightTog As Boolean, EFightTog As Boolean 'these toggles are for restoring hp, to max after death
Dim BattleCountTog As Single, PWinTog As Single, EWinTog As Single 'these are just counters

Private Sub FightButton_Click()
'change color of caption
BattleText.ForeColor = RGB(255, 255, 255)
'reset the fight round tog
BattleCountTog = 0
BattleCountText.Text = 0

'Check through all the text boxes for valid data
'if return false -
    If CheckInput = False Then
        BattleText.Caption = "There is invalid data, in one of the input boxes. Please corret it, then press Fight Button Again."
        Beep
        Exit Sub
    End If

'if return true =
'drop the fight button , to keep from being pressed
'until battle is over
    FightButton.Top = FightButton.Top + 1000
'disable all the text boxes, to prevent input during battle
    Dim Index As Integer
    For Index = 0 To 5
        HStatText(Index).Enabled = False
        EStatText(Index).Enabled = False
    Next Index

'put text box input in appropriate variables
    'Player Stats
    PHp = HStatText(0).Text
    PStr = HStatText(1).Text
    PDef = HStatText(2).Text
    PAtc = HStatText(3).Text
    PDdg = HStatText(4).Text
    PDmg = HStatText(5).Text
    'Enemy Stats
    EHp = EStatText(0).Text
    EStr = EStatText(1).Text
    EDef = EStatText(2).Text
    EAtc = EStatText(3).Text
    EDdg = EStatText(4).Text
    EDmg = EStatText(5).Text
    
'calculate the damage additive
'damage additive is basically taking the strikers str and subtracting
'the strikees def, this rnd(total +- 1) is added into the damageroll itself
    'calculate player first
    PDamageAdd = (PStr - EDef)
    'calculate enemy
    EDamageAdd = (EStr - PDef)
    
'caculate first strike
'Very Simple, player gets PAtc *3, enemy gets EAtc *3
'whoever has the highest gets first strike
'this formula will be explained in the players strike
'routine for it is the same as hit or miss

'randomize the number, or roll the di for you d&d fans
PHitRoll = Int(Rnd * (PAtc * 2)) + 1
EHitRoll = Int(Rnd * (EAtc * 2)) + 1

'check to see who goes first
    If PHitRoll > EHitRoll Then
        BattleText.Caption = "You get first strike!"
        PAttackButton.Top = PAttackButton.Top - 1000 'bring player button up
    Else: 'if missed to opposite
        BattleText.Caption = "Enemy gets first strike!"
        EAttackButton.Top = EAttackButton.Top - 1000 'bring up enemy button
    End If
    
'put hp values in dummys, so when battle is over you can put them back
'also a quick note, if you choose to use stat altering magic in your game
'you should put all stats in dummy variables, of course this is common sense
'but only store dummys if togs are false, when they are true, means one died
'and hp was set back to max, this keeps the dummys from loading lesser values

If PFightTog = False Then DPHp = PHp
If EFightTog = False Then DEHp = EHp
    
'set togs to true, until one or the other dies
PFightTog = True
EFightTog = True
'finally we exit the sub, the stage is set, because
'who gets first strike there button is showing, and from here its
'a matter of eventing back and forth, until someone is dead
End Sub

Private Sub Form_Load()
'randomize seed, you can not have a roleplaying game without this
'this is the true heart of all roleplaying games, randomness
Randomize Timer

'Move the attack buttons of page, until ready to be pressed
    PAttackButton.Top = PAttackButton.Top + 1000
    EAttackButton.Top = EAttackButton.Top + 1000
    
'put fight toggles at false, until fight starts
PFightTog = False
EFightTog = False

End Sub



Private Sub PAttackButton_Click()
'change color of caption
BattleText.ForeColor = RGB(0, 255, 255)
'update battlecounttog
BattleCountTog = BattleCountTog + 1
BattleCountText.Text = BattleCountTog

Randomize Timer
'the player attacks routine, basically three checks are perfomed here
'hit or miss, damage type if hit, and victory
    'calculate the hit or miss first
    'what the formula is is this, phitroll = rnd(patc*2)
    'ehitroll = rnd(eddg*2), if phitroll is highest its a hit
    'this formula is based on if player and enemy is equal, the
    'equal opurtunite to hit and miss for both, this seems
    'a pain, and i will get complaints, but it makes sense in
    'reality. Also for every one point the stat is over the other
    'that is two points to advantage, this can be changed...
        
    'note - with this system, if the enemy's dodge is greater
    'by 100 points, the player would never hit them, this
    'is where a little fair play comes in, how could you possibly
    'expect a enemy of that maganitude to be defeated, seriously.
    
    'roll number
    PHitRoll = Int(Rnd * (PAtc * 2)) + 1
    EHitRoll = Int(Rnd * (EDdg * 2)) + 1
        
    'check to see if hit, if miss, exit sub
    If PHitRoll < EHitRoll Then 'this misses and exits sub, bringing up enemys turn
        BattleText.Caption = "You Miss!!"
        GoTo SkipDamage
    End If
    
    'now this is where it gets a little tricky, i have added alot of variables
    'here to keep with the randomness of the battle and also the strengths
    'prevail, as it should be
    
    'First calculate the percentages
    'These percentages give random advantages of better hits
    'first i take the PAtc - eddg, then divide 100 by this total,
    'then multiply by 20, the 20 is to make percentages come out as whole.
    'also with this formula, the higher the stat difference the less number
    'to randomize, so therefore better odds, im proud of this formula, and anybody
    'who can make it better, more power to you..
    
    'if the if this randomized is lower than 75% of that difference then the hit is normal
    'if greater than 75% and less than 90% then double damage
    'anything greater will give excellenthit, basically totaldamage*2
    'now the math
    'first do pecentages
    Dim Total As Single, Total2 As Single
    Total = PAtc - EDdg                         'subtract the two
    If Total = 0 Or Total < 0 Then Total = 1    'check if 0 or less, and make 1
    If Total > 100 Then Total = 100             'this system is base 100, basically if the enemy is 100 times better
    Total = Int(100 / Total)                    'you wouldn't be here, because you would miss him, but if you are
                                                'you would get and excellant hit everytime, practicaly
                                                
    Total = Total * 20                  'multiply by 20 to make percentages come out even
                                        'otherwise, higher you go, total would be 1 or 2
                                        'and your really cant random that, and have percentages.
    
    Total2 = Int(Rnd * Total) + 1       'randomize total
    
    NormalHit = Int(0.75 * Total)       'multiply total x 75%
    DoubleHit = Int(0.9 * Total)        'mulitply total x 85%
    ExcellentHit = Total                'this is redundant, but keeps things clean
    Total = PDmg + PDamageAdd           'this places total damage for excellent hits
    If Total < 1 Then Total = 1         'notice excellent has no random, it is the max total * 2
    
    'next we calculate the random Damage
    'basically we roll the half the damage stat and add half the damage
    PDamageRoll = Int(Rnd * (PDmg / 2)) + 1 + Int((PDmg / 2))
    'add the damage additive, if it is negative it takes off of damage roll
    'this keeps in line with the defense of the opponent, therefore
    'if greater we inflict less damage, also checks to see if damage is negative
    'then makes it 1, the reason i take off at least one, is because
    'i have flashbacks to older roleplaying games and the no effect strike
    'oooo...i hated that, nothing made me more mad, if you make the effort
    'at least take off 1....im sorry its late and i digress...
    PDamageRoll = PDamageRoll + (Int(Rnd * PDamageAdd / 2) + 1 + Int((PDamageAdd / 2)))
    If PDamageRoll < 1 Then PDamageRoll = 1
    
    'now we compare hitpercentages with hitroll to see what damage
    'first instance normal hit
    If Total2 <= NormalHit Then
        BattleText.Caption = "You hit the enemy and inflict" + Str$(PDamageRoll) + " damage...."
        EHp = EHp - PDamageRoll
    End If
    'second instance double damage
    If Total2 > NormalHit And Total2 <= DoubleHit Then
        BattleText.Caption = "Double Hit!! You inflict" + Str$(PDamageRoll * 2) + " damage..."
        EHp = EHp - (PDamageRoll * 2)
    End If
    'third instance and rare
    If Total2 > DoubleHit And Total2 <= ExcellentHit Then
       BattleText.Caption = "Excellent Hit!!! You inflict" + Str$(Total * 2) + " damage..."
        EHp = EHp - (Total * 2)
    End If
    
    'update enemy hit points
    If EHp < 0 Then EHp = 0
        EStatText(0).Text = EHp
        
'check to see if enemy is dead
    If EHp = 0 Then
        EHp = DEHp: EFightTog = False 'reset enemys hp
        EStatText(0).Text = EHp
        BattleText.Caption = "You have defeated the enemy!!  Press Fight to begin another battle.."
        'reset buttons
        PAttackButton.Top = PAttackButton.Top + 1000
        FightButton.Top = FightButton.Top - 1000
        ResetInput 'enable textboxes
        'update pwintextobox
        PWinTog = PWinTog + 1
        PWinText.Text = PWinTog
        Exit Sub
    End If

SkipDamage: 'you missed or the enemy isn't dead
    'reset buttons
    PAttackButton.Top = PAttackButton.Top + 1000
    EAttackButton.Top = EAttackButton.Top - 1000
    'now ready for enemy's attack, which is the same, just opposite formulas
End Sub

Private Sub PAttackButton_GotFocus()
'When the first button has focus by hitting tab, it will set the focus
'to the first input, to avoid premature activation. also check to see if
'it isn't ready to be pressed
    If PAttackButton.Top <> 48 Then
        HStatText(0).SetFocus
    End If
    
End Sub

Private Function CheckInput() As Boolean

' Set error trap for function to return a false when
' there is an error
On Error GoTo NegativeInput

' variables for loop, and a dummy to convert text into
' string, and if an error occurs it returns a false
Dim Index As Integer, Dummy As Integer

' Loops for textboxes
For Index = 0 To 5
'convert text to 1 if it equals 0, saves errors later
    If HStatText(Index).Text = "0" Then HStatText(Index).Text = "1"
    If EStatText(Index).Text = "0" Then EStatText(Index).Text = "1"
'now you convert each text box, to a number and put it in dummy
    Dummy = HStatText(Index).Text
    Dummy = EStatText(Index).Text 'have to convert both sets
Next Index

' If there were no errors, then return true
    CheckInput = True
    Exit Function

NegativeInput:
'there was an error, so return false
CheckInput = False

End Function

Private Sub ResetInput()
'this sub basicaly enables the inputs after battles
Dim Index As Integer
For Index = 0 To 5
    HStatText(Index).Enabled = True
    EStatText(Index).Enabled = True
Next Index

End Sub

Private Sub ResetButton_Click()
'Reset the win togs
PWinTog = 0
EWinTog = 0
PWinText.Text = 0
EWinText.Text = 0

End Sub

Private Sub EAttackButton_Click()
'The enemies fight routine is a cookie cutter version of the players
'the only difference is the stats used are the enemies and the text displayed.
'Basically im lazy and just didn't want to comment this, but its real simple
'look at the players attack routine and change the p's to e's or vice versa
'and you get this one, so for interpretation, look at the players
'Also if you couldn't figure this out, are you sure you should be
'Programming, and if this comment hurt your feelings, i was joking...
'This is why i put this chunk of code at the end of the proggy...and
'after typing all this, i might as well have commented this section, oh well
'too late now...

BattleText.ForeColor = RGB(255, 0, 0)
BattleCountTog = BattleCountTog + 1
BattleCountText.Text = BattleCountTog

Randomize Timer
    
    EHitRoll = Int(Rnd * (EAtc * 2)) + 1
    PHitRoll = Int(Rnd * (PDdg * 2)) + 1
        
    If EHitRoll < PHitRoll Then
        BattleText.Caption = "The Enemy Misses!!"
        GoTo SkipDamage
    End If
    
    Dim Total As Single, Total2 As Single
    Total = EAtc - PDdg
    If Total = 0 Or Total < 0 Then Total = 1
    If Total > 100 Then Total = 100
    Total = Int(100 / Total)
                                                
    Total = Total * 20
    
    Total2 = Int(Rnd * Total) + 1
    
    NormalHit = Int(0.75 * Total)
    DoubleHit = Int(0.9 * Total)
    ExcellentHit = Total
    Total = EDmg + EDamageAdd
    If Total < 1 Then Total = 1
    
    EDamageRoll = Int(Rnd * (EDmg / 2)) + 1 + Int((EDmg / 2))
    
    EDamageRoll = EDamageRoll + (Int(Rnd * EDamageAdd / 2) + 1 + Int((EDamageAdd / 2)))
    If EDamageRoll < 1 Then EDamageRoll = 1
    
    If Total2 <= NormalHit Then
        BattleText.Caption = "You have been hit and the enemy has inflicted" + Str$(EDamageRoll) + " damage...."
        PHp = PHp - EDamageRoll
    End If
    If Total2 > NormalHit And Total2 <= DoubleHit Then
        BattleText.Caption = "Double Hit!! The Enemy inflicts" + Str$(EDamageRoll * 2) + " damage..."
        PHp = PHp - (EDamageRoll * 2)
    End If
    If Total2 > DoubleHit And Total2 <= ExcellentHit Then
       BattleText.Caption = "Critical Hit!!! The Enemy inflicts" + Str$(Total * 2) + " damage..."
        PHp = PHp - (Total * 2)
    End If
    
    If PHp < 0 Then PHp = 0
        HStatText(0).Text = PHp
        
    If PHp = 0 Then
        PHp = DPHp: PFightTog = False
        HStatText(0).Text = PHp
        BattleText.Caption = "You have been defeated by the enemy!!  Press Fight to begin another battle.."
        
        EAttackButton.Top = EAttackButton.Top + 1000
        FightButton.Top = FightButton.Top - 1000
        ResetInput
        
        EWinTog = EWinTog + 1
        EWinText.Text = EWinTog
        Exit Sub
    End If

SkipDamage:
    
    EAttackButton.Top = EAttackButton.Top + 1000
    PAttackButton.Top = PAttackButton.Top - 1000
    
End Sub

