VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  '2D
   BackColor       =   &H80000004&
   Caption         =   "Widerstands-Farbcode Berechnung"
   ClientHeight    =   3555
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6975
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6975
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chk5 
      Caption         =   "5 Streifenwiderstand"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   2160
      Value           =   1  'Aktiviert
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Beenden"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Line Line4 
      X1              =   3120
      X2              =   3120
      Y1              =   2160
      Y2              =   3360
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   3120
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3120
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   120
      Y1              =   2160
      Y2              =   3360
   End
   Begin VB.Label R 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0C0&
      Caption         =   "R ="
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEinheitmax 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblEinheitmin 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblRmax 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblRmin 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Rmax 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rmax ="
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Rmin 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rmin ="
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblProzent 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblEinheit 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblR 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblWert 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "10%"
      Height          =   255
      Index           =   4
      Left            =   4200
      TabIndex        =   11
      Top             =   240
      Width           =   570
   End
   Begin VB.Label lblWert 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "0.01"
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   10
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblWert 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   9
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblWert 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   8
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblWert 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   7
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Wert 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Index           =   4
      Left            =   4320
      TabIndex        =   6
      ToolTipText     =   "Toleranz"
      Top             =   600
      Width           =   330
   End
   Begin VB.Label Wert 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Index           =   3
      Left            =   3720
      TabIndex        =   5
      ToolTipText     =   "Multiplikator"
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Wert 
      BackColor       =   &H00000000&
      Height          =   1335
      Index           =   2
      Left            =   3240
      TabIndex        =   4
      ToolTipText     =   "Kennziffer 3"
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Wert 
      BackColor       =   &H00000000&
      Height          =   1335
      Index           =   1
      Left            =   2760
      TabIndex        =   3
      ToolTipText     =   "Kennziffer 2"
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Ausgefüllt
      Height          =   375
      Left            =   240
      Shape           =   4  'Gerundetes Rechteck
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Ausgefüllt
      Height          =   375
      Left            =   5400
      Shape           =   4  'Gerundetes Rechteck
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0FFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   1365
      Left            =   4920
      Shape           =   2  'Oval
      Top             =   585
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   1365
      Left            =   1560
      Shape           =   2  'Oval
      Top             =   585
      Width           =   495
   End
   Begin VB.Label Wert 
      BackColor       =   &H00000000&
      Height          =   1335
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      ToolTipText     =   "Kennziffer 1"
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0FFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   1365
      Left            =   1800
      Top             =   585
      Width           =   3330
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "Optionen"
      Begin VB.Menu mnuWiderstand 
         Caption         =   "&Widerstand"
         Begin VB.Menu mnu4 
            Caption         =   "4 Streifen"
         End
         Begin VB.Menu mnu5 
            Caption         =   "5 Streifen"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuDrossel 
         Caption         =   "&Drossel"
      End
      Begin VB.Menu mnuStrich 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBeenden 
         Caption         =   "&Beenden"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuSilver 
         Caption         =   "Silber"
      End
      Begin VB.Menu mnuGold 
         Caption         =   "Gold"
      End
      Begin VB.Menu mnuBlack 
         Caption         =   "Schwarz"
      End
      Begin VB.Menu mnuBrown 
         Caption         =   "Braun"
      End
      Begin VB.Menu mnuRed 
         Caption         =   "Rot"
      End
      Begin VB.Menu mnuOrange 
         Caption         =   "Orange"
      End
      Begin VB.Menu mnuYellow 
         Caption         =   "Gelb"
      End
      Begin VB.Menu mnuGreen 
         Caption         =   "Grün"
      End
      Begin VB.Menu mnuBlue 
         Caption         =   "Blau"
      End
      Begin VB.Menu mnuViolet 
         Caption         =   "Violett"
      End
      Begin VB.Menu mnuGrey 
         Caption         =   "Grau"
      End
      Begin VB.Menu mnuWhite 
         Caption         =   "Weiss"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim Prozent As Double
Dim Multiplikator As Currency
Dim Farbe As Integer
Dim Widerstand

Private Sub chk5_Click()

If chk5.Value = 1 Then
    Wert(2).Visible = True
    lblWert(2).Visible = True
    Wert(3).Left = 3720
    lblWert(3).Left = 3600
Else
    Wert(2).Visible = False
    lblWert(2).Visible = False
    Wert(3).Left = 3240
    lblWert(3).Left = 3120
End If
End Sub

Private Sub cmdExit_Click()

End
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdExit.FontUnderline = True
End Sub

Private Sub Form_Load()
Farbe = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


cmdExit.FontUnderline = False
End Sub

Private Sub mnu4_Click()

mnu5.Checked = False
mnu4.Checked = True

Call Widerstand_Ein

Wert(2).Visible = False
lblWert(2).Visible = False
Wert(3).Left = 3240
lblWert(3).Left = 3120

Call Back_Normal
End Sub

Private Sub mnu5_Click()

mnu5.Checked = True
mnu4.Checked = False

Call Widerstand_Ein

Wert(2).Visible = True
lblWert(2).Visible = True
Wert(3).Left = 3720
lblWert(3).Left = 3600

Call Back_Normal
End Sub

Private Sub mnuBeenden_Click()

Call cmdExit_Click
End Sub

Private Sub mnuBlack_Click()

Wert(i).BackColor = vbBlack
If i = 3 Then
    Multiplikator = "1"
    lblWert(i).Caption = Multiplikator
Else
    lblWert(i).Caption = "0"
End If

End Sub

Private Sub mnuBlue_Click()

Wert(i).BackColor = vbBlue
Select Case i
Case 3
    Multiplikator = "1000000"
    lblWert(i).Caption = "10^6"
Case 4
    Prozent = "0.25"
    lblWert(i).Caption = Prozent & "%"
Case Else
    lblWert(i).Caption = "6"
End Select
End Sub

Private Sub mnuBrown_Click()

Wert(i).BackColor = 16512
Select Case i
Case 3
    Multiplikator = "10"
    lblWert(i).Caption = Multiplikator
Case 4
    Prozent = "1"
    lblWert(i).Caption = Prozent & "%"
Case Else
    lblWert(i).Caption = "1"
End Select

'Shape1.FillColor = &HFF8080
'Shape1.BorderColor = &HFF8080
'Shape2.FillColor = &HFF8080
'Shape2.BorderColor = &HFF8080
'Shape3.FillColor = &HFF8080
'Shape3.BorderColor = &HFF8080
End Sub

Private Sub mnuDrossel_Click()

Call mnu4_Click
mnu4.Checked = False
mnuDrossel.Checked = True

Shape1.FillColor = &H40C0&
Shape1.BorderColor = &H40C0&
Shape2.FillColor = &H40C0&
Shape2.BorderColor = &H40C0&
Shape3.FillColor = &H40C0&
Shape3.BorderColor = &H40C0&

frmMain.Caption = "Drossel-Farbcode Berechnung"
End Sub

Private Sub mnuGold_Click()

Wert(i).BackColor = &HC0C0&
If i = 3 Then
    Multiplikator = "0.1"
    lblWert(i).Caption = Multiplikator
Else
    Prozent = "5"
    lblWert(i).Caption = Prozent & "%"
End If
End Sub

Private Sub mnuGreen_Click()

Wert(i).BackColor = QBColor(2)
Select Case i
Case 3
    Multiplikator = "100000"
    lblWert(i).Caption = "10^5"
Case 4
    Prozent = "0.5"
    lblWert(i).Caption = Prozent & "%"
Case Else
    lblWert(i).Caption = "5"
End Select
End Sub

Private Sub mnuGrey_Click()

Wert(i).BackColor = QBColor(8)
lblWert(i).Caption = "8"
End Sub

Private Sub mnuOrange_Click()

Wert(i).BackColor = 33023
If i = 3 Then
    Multiplikator = "1000"
    lblWert(i).Caption = Multiplikator
Else
    lblWert(i).Caption = "3"
End If
End Sub

Private Sub mnuRed_Click()

Wert(i).BackColor = vbRed
Select Case i
Case 3
    Multiplikator = "100"
    lblWert(i).Caption = Multiplikator
Case 4
    Prozent = "2"
    lblWert(i).Caption = Prozent & "%"
Case Else
    lblWert(i).Caption = "2"
End Select
End Sub

Private Sub mnuSilver_Click()

Wert(i).BackColor = QBColor(7)
If i = 3 Then
    Multiplikator = "0.01"
    lblWert(i).Caption = Multiplikator
Else
    Prozent = "10"
    lblWert(i).Caption = Prozent & "%"
End If
End Sub

Private Sub mnuViolet_Click()

Wert(i).BackColor = QBColor(5)
Select Case i
Case 3
    Multiplikator = "10000000"
    lblWert(i).Caption = "10^7"
Case 4
    Prozent = "0.1"
    lblWert(i).Caption = Prozent & "%"
Case Else
    lblWert(i).Caption = "7"
End Select
End Sub

Private Sub mnuWhite_Click()

Wert(i).BackColor = vbWhite
lblWert(i).Caption = "9"
End Sub

Private Sub mnuYellow_Click()

Wert(i).BackColor = vbYellow
If i = 3 Then
    Multiplikator = "10000"
    lblWert(i).Caption = "10^4"
Else
    lblWert(i).Caption = "4"
End If
End Sub

Private Sub Wert_Click(Index As Integer)
If i = 0 Then
Farbe = Farbe + 1
'Label1.Caption = Farbe
End If
End Sub

Private Sub Wert_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim vRmin As Double
Dim vRmax As Double

i = Index

If Button = 2 Then
        Select Case i
        Case 3
            mnuSilver.Visible = True
            mnuGold.Visible = True
            mnuBlack.Visible = True
            mnuOrange.Visible = True
            mnuYellow.Visible = True
            mnuGrey.Visible = False
            mnuWhite.Visible = False
        Case 4
            mnuSilver.Visible = True
            mnuGold.Visible = True
            mnuBlack.Visible = False
            mnuOrange.Visible = False
            mnuYellow.Visible = False
            mnuGrey.Visible = False
            mnuWhite.Visible = False
        Case Else
            mnuSilver.Visible = False
            mnuGold.Visible = False
            mnuBlack.Visible = True
            mnuOrange.Visible = True
            mnuYellow.Visible = True
            mnuGrey.Visible = True
            mnuWhite.Visible = True
        End Select

        PopupMenu mnuPopUp
End If

If chk5.Value = 1 Then
    Widerstand = (lblWert(0).Caption * 100 + lblWert(1).Caption * 10 + lblWert(2).Caption) * Multiplikator
Else
    Widerstand = (lblWert(0).Caption * 10 + lblWert(1).Caption) * Multiplikator
End If

If i = 3 Then
    R.Visible = True
    Rmin.Visible = True
    Rmax.Visible = True
    'DeltaR.Visible = True
    lblR.Visible = True
    lblRmin.Visible = True
    lblRmax.Visible = True
    'lblDeltaR.Visible = True
    lblEinheit.Visible = True
    lblEinheitmin.Visible = True
    lblEinheitmax.Visible = True
    'lblEinheitR.Visible = True
    lblProzent.Visible = True
End If

If mnuDrossel.Checked = True Then
    R.Caption = "D"
    Rmin.Caption = "Dmin"
    Rmax.Caption = "Dmax"
Else
    R.Caption = "R"
    Rmin.Caption = "Rmin"
    Rmax.Caption = "Rmax"
End If

Select Case Widerstand
    Case Is >= 1000000
        lblR.Caption = Widerstand / 1000000
        lblRmin.Caption = lblRmin.Caption / 1000000
        lblRmax.Caption = lblRmax.Caption / 1000000
        'lblDeltaR.Caption = lblDeltaR.Caption / 1000000
        If mnuDrossel.Checked = True Then
            lblEinheit.Caption = "kH"
            lblEinheitmin.Caption = "kH"
            lblEinheitmax.Caption = "kH"
        Else
            lblEinheit.Caption = "Mohm"
            lblEinheitmin.Caption = "Mohm"
            lblEinheitmax.Caption = "Mohm"
        'lblEinheitR.Caption = "Mohm"
        End If
    Case Is >= 1000
        lblR.Caption = Widerstand / 1000
        lblRmin.Caption = lblRmin.Caption / 1000
        lblRmax.Caption = lblRmax.Caption / 1000
        'lblDeltaR.Caption = lblDeltaR.Caption / 1000
        If mnuDrossel.Checked = True Then
            lblEinheit.Caption = "H"
            lblEinheitmin.Caption = "H"
            lblEinheitmax.Caption = "H"
        Else
            lblEinheit.Caption = "kohm"
            lblEinheitmin.Caption = "kohm"
            lblEinheitmax.Caption = "kohm"
        'lblEinheitR.Caption = "kohm"
        End If
    Case Else
        lblR.Caption = Widerstand
        lblRmin.Caption = lblRmin.Caption
        lblRmax.Caption = lblRmax.Caption
        'lblDeltaR.Caption = lblDeltaR.Caption
        If mnuDrossel.Checked = True Then
            lblEinheit.Caption = "uH"
            lblEinheitmin.Caption = "uH"
            lblEinheitmax.Caption = "uH"
        Else
            lblEinheit.Caption = "ohm"
            lblEinheitmin.Caption = "ohm"
            lblEinheitmax.Caption = "ohm"
        'lblEinheitR.Caption = "ohm"
        End If
End Select

lblProzent.Caption = lblWert(4).Caption
lblRmin.Caption = lblR.Caption * (100 - Prozent) / 100
lblRmax.Caption = lblR.Caption * (100 + Prozent) / 100
lblEinheitmin.Caption = lblEinheit.Caption
lblEinheitmax.Caption = lblEinheit.Caption

'lblDeltaR.Caption = vRmax - vRmin
End Sub

Private Sub Widerstand_Ein()

frmMain.Caption = "Widerstands-Farbcode Berechnung"
mnuDrossel.Checked = False
End Sub

Private Sub Back_Normal()

Shape1.FillColor = 12648447
Shape1.BorderColor = 12648447
Shape2.FillColor = 12648447
Shape2.BorderColor = 12648447
Shape3.FillColor = 12648447
Shape3.BorderColor = 12648447
End Sub

