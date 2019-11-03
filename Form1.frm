VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Easy Color Code"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   FillColor       =   &H00C0FFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "?"
      Height          =   255
      Left            =   5880
      TabIndex        =   14
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtTol 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtVal 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox cmbTol 
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Text            =   "Select..."
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbMul 
      Height          =   315
      Left            =   3120
      TabIndex        =   2
      Text            =   "Select..."
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbSecond 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Text            =   "Select..."
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbFirst 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Text            =   "Select..."
      ToolTipText     =   "First band"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Shape shTol 
      BorderColor     =   &H80000006&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   1
      Left            =   3960
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape shMul 
      BorderColor     =   &H80000006&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   3120
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape shSecond 
      BorderColor     =   &H80000006&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2640
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape shFirst 
      BorderColor     =   &H80000006&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   2160
      Top             =   360
      Width           =   255
   End
   Begin VB.Line Line2 
      X1              =   4560
      X2              =   5280
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   1200
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label6 
      Caption         =   "Tolerance"
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Forth Band"
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Third Band"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Second Band"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "First Band"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Value in ohms"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000E&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1920
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAbout_Click()
frmAbout.Show
'Show about frame


End Sub

Private Sub cmdCalc_Click()
'Select the color of the first band, set its value and change the shape`s fill color accordingly

Select Case cmbFirst.Text
        Case "Black"
            first = 0
            shFirst(0).FillColor = &H0&
        Case "Brown"
            first = 1
            shFirst(0).FillColor = &H404080
        Case "Red"
            first = 2
             shFirst(0).FillColor = &HFF&
        Case "Orange"
            first = 3
            shFirst(0).FillColor = &H80FF&
        Case "Yellow"
            first = 4
            shFirst(0).FillColor = &HFFFF&
        Case "Green"
            first = 5
            shFirst(0).FillColor = &HFF00&
        Case "Blue"
            first = 6
            shFirst(0).FillColor = &HFF0000
        Case "Violet"
            first = 7
            shFirst(0).FillColor = &HC000C0
        Case "Grey"
            first = 8
            shFirst(0).FillColor = &HC0C0C0
        Case "White"
            first = 9
            shFirst(0).FillColor = &HFFFFFF
        Case Else
            first = 0
End Select

'Select the color of the second band, set its value and change the shape`s fill color accordingly
Select Case cmbSecond.Text
        Case "Black"
            sec = 0
            shSecond.FillColor = &H0&
        Case "Brown"
            sec = 1
            shSecond.FillColor = &H404080
        Case "Red"
            sec = 2
            shSecond.FillColor = &HFF&
        Case "Orange"
            sec = 3
            shSecond.FillColor = &H80FF&
        Case "Yellow"
            sec = 4
            shSecond.FillColor = &HFFFF&
        Case "Green"
            sec = 5
            shSecond.FillColor = &HFF00&
        Case "Blue"
            sec = 6
            shSecond.FillColor = &HFF0000
        Case "Violet"
            sec = 7
            shSecond.FillColor = &HC000C0
        Case "Grey"
            sec = 8
            shSecond.FillColor = &HC0C0C0
        Case "White"
            sec = 9
            shFirst(0).FillColor = &HFFFFFF
        Case Else
            sec = 0
End Select

'Select the color of the third band, set its value and change the shape`s fill color accordingly
Select Case cmbMul.Text
        Case "Black"
            expon = 0
            shMul.FillColor = &H0&
        Case "Brown"
            expon = 1
            shMul.FillColor = &H404080
        Case "Red"
            expon = 2
            shMul.FillColor = &HFF&
        Case "Orange"
            expon = 3
            shMul.FillColor = &H80FF&
        Case "Yellow"
            expon = 4
            shMul.FillColor = &HFFFF&
        Case "Green"
            expon = 5
            shMul.FillColor = &HFF00&
        Case "Blue"
            expon = 6
            shMul.FillColor = &HFF0000
        Case "Violet"
            expon = 7
            shMul.FillColor = &HC000C0
        Case "Silver"
            expon = -2
            shMul.FillColor = &HE0E0E0
        Case "Gold"
            expon = -1
            shMul.FillColor = &HC0FFFF
        Case Else
            expon = 0
End Select

'Select the color of the fourth band, set its value and change the shape`s fill color accordingly
Select Case cmbTol.Text
        
        Case "Brown"
            txtTol.Text = "±1%"
            shTol(1).FillColor = &H404080
        Case "Red"
            txtTol.Text = "±2%"
            shTol(1).FillColor = &HFF&
        Case "Green"
            txtTol.Text = "±0.5%"
            shTol(1).FillColor = &HFF00&
        Case "Blue"
            txtTol.Text = "±0.25%"
            shTol(1).FillColor = &HFF0000
        Case "Violet"
            txtTol.Text = "±0.1%"
            shTol(1).FillColor = &HC000C0
        Case "Grey"
            txtTol.Text = "±0.05%"
            shTol(1).FillColor = &HC0C0C0
        Case "Silver"
            txtTol.Text = "±10%"
            shTol(1).FillColor = &HE0E0E0
        Case "Gold"
            txtTol.Text = "±5%"
            shTol(1).FillColor = &HC0FFFF
        Case Else
            txtVal.Text = "Choose a color"
End Select

'Compute the resistors value
value = (10 * first + sec) * (10 ^ expon)

'Display the value
txtVal.Text = value

End Sub

Private Sub cmdExit_Click()
'Quit
End
End Sub



Private Sub Form_Load()
'Initialise variables
Dim first, sec, expon, tol, value As Integer


'Load the strings into the combo boxes

cmbFirst.AddItem "Black"
cmbFirst.AddItem "Brown"
cmbFirst.AddItem "Red"
cmbFirst.AddItem "Orange"
cmbFirst.AddItem "Yellow"
cmbFirst.AddItem "Green"
cmbFirst.AddItem "Blue"
cmbFirst.AddItem "Violet"
cmbFirst.AddItem "Grey"
cmbFirst.AddItem "White"

cmbSecond.AddItem "Black"
cmbSecond.AddItem "Brown"
cmbSecond.AddItem "Red"
cmbSecond.AddItem "Orange"
cmbSecond.AddItem "Yellow"
cmbSecond.AddItem "Green"
cmbSecond.AddItem "Blue"
cmbSecond.AddItem "Violet"
cmbSecond.AddItem "Grey"
cmbSecond.AddItem "White"

cmbMul.AddItem "Silver"
cmbMul.AddItem "Gold"
cmbMul.AddItem "Black"
cmbMul.AddItem "Brown"
cmbMul.AddItem "Red"
cmbMul.AddItem "Orange"
cmbMul.AddItem "Yellow"
cmbMul.AddItem "Green"
cmbMul.AddItem "Blue"
cmbMul.AddItem "Violet"

cmbTol.AddItem "Brown"
cmbTol.AddItem "Red"
cmbTol.AddItem "Green"
cmbTol.AddItem "Blue"
cmbTol.AddItem "Violet"
cmbTol.AddItem "Grey"
cmbTol.AddItem "Silver"
cmbTol.AddItem "Gold"

End Sub

