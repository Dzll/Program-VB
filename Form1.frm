VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   ScaleHeight     =   26646.33
   ScaleMode       =   0  'User
   ScaleWidth      =   54706.53
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "Lingkaran"
      Height          =   375
      Left            =   15960
      TabIndex        =   25
      Top             =   11400
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   15960
      TabIndex        =   23
      Top             =   10800
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   13680
      TabIndex        =   20
      Top             =   11400
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   13680
      TabIndex        =   19
      Top             =   10800
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Garis"
      Height          =   375
      Left            =   10440
      TabIndex        =   18
      Top             =   11400
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   8400
      TabIndex        =   13
      Top             =   11400
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   8400
      TabIndex        =   12
      Top             =   10800
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   11400
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   6480
      TabIndex        =   10
      Top             =   10800
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   375
      Left            =   20640
      TabIndex        =   9
      Top             =   11520
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   20640
      TabIndex        =   8
      Top             =   10920
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Titik"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   11400
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Blue"
      Height          =   375
      Left            =   19440
      TabIndex        =   4
      Top             =   11640
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Green"
      Height          =   375
      Left            =   19440
      TabIndex        =   3
      Top             =   10920
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Red"
      Height          =   375
      Left            =   19440
      TabIndex        =   2
      Top             =   10200
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   11400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   10800
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Radius"
      Height          =   255
      Left            =   16200
      TabIndex        =   24
      Top             =   10320
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Y"
      Height          =   255
      Left            =   12480
      TabIndex        =   22
      Top             =   11400
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "X"
      Height          =   255
      Left            =   12480
      TabIndex        =   21
      Top             =   10800
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Titik 2"
      Height          =   255
      Left            =   8640
      TabIndex        =   17
      Top             =   10320
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Titik 1"
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   10320
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Y"
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   11520
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "X"
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   10920
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Y"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   11400
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "X"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   10800
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = "" Or Text2.Text = "" Then
    MsgBox "Data Belum di inputkan "
    Else
        If Option1.Value = True Then
            FillStyle = vbSolid
            FillColor = vbRed
            Circle (Text1, Text2), 50, vbRed
        End If
        If Option2.Value = True Then
            FillStyle = vbSolid
            FillColor = vbGreen
            Circle (Text1, Text2), 50, vbGreen
        End If
        If Option3.Value = True Then
            FillStyle = vbSolid
            FillColor = vbBlue
            Circle (Text1, Text2), 50, vbBlue
        End If
    End If
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    Cls
End Sub

Private Sub Command4_Click()
    If Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
    MsgBox "Data belum di inputkan "
    Else
        If Option1.Value = True Then
            FillStyle = vbSolid
            FillColor = vbRed
            Line (Text3, Text4)-(Text5, Text6), vbRed
        End If
        If Option2.Value = True Then
            FillStyle = vbSolid
            FillColor = vbGreen
            Line (Text3, Text4)-(Text5, Text6), vbGreen
        End If
        If Option3.Value = True Then
            FillStyle = vbSolid
            FillColor = vbBlue
            Line (Text3, Text4)-(Text5, Text6), vbBlue
        End If
    End If
End Sub

Private Sub Command5_Click()
    If Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Then
    MsgBox "Data Belum di inputkan "
    Else
        If Option1.Value = True Then
            FillStyle = vbFSTransparent
            FillColor = vbRed
            Circle (Text7, Text8), Text9, vbRed
        End If
        If Option2.Value = True Then
            FillStyle = vbFSTransparent
            FillColor = vbGreen
            Circle (Text7, Text8), Text9, vbGreen
        End If
        If Option3.Value = True Then
            FillStyle = vbFSTransparent
            FillColor = vbBlue
            Circle (Text7, Text8), Text9, vbBlue
        End If
    End If
End Sub

