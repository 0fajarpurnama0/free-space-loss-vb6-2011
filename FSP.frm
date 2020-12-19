VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "km"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   360
      Picture         =   "FSP.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   4755
      TabIndex        =   31
      Top             =   2160
      Width           =   4815
   End
   Begin VB.TextBox d3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   25
      Text            =   "d"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox f3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2760
      TabIndex        =   24
      Text            =   "f"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox FPS3 
      Height          =   285
      Left            =   3720
      TabIndex        =   23
      Text            =   "FPS"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Calculate3 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   4800
      TabIndex        =   22
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox d2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Text            =   "d"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2760
      TabIndex        =   14
      Text            =   "f"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3720
      TabIndex        =   13
      Text            =   "FPS"
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Calculate2 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   4800
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Calculate1 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox FPS1 
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Text            =   "FPS"
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox f1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Text            =   "f"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox d1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "d"
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "MHz"
      Height          =   255
      Left            =   2880
      TabIndex        =   30
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "miles"
      Height          =   255
      Left            =   1560
      TabIndex        =   29
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label16 
      Caption         =   "dB"
      Height          =   255
      Left            =   4560
      TabIndex        =   28
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label15 
      Caption         =   "+ 20log"
      Height          =   255
      Left            =   2160
      TabIndex        =   27
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   375
      Left            =   3360
      TabIndex        =   26
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "FSP = 36.6 + 20log"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "GHz"
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "km"
      Height          =   255
      Left            =   1560
      TabIndex        =   19
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label10 
      Caption         =   "dB"
      Height          =   255
      Left            =   4560
      TabIndex        =   18
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "+ 20log"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   375
      Left            =   3360
      TabIndex        =   16
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "MHz"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "km"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "FSP = 92.5 + 20log"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      Caption         =   "FSP (Free Space Loss) in Signal Transmission"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "dB"
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "+ 20log"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "FSP = 32.5 + 20log"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calculate1_Click()
FPS1 = 32.5 + (10 * Log(Val(d1))) + (10 * Log(Val(f1)))
End Sub

Private Sub Calculate2_Click()
FPS2 = 92.5 + (10 * Log(Val(d2))) + (10 * Log(Val(f2)))
End Sub

Private Sub Calculate3_Click()
FPS3 = 36.6 + (10 * Log(Val(d3))) + (10 * Log(Val(f3)))
End Sub




Private Sub Load_Click()
Test = Log(1000)
End Sub
