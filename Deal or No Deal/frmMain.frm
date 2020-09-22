VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0000FF00&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Deal or No Deal - Loading"
   ClientHeight    =   3495
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox brdWin 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   1560
      ScaleHeight     =   3225
      ScaleWidth      =   4185
      TabIndex        =   59
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.Label lblWinnings 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   61
         Top             =   1560
         Width           =   3975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "You Have Won"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   60
         Top             =   960
         Width           =   3975
      End
   End
   Begin VB.PictureBox brdDeal 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   1560
      ScaleHeight     =   3225
      ScaleWidth      =   4185
      TabIndex        =   53
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton btnNoDeal 
         BackColor       =   &H00FF0000&
         Caption         =   "No Deal"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton btnDeal 
         BackColor       =   &H00FF0000&
         Caption         =   "Deal"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Offer"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label lblOffer 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   55
         Top             =   840
         Width           =   3975
      End
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   58
      Text            =   "Pick a case to keep!"
      Top             =   120
      Width           =   4215
   End
   Begin VB.PictureBox brdCase 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   1560
      ScaleHeight     =   3105
      ScaleWidth      =   4185
      TabIndex        =   26
      Top             =   240
      Width           =   4215
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   13
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   14
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   15
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   16
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   17
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   18
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   19
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   20
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   21
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   22
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   23
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   24
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton btnCase 
         BackColor       =   &H0064E1E1&
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   25
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2520
         Width           =   615
      End
   End
   Begin VB.Line Line2 
      X1              =   1560
      X2              =   1560
      Y1              =   240
      Y2              =   3360
   End
   Begin VB.Line Line3 
      X1              =   5760
      X2              =   5760
      Y1              =   240
      Y2              =   3360
   End
   Begin VB.Shape Shape1 
      Height          =   3135
      Left            =   120
      Top             =   240
      Width           =   7095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   25
      X1              =   5760
      X2              =   7200
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   23
      X1              =   5760
      X2              =   7200
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   22
      X1              =   5760
      X2              =   7200
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   21
      X1              =   5760
      X2              =   7200
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   20
      X1              =   5760
      X2              =   7200
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   19
      X1              =   5760
      X2              =   7200
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   18
      X1              =   5760
      X2              =   7200
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   17
      X1              =   5760
      X2              =   7200
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   16
      X1              =   5760
      X2              =   7200
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   15
      X1              =   5760
      X2              =   7200
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   14
      X1              =   5760
      X2              =   7200
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   13
      X1              =   5760
      X2              =   7200
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   11
      X1              =   120
      X2              =   1560
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   10
      X1              =   120
      X2              =   1560
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   9
      X1              =   120
      X2              =   1560
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   8
      X1              =   120
      X2              =   1560
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   7
      X1              =   120
      X2              =   1560
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   6
      X1              =   120
      X2              =   1560
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   5
      X1              =   120
      X2              =   1560
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   4
      X1              =   120
      X2              =   1560
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   120
      X2              =   1560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   120
      X2              =   1560
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   120
      X2              =   1560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   120
      X2              =   1560
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   5760
      TabIndex        =   25
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   5760
      TabIndex        =   24
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   5760
      TabIndex        =   23
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   5760
      TabIndex        =   22
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   5760
      TabIndex        =   21
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   5760
      TabIndex        =   20
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   5760
      TabIndex        =   19
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   5760
      TabIndex        =   18
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   5760
      TabIndex        =   17
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   5760
      TabIndex        =   16
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   5760
      TabIndex        =   15
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   5760
      TabIndex        =   14
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   5760
      TabIndex        =   13
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHow 
         Caption         =   "&How to play"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Suitcase(1 To 26) As Currency
Dim FilledCases As Integer
Dim PrizeMoney(1 To 26) As Currency
Dim Playing As Boolean
Dim CasesToOpen As Integer
Dim Round As Integer
Dim YourCase As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Game Preperation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MakePrizeMoney()
Dim x As Integer
    PrizeMoney(1) = 0.5
    PrizeMoney(2) = 1#
    PrizeMoney(3) = 2#
    PrizeMoney(4) = 5#
    PrizeMoney(5) = 10#
    PrizeMoney(6) = 25#
    PrizeMoney(7) = 50#
    PrizeMoney(8) = 100#
    PrizeMoney(9) = 150#
    PrizeMoney(10) = 250#
    PrizeMoney(11) = 500#
    PrizeMoney(12) = 750#
    PrizeMoney(13) = 1000#
    PrizeMoney(14) = 1500#
    PrizeMoney(15) = 2000#
    PrizeMoney(16) = 2500#
    PrizeMoney(17) = 3000#
    PrizeMoney(18) = 5000#
    PrizeMoney(19) = 7500#
    PrizeMoney(20) = 10000#
    PrizeMoney(21) = 15000#
    PrizeMoney(22) = 25000#
    PrizeMoney(23) = 50000#
    PrizeMoney(24) = 75000#
    PrizeMoney(25) = 100000#
    PrizeMoney(26) = 200000#
    
    For x = 0 To 25
        lblMoney(x).Caption = "   " & Format(PrizeMoney(x + 1), "$#,###,###.#0") & "   "
    Next x
End Sub

Private Sub FillCases()
Dim RandomCase As Integer
Dim x As Integer
FilledCases = 0
MakePrizeMoney

    Do
Randomness:
        Randomize
        RandomCase = Int(Rnd * 26 + 1)
        If Suitcase(RandomCase) = 0 Then
            Suitcase(RandomCase) = PrizeMoney(FilledCases + 1)
            If PrizeMoney(FilledCases + 1) = 0 Then MsgBox "meh", , RandomCase & "  " & FilledCases + 1
            PrizeMoney(FilledCases + 1) = 0
        Else
            DoEvents
            GoTo Randomness
        End If
                    
        FilledCases = FilledCases + 1
        DoEvents
        
    Loop Until (PrizeMoney(26) = 0)
    
    For x = 0 To 25
        lblMoney(x).BackColor = RGB(0, 0, 255)
    Next x
    Me.Caption = "Deal or No Deal"
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Game Play
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    brdDeal.Visible = False
    brdWin.Visible = False
    MakePrizeMoney
    FillCases
    CasesToOpen = 6
    Round = 1
    brdCase.Enabled = True
End Sub

Private Sub btnCase_Click(Index As Integer)
Dim x As Integer
Dim Offer As Long

    If Playing = True Then
        btnCase(Index).Enabled = False
        MsgBox "Case number " & Index + 1 & " contained " _
        & Format(Suitcase(Index + 1), "$#,###,###.#0"), _
        vbInformation, "Opened case " & Index + 1
               
        For x = 0 To 25
            If lblMoney(x).Caption = Format(Suitcase(Index + 1), "   $#,###,###.#0   ") Then
                lblMoney(x).BackColor = RGB(100, 100, 100)
            End If
        Next x
        
        Suitcase(Index + 1) = 0
        FilledCases = FilledCases - 1
        
        If FilledCases = 1 Then
            brdWin.Visible = True
            lblWinnings.Caption = Format(Suitcase(YourCase), "$#,###,###.#0")
            Playing = False
            Exit Sub
        End If
        CasesToOpen = CasesToOpen - 1
    Else
        btnCase(Index).Enabled = False
        btnCase(Index).BackColor = RGB(0, 0, 255)
        YourCase = Index + 1
        MsgBox "You have picked case number " & Index + 1 & _
        " to keep. Good Luck", vbInformation, "Picked case " & Index + 1
        
        Playing = True
    End If
    
    Offer = 0
    For x = 1 To 26
        Offer = Offer + Suitcase(x)
    Next x
    lblOffer.Caption = Format(Offer / FilledCases, "$#,###,##0.00")
        
    txtA.Text = CasesToOpen & " cases to open."
    
    If CasesToOpen = 0 Then
        brdCase.Enabled = False
        brdDeal.Visible = True
    End If
End Sub

Private Sub btnDeal_Click()
    Beep
    brdDeal.Enabled = False
    brdWin.Visible = True
    lblWinnings.Caption = lblOffer.Caption
    Playing = False
End Sub

Private Sub btnNoDeal_Click()
    Beep
    
    Round = Round + 1
    If Round = 2 Then CasesToOpen = 5
    If Round = 3 Then CasesToOpen = 4
    If Round = 4 Then CasesToOpen = 3
    If Round = 5 Then CasesToOpen = 2
    If Round = 6 Then CasesToOpen = 2
    If Round >= 7 Then CasesToOpen = 1

    brdDeal.Visible = False
    brdCase.Enabled = True
    
    txtA.Text = CasesToOpen & " cases to open."
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Menus
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuAbout_Click()
    MsgBox "Deal or No Deal" & vbCrLf & vbCrLf & _
    "Version 1" & vbCrLf & "by Chris Cummerford", vbInformation, _
    "About Deal or No Deal       "
End Sub

Private Sub mnuExit_Click()
Dim Answer As Integer
    Answer = MsgBox("Are you sure you want to exit?", _
    vbExclamation + vbYesNo, "Exit Deal or No Deal?")
    
    If Answer = vbYes Then Unload Me
End Sub

Private Sub mnuHow_Click()
    MsgBox "So you want to know how to play eh?" _
    & vbCrLf & "Well... watch Deal or No Deal weekdays at 5:30pm on 7 to find out!" _
    & vbCrLf & "HA! Its a pretty easy game anyway...", vbInformation, "How to Play"
End Sub

Private Sub mnuNew_Click()
Dim x As Integer
    If Playing = True Then
        x = MsgBox("You are currently playing a game." & vbCrLf & _
        "Are you sure you want to start a new game?", _
        vbExclamation + vbYesNo, "New game?")
        
        If x = vbNo Then Exit Sub
    End If

    brdDeal.Visible = False
    brdWin.Visible = False

    For x = 0 To 25
        Suitcase(x + 1) = 0
        btnCase(x).BackColor = 6611425
        btnCase(x).Enabled = True
    Next x
    
    Playing = False
    FilledCases = 0
    YourCase = 0
    txtA.Text = "Pick a case to keep!"
    
    Form_Load
End Sub
