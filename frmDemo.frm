VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "Test Alert"
   ClientHeight    =   4665
   ClientLeft      =   3195
   ClientTop       =   2685
   ClientWidth     =   6360
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   6360
   Begin VB.Frame fraGrad 
      Caption         =   "Gradiant Type"
      Height          =   885
      Left            =   2010
      TabIndex        =   16
      Top             =   2100
      Width           =   1740
      Begin VB.OptionButton optGradiantType 
         Caption         =   "Tube"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   18
         Top             =   570
         Width           =   1575
      End
      Begin VB.OptionButton optGradiantType 
         Caption         =   "Corner (Default)"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   270
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame fraAction 
      Caption         =   "Action"
      Height          =   885
      Left            =   270
      TabIndex        =   13
      Top             =   2100
      Width           =   1725
      Begin VB.OptionButton optAction 
         Caption         =   "Fade"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Top             =   525
         Width           =   1485
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Scroll (Default)"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   270
         Value           =   -1  'True
         Width           =   1485
      End
   End
   Begin VB.Frame fraShape 
      Caption         =   "Shape"
      Height          =   915
      Left            =   255
      TabIndex        =   10
      Top             =   3000
      Width           =   1740
      Begin VB.OptionButton optShape 
         Caption         =   "Round corners"
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   12
         Top             =   555
         Width           =   1560
      End
      Begin VB.OptionButton optShape 
         Caption         =   "Square (Default)"
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   270
         Value           =   -1  'True
         Width           =   1560
      End
   End
   Begin VB.Frame fraColorOption 
      Caption         =   "Color"
      Height          =   1500
      Left            =   2010
      TabIndex        =   5
      Top             =   3000
      Width           =   1740
      Begin VB.OptionButton optColor 
         Caption         =   "Yellow (Default)"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   270
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Blue"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   8
         Top             =   550
         Width           =   1560
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Red"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   7
         Top             =   830
         Width           =   1560
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Green"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   6
         Top             =   1110
         Width           =   1560
      End
   End
   Begin VB.CommandButton cmdCodeClose 
      Caption         =   "Code Close Alert"
      Height          =   435
      Left            =   150
      TabIndex        =   4
      Top             =   1470
      Width           =   1635
   End
   Begin VB.CommandButton cmdUserClose 
      Caption         =   "User Close Alert"
      Height          =   435
      Left            =   150
      TabIndex        =   2
      Top             =   795
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   1845
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmDemo.frx":000C
      Top             =   105
      Width           =   4365
   End
   Begin VB.CommandButton cmdAutoClose 
      Caption         =   "Auto Close Alert"
      Height          =   435
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "The first window displayed will resize itself to match the amount of text to display if AutoSize option is not set to False"
      Height          =   1845
      Left            =   3975
      TabIndex        =   3
      Top             =   2325
      Width           =   2115
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_iBackColor As Long

Private Sub cmdAutoClose_Click()
  Dim AlertWindow As frmAlertWindow
  Dim SMessage As String
  
    Set AlertWindow = New frmAlertWindow
    
    SMessage = "Auto Close after 4 seconds" & vbNewLine
    If m_iActiveAlertWindows < 1 Then
        SMessage = SMessage & "AutoSize = True" & vbNewLine
    Else
        SMessage = SMessage & "AutoSize = False" & vbNewLine
    End If

        
    AlertWindow.DisplayMessage SMessage & Text1.Text, 4, _
        optAction(1).Value, , optShape(0).Value, m_iBackColor, optGradiantType(1).Value
        
End Sub

Private Sub cmdUserClose_Click()
  Dim AlertWindow As frmAlertWindow
  Dim SMessage As String
    
    Set AlertWindow = New frmAlertWindow
        
    SMessage = "Close on User Click" & vbNewLine
    If m_iActiveAlertWindows < 1 Then
        SMessage = SMessage & "AutoSize = True" & vbNewLine
    Else
        SMessage = SMessage & "AutoSize = False" & vbNewLine
    End If
    
    AlertWindow.DisplayMessage SMessage & Text1.Text, 0, _
        optAction(1).Value, True, optShape(0).Value, m_iBackColor, optGradiantType(1).Value
End Sub


Private Sub cmdCodeClose_Click()
  Static bShowClose As Boolean
  Dim Frm As Form
  Dim SMessage As String
  
    If Not bShowClose Then
        
        SMessage = "Close By Code Only" & vbNewLine & _
            "AutoSize = False" & vbNewLine
            
        frmAlertWindow.DisplayMessage SMessage & Text1.Text, -1, _
            optAction(1).Value, False, optShape(0).Value, m_iBackColor, optGradiantType(1).Value
        
        bShowClose = True
        cmdCodeClose.Caption = "Click Here to Close"
    Else
        bShowClose = False
        cmdCodeClose.Caption = "Code Close Alert"
        
        '/* This is only necessary if you are allowing more than 1 copy to be
        '/* shown at the same time; else use frmAlertWindow.CloseActivate = True
        For Each Frm In Forms
            If Frm.Name = "frmAlertWindow" Then
                Frm.CloseActivate = True
            End If
        Next Frm
    End If

End Sub


Private Sub Form_Load()
    Text1.Text = "Today is " & Date & vbNewLine & Time
    m_iBackColor = &HC0FFFF
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDemo = Nothing
End Sub


Private Sub optColor_Click(Index As Integer)
    Select Case Index
    Case 0 '/* Yellow
        m_iBackColor = &HC0FFFF
    Case 1 '/* Blue
        m_iBackColor = RGB(160, 195, 255)
    Case 2 '/* Red
        m_iBackColor = RGB(255, 200, 200)
    Case 3 '/* Green
        m_iBackColor = RGB(200, 255, 200)
    End Select

End Sub

