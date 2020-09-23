VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frm_Main 
   Caption         =   "(Fun_Font)"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   585
   ClientWidth     =   4620
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   211
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   308
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4260
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm_Main.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuclear 
      Caption         =   "&Clear"
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'i don't know what the purpose of this thing. what i can
'think of is maybe it can be use in chat, instant messaging
'or sylish notepad program. thats why i put it here
'mil_xp@hotmail.com
'29.06.2002

Private Sub colorizeFont()
R = 255 * Rnd
G = 255 * Rnd
B = 255 * Rnd
RichTextBox1.SelColor = RGB(R, G, B)
End Sub

Private Sub Form_Resize()
RichTextBox1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub mnuclear_Click()
RichTextBox1.Text = ""
End Sub

Private Sub RichTextBox1_Change()
DoEvents
colorizeFont
Me.Caption = "(Fun_Font) = Font Color (" & RichTextBox1.SelColor & ")"
DoEvents
End Sub
