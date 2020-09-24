VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Unicode"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5820
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   810
      TabIndex        =   0
      Top             =   420
      Width           =   3990
      Begin VB.CommandButton Command1 
         Caption         =   "Click Me"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   390
         TabIndex        =   1
         Top             =   510
         Width           =   3210
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SETTEXT    As Long = &HC

Private Sub Command1_Click()

  'This works for all controls having an hWnd which respond to the WM_SETTEXT message.
  'Unfortunately a Label has no hWnd and a textbox does not respond to WM_SETTEXT.
  'You can find the hexcodes in the character table.
  
  'PS
  'I speak none of the languages; I just picked some characters that looked typical too me.

  Dim UniCode As String

    UniCode = "Greek: " & ChrW$(&H394) & ChrW$(&H3A3) & ChrW$(&H3A8) & ChrW$(&H3A9) & ", " & _
              "Kyrillic: " & ChrW$(&H40A) & ChrW$(&H414) & ChrW$(&H416) & ChrW$(&H42F) & ", " & _
              "Hebrew: " & ChrW$(&H5D0) & ChrW$(&H5D1) & ChrW$(&H5D2) & ChrW$(&H5D3) & ", " & _
              "Arabic: " & ChrW$(&HFDF2)

    DefWindowProc Command1.hWnd, WM_SETTEXT, 0, ByVal StrPtr(UniCode)
    Command1.Refresh

    UniCode = " " & ChrW$(&H394) & ChrW$(&H3A8) & ChrW$(&H3A9) & ChrW$(&H3A3) & " "

    DefWindowProc Frame1.hWnd, WM_SETTEXT, 0, ByVal StrPtr(UniCode)
    Frame1.Refresh

    UniCode = "Hey - it works for the Form Caption" & ChrW$(&H203C) & UniCode & ChrW$(&H25AC) & ChrW$(&H25BA) & ChrW$(&H2665)
    DefWindowProc hWnd, WM_SETTEXT, 0, ByVal StrPtr(UniCode)

End Sub

':) Ulli's VB Code Formatter V2.24.25 (2009-Apr-30 14:03)  Subs: 1  Decl: 4  Code: 28  Total: 32 Lines
':) CommentOnly: 4 (12,5%)  Commented: 0 (0%)  Filled: 21 (65,6%)  Empty: 11 (34,4%)  Max Logic Depth: 1
