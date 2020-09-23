VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "Form1.frx":0000
      Left            =   810
      List            =   "Form1.frx":000D
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   3465
   End
   Begin VB.TextBox TextTofind 
      Height          =   405
      Left            =   810
      TabIndex        =   0
      Top             =   870
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
    Private Const LB_FINDSTRING = &H18F
    Private Const LB_FINDSTRINGEXACT = &H1A2


Private Sub TextTofind_KeyPress(KeyAscii As Integer)
Dim tempstr As String
 
 If KeyAscii = 8 Then
  Exit Sub
 End If
 
 TextTofind.SelStart = Len(TextTofind) + 1

 If KeyAscii = 13 Then
        Add2AutoComplet TextTofind.Text
        TextTofind.Text = ""
 Else
 
        tempstr = list_find_mutch(TextTofind.Text, KeyAscii)
        
        If Len(tempstr) <> 0 Then
          KeyAscii = 0
          TextTofind.Text = tempstr
          TextTofind.SelStart = 0
          TextTofind.SelLength = Len(tempstr) + 1
        End If
 
End If
End Sub

Private Sub Add2AutoComplet(str As String)
    If Find_Exact_Mutch(str) = False Then
         List1.AddItem str
    End If
End Sub


Private Function Find_Exact_Mutch(str As String) As Boolean
    Dim i As Long
    Dim strTemp As String
    
    If Len(str) = 0 Then
        list_KeyPress = ""
        Find_Exact_Mutch = False
    Else
        strTemp = str
    End If
    
    i = SendMessage(List1.hwnd, LB_FINDSTRINGEXACT, -1, ByVal strTemp)
    
    If i <> -1 Then
        Find_Exact_Mutch = True
    End If
End Function

Private Function list_find_mutch(str As String, KeyAscii As Integer) As String
    Dim i As Long
    Dim strTemp As String
    
    If Len(str) = 0 Then
        list_find_mutch = ""
        Exit Function
    Else
        strTemp = str & Chr(KeyAscii)
    End If
    
    i = SendMessage(List1.hwnd, LB_FINDSTRING, -1, ByVal strTemp)
    
    If i <> -1 Then
        list_find_mutch = List1.List(i)
    End If

End Function
