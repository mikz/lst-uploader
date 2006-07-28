VERSION 5.00
Begin VB.Form Uploader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libimseti Uploader"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picProgress 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   3735
      TabIndex        =   12
      Top             =   2280
      Width           =   3735
   End
   Begin VB.CommandButton alternative 
      Caption         =   "Start!"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox i_heslo 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox i_jmeno 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "uploader.frx":0000
      Left            =   960
      List            =   "uploader.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton b_stop 
      Caption         =   "Exit"
      Height          =   975
      Left            =   2640
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox i_album 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox i_uid 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label status 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   3735
   End
   Begin VB.Label l_login 
      Caption         =   "Heslo:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label l_login 
      Caption         =   "Jméno:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label l_prihlaseni 
      Caption         =   "Pøihlášení:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "id alba:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label l_uid 
      Caption         =   "uid:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Uploader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_cProgress As Collection
Dim id_album, curl, uid, txt As String
Dim FreeImage1 As Long
Dim FreeImage2 As Long
Dim orig_width, orig_height, new_width, new_height As Long
Dim new_scale As Double
Dim fif As FREE_IMAGE_FORMAT
Dim filename_obrazek As String
Dim bOK As Long
Dim b_album As Boolean
Dim b_heslo As Boolean
Dim b_uid As Boolean
Dim b_jmeno As Boolean
 
Public Function AppPath() As String
    
    Dim sAns As String
    sAns = App.path
    If Right(App.path, 1) <> "\" Then sAns = sAns & "\"
    AppPath = sAns

End Function


Public Function adresar_lomitko(adresar As String) As String
 
    If Right(adresar, 1) <> "\" Then adresar_lomitko = adresar & "\"

End Function




Private Sub b_stop_Click()
End
End Sub

Private Sub Form_Paint()
Dim cProgress As cProgressBar
   If Not m_cProgress Is Nothing Then
      If Not cProgress Is Nothing Then
         cProgress.DrawToDC _
            Me.hWnd, Me.hDC, _
            0, Me.ScaleHeight \ Screen.TwipsPerPixelY - 16, _
            Me.ScaleWidth \ Screen.TwipsPerPixelX, _
            Me.ScaleHeight \ Screen.TwipsPerPixelY
      End If
   End If
   End Sub
   

Private Sub Combo1_Click()
Select Case Combo1.ListIndex
Case 1
l_login(0).Visible = True
l_login(1).Visible = True
l_uid.Visible = False
i_uid.Visible = False
i_jmeno.Visible = True
i_heslo.Visible = True
Case 0
l_login(0).Visible = False
l_login(1).Visible = False
l_uid.Visible = True
i_uid.Visible = True
i_jmeno.Visible = False
i_heslo.Visible = False
End Select

End Sub

Private Sub Form_Load()
 Combo1.AddItem ("pomocí uid")
  Combo1.AddItem ("jméno/heslo")

   ' This collection controls animation:
  
 

End Sub

Private Sub l_percent_Click()

End Sub

Private Sub picProgress_Paint()
Dim cProgress As cProgressBar
   If Not m_cProgress Is Nothing Then
      On Error Resume Next
      Set cProgress = m_cProgress("picProgress")
      On Error GoTo 0
      If Not (cProgress Is Nothing) Then
         cProgress.Draw
      End If
   End If
End Sub


Private Sub tmrUpd_Timer()
Dim cProgress As cProgressBar
   If Not m_cProgress Is Nothing Then
      For Each cProgress In m_cProgress
     
         With cProgress
            .Value = .Value + .Tag
            If .ShowText Then
               .Text = CLng(.Percent) & "%"
            End If
            If .Value >= .Max Then
               .Tag = -1 * Abs(.Tag)
            ElseIf .Value < 1 Then
               .Tag = Abs(.Tag)
            End If
            If .DrawObject Is Nothing Then
               Form_Paint
            End If
         End With
      Next
   End If
End Sub
Private Sub cmdStep_Click()
   tmrUpd_Timer
End Sub

Private Sub i_album_Change()
If i_album.Text <> "" Then
b_album = True
Else
b_album = False
End If

If b_album And b_uid Or b_heslo And b_jmeno And b_album Then
alternative.enabled = True
Else
alternative.enabled = False
End If

End Sub

Private Sub i_heslo_Change()
If i_heslo.Text <> "" Then
b_heslo = True
Else
b_heslo = False
End If

If b_album And b_uid Or b_heslo And b_jmeno And b_album Then
alternative.enabled = True
Else
alternative.enabled = False
End If
End Sub

Private Sub i_jmeno_Change()
If i_jmeno.Text <> "" Then
b_jmeno = True
Else
b_jmeno = False
End If
If b_album And b_uid Or b_heslo And b_jmeno And b_album Then
alternative.enabled = True
Else
alternative.enabled = False
End If
End Sub

Private Sub i_uid_Change()
If Len(i_uid.Text) = 32 Then
b_uid = True
Else
b_uid = False
End If
If b_album And b_uid Or b_heslo And b_jmeno And b_album Then
alternative.enabled = True
Else
alternative.enabled = False
End If

End Sub


Private Sub alternative_Click()
Dim sOpen As SelectedFile
Dim Count As Integer
Dim FileList As String
alternative.enabled = False
i_album.enabled = False
i_jmeno.enabled = False
i_heslo.enabled = False
i_uid.enabled = False
Combo1.enabled = False


Select Case Combo1.ListIndex
Case 1
   Dim objLink As HTMLLinkElement
    Dim objMSHTML As New MSHTML.HTMLDocument
    Dim objDocument As MSHTML.HTMLDocument
    
    status.Caption = "Pøihlašuji se na Libimseti.cz"
    
    Set objDocument = objMSHTML.createDocumentFromUrl("http://libimseti.cz/?a=l&e_login=" & Trim$(i_jmeno.Text) & "&e_pass=" & Trim$(i_heslo.Text), vbNullString)
     status.Caption = "Stahuji stránku Libimseti.cz"
    While objDocument.readyState <> "complete"
        DoEvents
        Sleep 100
    Wend
    status.Caption = "Stránka stažena..."
    
    DoEvents
    
    
     Dim RegEx As Object, RegM As Object
 Set RegEx = CreateObject("vbscript.regexp")
 RegEx.Pattern = "uid=(.{32})"
 If RegEx.Test(objDocument.documentElement.outerHTML) Then
  Set RegM = RegEx.Execute(objDocument.documentElement.outerHTML).Item(0)
   uid = Mid(objDocument.documentElement.outerHTML, RegM.FirstIndex + 5, 32)
   RegEx.Pattern = "Nespr.vn. heslo nebo neexistuj.c. p.ezd.vka"
  If RegEx.Test(objDocument.documentElement.outerHTML) Then
  Set RegM = RegEx.Execute(objDocument.documentElement.outerHTML).Item(0)
  status.Caption = "Špatné jméno/heslo"
  alternative.enabled = True
  i_album.enabled = True
i_jmeno.enabled = True
i_heslo.enabled = True
i_uid.enabled = True
Combo1.enabled = True

  Exit Sub
  End If
        status.Caption = "Pøihlášeno..."
   Else
   status.Caption = "Nepodarilo se prihlásit :("
   alternative.enabled = True
     i_album.enabled = True
i_jmeno.enabled = True
i_heslo.enabled = True
i_uid.enabled = True
Combo1.enabled = True
   Exit Sub
   End If
   
  Set RegEx = Nothing
 Set RegM = Nothing
Case 0
uid = Trim$(i_uid.Text)
End Select
id_album = Trim$(i_album.Text)
  'working variables
   Dim cnt As Integer
   Dim tmp As String
   Dim adresar As String
  'dim an array to hold the files selected
   Dim sFileArray() As String
Dim Percent As Double
 
        
  

    FileDialog.sFilter = "Obrázky (*.jpg;*.jpeg;*.bmp)" & Chr$(0) & "*.jpg;*.jpeg;*.bmp" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*"
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
    FileDialog.sDlgTitle = "Show Open"
    FileDialog.sInitDir = App.path & "\"
    sOpen = ShowOpen(Me.hWnd)
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
     Set m_cProgress = New Collection
   
   Dim cProgress As cProgressBar
   
       
   ' Progress in picProgress(1) with segments:
   Set cProgress = New cProgressBar
   cProgress.DrawObject = picProgress
   cProgress.Min = 0
   cProgress.Max = Val(sOpen.nFilesSelected) * 100
   cProgress.Tag = 1
   cProgress.ShowText = True
   cProgress.BackColor = vbButtonFace
   cProgress.BarColor = vbHighlight
   
   m_cProgress.Add cProgress, "picProgress"
     
     If Not m_cProgress Is Nothing Then
      For Each cProgress In m_cProgress
     cProgress.Value = 0
            If cProgress.DrawObject Is Nothing Then
               Form_Paint
            End If
      Next
   End If
    Percent = 100 / Val(sOpen.nFilesSelected)
    
        FileList = "Directory : " & sOpen.sLastDirectory & vbCr
        For Count = 1 To sOpen.nFilesSelected
        adresar = sOpen.sLastDirectory
         If Not m_cProgress Is Nothing Then
             For Each cProgress In m_cProgress
                cProgress.Value = Count * 100
                cProgress.Text = Round(Percent * Count, 1) & " %"
                 If cProgress.DrawObject Is Nothing Then
                    Form_Paint
                 End If
            Next
        End If
         
            
             filename_obrazek = adresar_lomitko(adresar) & sOpen.sFiles(Count)
            fif = FreeImage_GetFileType(filename_obrazek, 0)
            
If fif = FIF_UNKNOWN Then
fif = FreeImage_GetFIFFromFilename(filename_obrazek)
End If

             FreeImage1 = FreeImage_Load(fif, filename_obrazek, 0)
             orig_width = FreeImage_GetWidth(FreeImage1)
             orig_height = FreeImage_GetHeight(FreeImage1)
            If orig_width > Val("460") Or orig_height > Val("345") Then
            'MsgBox ("vetsi nez povoleno")
            If orig_width >= orig_height Then
            'MsgBox ("sirka > vyska")
                 new_scale = Val("460") / Val(orig_width)
                new_width = orig_width * new_scale
                 new_height = orig_height * new_scale
               '  MsgBox (new_scale)
              Else
             ' MsgBox ("sirka < vyska")
                new_scale = Val("345") / Val(orig_height)
                 new_width = orig_width * new_scale
                new_height = orig_height * new_scale
                'MsgBox (new_scale)
             End If
            
            Else
                new_width = orig_width
                new_height = orig_height
            End If
            
            'MsgBox (filename_obrazek & "   : " & orig_width & " x " & orig_height)
            'MsgBox (new_width & " x " & new_height)
             FreeImage2 = FreeImage_Rescale(FreeImage1, new_width, new_height, 2)
             bOK = FreeImage_Save(FIF_JPEG, FreeImage2, AppPath & "temp.jpg", &H80)
             FreeImage_Unload (FreeImage1)
             
             curl = "curl.exe -F " & Chr(34) & "userfile=@" & AppPath & "\temp.jpg" & Chr(34) & " -F " & Chr(34) & "n_a=a" & Chr(34) & " -F " & Chr(34) & "uid=" & uid & Chr(34) & " -F " & Chr(34) & "id_album=" & id_album & Chr(34) & " -F " & Chr(34) & "a=dal" & Chr(34) & " -F " & Chr(34) & "txt=" & sOpen.sFiles(Count) & Chr(34) & " http://libimseti.cz/index.php"
            status.Caption = " Odesílám fotku : " & sOpen.sFiles(Count) & "  (" & Count & " z " & sOpen.nFilesSelected & ")"
           ShellAndLoop (AppPath & curl)
                       
            FileList = FileList & sOpen.sFiles(Count) & vbCr
        Next Count
        status.Caption = "Všechny soubory odeslány ;)"
          
             If Not m_cProgress Is Nothing Then
             For Each cProgress In m_cProgress
                cProgress.Value = cProgress.Max
                cProgress.Text = "100 %"
                 If cProgress.DrawObject Is Nothing Then
                    Form_Paint
                 End If
            Next
        End If
        Call MsgBox(FileList, vbOKOnly + vbInformation, "Show Open Selected")
    End If
       alternative.enabled = True
     i_album.enabled = True
i_jmeno.enabled = True
i_heslo.enabled = True
i_uid.enabled = True
Combo1.enabled = True
Beep
   End Sub
