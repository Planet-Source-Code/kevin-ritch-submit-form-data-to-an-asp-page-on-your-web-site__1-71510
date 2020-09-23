VERSION 5.00
Begin VB.Form PostDataToYourWebSiteForm 
   BackColor       =   &H00B8DEFA&
   Caption         =   "Using XMLHTTP in VB6 to ""POST"" Data to your ASP Page at GoDaddy.com"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8325
   Icon            =   "PostDataToYourWebSiteForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "PostDataToYourWebSiteForm.frx":0CCA
      Top             =   3240
      Width           =   8055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00B8DEFA&
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton SubmitButton 
         Caption         =   """POST"" DATA TO ASP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   4
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox Company 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1440
         TabIndex        =   3
         Text            =   "This & That Groovy Corporation, Inc."
         Top             =   360
         Width           =   6375
      End
      Begin VB.TextBox Contact 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1440
         TabIndex        =   2
         Text            =   "Michael O'Hara"
         Top             =   840
         Width           =   6375
      End
      Begin VB.TextBox EMail 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1440
         TabIndex        =   1
         Text            =   "Michael.OHara@V8Software.com"
         Top             =   1320
         Width           =   6375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Developed Dec 2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   3
         Left            =   5040
         TabIndex        =   10
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "EMail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is the actual ASP page on my website : SimpleFormResponse.asp  ( For your FREE use )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   7935
   End
End
Attribute VB_Name = "PostDataToYourWebSiteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SubmitButton_Click()
 SubmitButton.Enabled = False
 Screen.MousePointer = 11
 PostURL$ = "http://ACTBrowser.com/SimpleFormResponse.asp"
 PostURL$ = PostURL$ & "?Company=" & AsciiToHex$(Company)
 PostURL$ = PostURL$ & "&Contact=" & AsciiToHex$(Contact)
 PostURL$ = PostURL$ & "&EMail=" & AsciiToHex$(EMail)
 a$ = PostURLSource$(PostURL$)
 Screen.MousePointer = Default
 SubmitButton.Enabled = True
 MsgBox a$, vbApplicationModal + vbInformation, "RESPONSE SENT BACK BY MY GODADDY.COM ASP PAGE!        "
End Sub

Public Function AsciiToHex(TypedData$) As String
'===========================================
' THIS CONVERSION FUNCTION STOPS ISSUES WITH
' AMPERSANDS AND APOSTROPHES IN YOUR DATA.
' ========================================
' IT ALSO FACILITATES MEMO FIELD DATA TOO!
'=========================================
 On Error Resume Next
 tmp$ = ""
 For i = 1 To Len(TypedData$)
  n = Asc(Mid$(TypedData$, i, 1))
  HV$ = "00" & Hex$(n)
  HV$ = Right$(HV$, 2)
  tmp$ = tmp$ & HV$
 Next i
 AsciiToHex$ = tmp$
End Function

Private Sub BuildPostData(BYTEARRAY() As Byte, ByVal strPostData As String)
 Dim intNewBytes As Long
 Dim strCH As String
 Dim i As Long
 intNewBytes = Len(strPostData) - 1
 If intNewBytes < 0 Then
  Exit Sub
 End If
 ReDim BYTEARRAY(intNewBytes)
 For i = 0 To intNewBytes
  strCH = Mid$(strPostData, i + 1, 1)
  BYTEARRAY(i) = Asc(strCH)
 Next
End Sub

Function PostURLSource(TheURL As String) As String
'=======================================================
'"?" Prefix 1st Variable, "&" prefix for subsequent vars
'=======================================================
 s = InStr(TheURL, "?")
 If s = 0 Then
  Exit Function
 End If
 SiteASP$ = Left$(TheURL, s - 1)
 StringtoPost = Right$(TheURL, Len(TheURL) - s)
 Dim bytpostdata() As Byte
 Dim strPostData As String
 Dim strHeader As String
 Dim varPostData As Variant
'====================================
'Pack the post data into a byte array
'====================================
 strPostData = StringtoPost
 BuildPostData bytpostdata(), strPostData
'=============================
'Write the byte into a variant
'=============================
 varPostData = bytpostdata
'=================
'Create the Header
'=================
 strHeader = "application/x-www-form-urlencoded" + Chr(10) + Chr(13)
'=============
'Post the data
'=============
 Dim xmlhttp As Object
 Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
 xmlhttp.Open "POST", SiteASP$, False
 xmlhttp.setRequestHeader "Content-Type", strHeader
 xmlhttp.send varPostData
 HTTPText$ = xmlhttp.responseText
 Set xmlhttp = Nothing
 PostURLSource = HTTPText$
End Function
