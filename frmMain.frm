VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6240
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   8070
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8070
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy to clipboard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5940
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4950
      Width           =   1995
   End
   Begin VB.PictureBox picProgressBar 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   90
      ScaleHeight     =   315
      ScaleWidth      =   7785
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4365
      Width           =   7845
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "frmMain.frx":0442
      Top             =   3510
      Width           =   7845
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "frmMain.frx":044E
      Top             =   1125
      Width           =   7845
   End
   Begin VB.TextBox txtExpected 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "frmMain.frx":0457
      Top             =   2475
      Width           =   7845
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Index           =   2
      Left            =   6630
      Picture         =   "frmMain.frx":0463
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Display About Screen"
      Top             =   5490
      Width           =   615
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Index           =   3
      Left            =   7320
      Picture         =   "frmMain.frx":076D
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Terminate this application"
      Top             =   5490
      Width           =   615
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Index           =   1
      Left            =   5940
      Picture         =   "frmMain.frx":0A77
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Stop the active process"
      Top             =   5490
      Width           =   615
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Index           =   0
      Left            =   5940
      Picture         =   "frmMain.frx":0EB9
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Start the wiping process"
      Top             =   5490
      Width           =   615
   End
   Begin VB.ComboBox cboHashType 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   105
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   5745
      Width           =   2130
   End
   Begin VB.ComboBox cboInput 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmMain.frx":11C3
      Left            =   105
      List            =   "frmMain.frx":11C5
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   5070
      Width           =   2130
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "frmMain.frx":11C7
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   180
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   7320
      Picture         =   "frmMain.frx":1609
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   150
      TabIndex        =   24
      Top             =   2160
      Width           =   7755
   End
   Begin VB.Label lblAES_Msg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SHA2 hashing not available"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2445
      TabIndex        =   22
      Top             =   5085
      Width           =   3285
   End
   Begin VB.Label lblAppTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CryptoAPI Hash"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2415
      TabIndex        =   18
      Top             =   90
      Width           =   3225
   End
   Begin VB.Label lblDisclaimer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblDisclaimer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   2880
      TabIndex        =   17
      Top             =   5490
      Width           =   2640
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Select hash algorithm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   150
      TabIndex        =   16
      Top             =   5520
      Width           =   2070
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Input data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   195
      TabIndex        =   15
      Top             =   4800
      Width           =   1740
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Expected output"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   150
      TabIndex        =   14
      Top             =   1935
      Width           =   1665
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Actual output "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   150
      TabIndex        =   13
      Top             =   3255
      Width           =   1410
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Data input"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   150
      TabIndex        =   12
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kenneth Ives"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   3495
      TabIndex        =   11
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblTitle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3690
      TabIndex        =   10
      Top             =   540
      Width           =   675
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data Length "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   5085
      TabIndex        =   9
      Top             =   840
      Width           =   2730
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Project:       Test hash routines
'
' REFERENCE:
'
' NIST (National Institute of Standards and Technology) Publications
' (FIPS, Special Publications)
' http://csrc.nist.gov/publications/PubsFIPS.html
'
' FIPS 180-2 (Federal Information Processing Standards Publication)
' dated 1-Aug-2002, with Change Notice 1, dated 25-Feb-2004
' http://csrc.nist.gov/publications/fips/fips180-2/FIPS180-2_changenotice.pdf
'
' FIPS 180-3 (Federal Information Processing Standards
' Publication) dated Oct-2008 (supercedes FIPS 180-2)
' http://csrc.nist.gov/publications/fips/fips180-3/fips180-3_final.pdf
'
' FIPS 180-4 (Federal Information Processing Standards Publication)
' dated Mar-2012 (Supercedes FIPS-180-3)
' http://csrc.nist.gov/publications/fips/fips180-4/fips-180-4.pdf
'
' Examples of the implementation of the secure hash algorithms
' SHA-1, SHA-224, SHA-256, SHA-384, SHA-512 can be found at:
' http://csrc.nist.gov/groups/ST/toolkit/examples.html
' http://csrc.nist.gov/groups/ST/toolkit/documents/Examples/SHA2_Additional.pdf
'
' Aaron Gifford 's additional test vectors
' http://www.adg.us/computers/sha.html
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Dec-2006  Kenneth Ives  kenaso@tx.rr.com
' 05-Apr-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added progressbar display
' 01-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for the speed enhancement.
'              This way the progress bar is only initialized once.
'              See ProgressBar() routine.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const ONE_MIL          As Long = 1000000           ' 1,000,000
  ' Constants - for Hyperlinks
  Private Const SW_SHOWNORMAL    As Long = 1

' ***************************************************************************
' API declares
' ***************************************************************************
  ' API declares - for Hyperlinks
  '
  ' The ShellExecute function opens or prints a specified file.  The file
  ' can be an executable file or a document file.
  Private Declare Function ShellExecute Lib "shell32.dll" _
          Alias "ShellExecuteA" (ByVal hwnd As Long, _
          ByVal lpOperation As String, ByVal lpFile As String, _
          ByVal lpParameters As String, ByVal lpDirectory As String, _
          ByVal nShowCmd As Long) As Long

  ' The GetDesktopWindow function returns a handle to the desktop window.
  ' The desktop window covers the entire screen. The desktop window is
  ' the area on top of which other windows are painted.
  Private Declare Function GetDesktopWindow Lib "user32" () As Long

' ***************************************************************************
' Module variables
' ***************************************************************************
  Private mstrPath            As String
  Private mlngAlgorithm       As Long
  Private mlngExpectedResults As Long
  Private mobjKeyEdit         As cKeyEdit

  ' Used to update progressbar display
  Private WithEvents mobjAPI_Hash As cAPI_Hash
Attribute mobjAPI_Hash.VB_VarHelpID = -1

Private Sub cboHashType_Click()
    
    ResetProgressBar
        
    With frmMain
        mlngAlgorithm = .cboHashType.ListIndex     ' hash algorithm desired (0-6)
        .lblTitle(0).Caption = .cboHashType.Text   ' Text from combobox
        
        ' determlne test data
        With .cboInput
             .Clear
             .AddItem "abc"
             .AddItem "Short phrase"
             .AddItem "56 Characters"
             .AddItem "112 Characters"
             .AddItem "1000 Letter 'A'"
             .AddItem "1515 Text file"
             .AddItem "2175 Binary file"
             .AddItem "12,271 Binary file"
             .AddItem "1,000,000 Letter 'a'"
             .AddItem "1,000,000 Binary '0'"
             .ListIndex = 0
        End With
    
       ' cboInput_Click   ' Update display
    End With
                
End Sub

Private Sub cboInput_Click()
  
    Dim strTestData   As String
    Dim strDataLength As String
    Dim strOutput     As String
    
    On Error GoTo cboInput_Click_Error

    ResetProgressBar
    
    With frmMain
        mlngExpectedResults = .cboInput.ListIndex   ' capture test data desired
        .cmdCopy.Enabled = False                    ' disable copy button
        .txtInput.Text = vbNullString                         ' empty input textbox
        .txtExpected.Text = vbNullString                      ' empty expected output textbox
        .txtOutput.Text = vbNullString                        ' empty actual output textbox
        .lblURL.Caption = vbNullString                        ' empty URL label
        .lblURL.Enabled = True                      ' enable URL hyperlink
        .lblTitle(6).Caption = vbNullString                   ' empty test data size display
        
        ' Safety during initial application load
        If mlngExpectedResults < 0 Then
            Exit Sub
        End If
            
        SelectResults mlngAlgorithm, mlngExpectedResults, strTestData, strDataLength, strOutput
        
        ' load input data, input data length, and expected output data to the screen
        .txtInput.Text = strTestData
        .lblTitle(6).Caption = "Data length:  " & Format$(strDataLength, "#,##0")
        .txtExpected.Text = LCase$(strOutput)
          
        Select Case mlngAlgorithm
               
               Case 0         ' MD2
                    Select Case mlngExpectedResults
                           Case 0
                                .lblURL.Caption = "http://tools.ietf.org/html/rfc1319"
                           Case 1
                                .lblURL.Caption = "http://en.wikipedia.org/wiki/MD2_(cryptography)"
                           Case 2
                                .lblURL.Caption = "http://home2.paulschou.net/tools/xlate/"
                           Case Else
                                .lblURL.Enabled = False   ' Disable hyperlink label
                    End Select
                    
               Case 1         ' MD4
                    Select Case mlngExpectedResults
                           Case 0, 2
                                .lblURL.Caption = "http://www.faqs.org/rfcs/rfc1320.html"
                           Case 1
                                .lblURL.Caption = "http://en.wikipedia.org/wiki/MD4"
                           Case Else
                                .lblURL.Enabled = False   ' Disable hyperlink label
                    End Select
                    
               Case 2         ' MD5
                    Select Case mlngExpectedResults
                           Case 0, 2
                                .lblURL.Caption = "http://www.faqs.org/rfcs/rfc1321.html"
                           Case 1
                                .lblURL.Caption = "http://en.wikipedia.org/wiki/MD5"
                           Case Else
                                .lblURL.Enabled = False   ' Disable hyperlink label
                    End Select
                    
               Case 3         ' SHA-1
                    Select Case mlngExpectedResults
                           Case 0, 2
                                .lblURL.Caption = "http://csrc.nist.gov/groups/ST/toolkit/documents/Examples/SHA1.pdf"
                           Case 1
                                .lblURL.Caption = "http://en.wikipedia.org/wiki/Examples_of_SHA_digests"
                           Case 8
                                .lblURL.Caption = "http://csrc.nist.gov/publications/fips/fips180-2/FIPS180-2_changenotice.pdf"
                           Case Else
                                .lblURL.Enabled = False   ' Disable hyperlink label
                    End Select
                    
               Case 4      ' SHA-256
                    Select Case mlngExpectedResults
                           Case 0, 2
                                .lblURL.Caption = "http://csrc.nist.gov/groups/ST/toolkit/documents/Examples/SHA256.pdf"
                           Case 1
                                .lblURL.Caption = "http://en.wikipedia.org/wiki/Examples_of_SHA_digests"
                           Case 4, 9
                                .lblURL.Caption = "http://csrc.nist.gov/groups/ST/toolkit/documents/Examples/SHA2_Additional.pdf"
                           Case 5 To 7
                                .lblURL.Caption = "http://www.aarongifford.com/computers/sha.html"
                           Case 8
                                .lblURL.Caption = "http://csrc.nist.gov/publications/fips/fips180-2/FIPS180-2_changenotice.pdf"
                           Case Else
                                .lblURL.Enabled = False   ' Disable hyperlink label
                    End Select
                    
               Case 5      ' SHA-384
                    Select Case mlngExpectedResults
                           Case 0, 3
                                .lblURL.Caption = "http://csrc.nist.gov/groups/ST/toolkit/documents/Examples/SHA384.pdf"
                           Case 1
                                .lblURL.Caption = "http://en.wikipedia.org/wiki/Examples_of_SHA_digests"
                           Case 4, 9
                                .lblURL.Caption = "http://csrc.nist.gov/groups/ST/toolkit/documents/Examples/SHA2_Additional.pdf"
                           Case 5 To 7
                                .lblURL.Caption = "http://www.aarongifford.com/computers/sha.html"
                           Case 8
                                .lblURL.Caption = "http://csrc.nist.gov/publications/fips/fips180-2/FIPS180-2_changenotice.pdf"
                           Case Else
                                .lblURL.Enabled = False   ' Disable hyperlink label
                    End Select
                    
               Case 6      ' SHA-512
                    Select Case mlngExpectedResults
                           Case 0, 3
                                .lblURL.Caption = "http://csrc.nist.gov/groups/ST/toolkit/documents/Examples/SHA512.pdf"
                           Case 1
                                .lblURL.Caption = "http://en.wikipedia.org/wiki/Examples_of_SHA_digests"
                           Case 5 To 7
                                .lblURL.Caption = "http://www.aarongifford.com/computers/sha.html"
                           Case 4, 9
                                .lblURL.Caption = "http://csrc.nist.gov/groups/ST/toolkit/documents/Examples/SHA2_Additional.pdf"
                           Case 8
                                .lblURL.Caption = "http://csrc.nist.gov/publications/fips/fips180-2/FIPS180-2_changenotice.pdf"
                           Case Else
                                .lblURL.Enabled = False   ' Disable hyperlink label
                    End Select
               
               Case Else
                   .lblURL.Enabled = False   ' Disable hyperlink label
        End Select
    End With
    
cboInput_Click_CleanUp:
    On Error GoTo 0
    Exit Sub

cboInput_Click_Error:
    Err.Clear
    Resume cboInput_Click_CleanUp
    
End Sub

Private Sub cmdChoice_Click(Index As Integer)

    Dim hFile      As Long
    Dim lngIndex   As Long
    Dim strSource  As String
    Dim abytData() As Byte
    
    Erase abytData()  ' Always start with an empty array
    
    Select Case Index
           
           Case 0 ' Start processing
                gblnStopProcessing = False
                mobjAPI_Hash.StopProcessing = False
                ResetProgressBar
                
                DoEvents
                Screen.MousePointer = vbHourglass   ' Change mouse pointer
                frmMain.txtOutput.Text = vbNullString         ' Empty output textbox
                ResetCmdButtons                     ' Display command buttons correctly
                
                ' Set API_Hash property values
                With mobjAPI_Hash
                    .HashMethod = mlngAlgorithm     ' Selected hash method
                    .HashRounds = 1                 ' Number of hashing iterations
                    .ReturnHexString = True         ' Return hashed data in hex format
                    .ReturnLowercase = True        ' Return in uppercase characters
                End With
                
                With frmMain
                    Select Case mlngExpectedResults
                    
                           Case 0 To 3  ' various test strings
                                ' Convert string data to byte array
                                abytData() = StringToByteArray(Trim$(.txtInput.Text))
                                
                                ' Perform hash string function
                                abytData() = mobjAPI_Hash.HashString(abytData())
                           
                                ' Display hashed results
                                .txtOutput.Text = ByteArrayToString(abytData())
                           
                           Case 4  ' 1000 letter 'A' (0x41)
                                ' Convert string data to byte array
                                abytData() = StringToByteArray(String$(1000, 65))
                                
                                ' Perform hash string function
                                abytData() = mobjAPI_Hash.HashString(abytData())
                           
                                ' Display hashed results
                                .txtOutput.Text = ByteArrayToString(abytData())
                           
                           Case 5   ' Excert from President Abraham Lincoln
                                ' format path\filename to be hashed
                                strSource = mstrPath & "TestFiles\" & TEST_FILE1
    
                                ' Convert path\filename string to byte array
                                abytData() = StringToByteArray(strSource)
                                
                                ' Perform hash file function.
                                abytData() = mobjAPI_Hash.HashFile(abytData())
    
                                ' Display hashed results
                                .txtOutput.Text = ByteArrayToString(abytData())
                           
                           Case 6   ' Binary test file
                                ' format path\filename to be hashed
                                strSource = mstrPath & "TestFiles\" & TEST_FILE2
    
                                ' Convert path\filename string to byte array
                                abytData() = StringToByteArray(strSource)
                                
                                ' Perform hash file function.
                                abytData() = mobjAPI_Hash.HashFile(abytData())
    
                                ' Display hashed results
                                .txtOutput.Text = ByteArrayToString(abytData())
    
                           Case 7   ' Binary test file
                                ' format path\filename to be hashed
                                strSource = mstrPath & "TestFiles\" & TEST_FILE3
    
                                ' Convert path\filename string to byte array
                                abytData() = StringToByteArray(strSource)
                                
                                ' Perform hash file function.
                                abytData() = mobjAPI_Hash.HashFile(abytData())
    
                                ' Display hashed results
                                .txtOutput.Text = ByteArrayToString(abytData())
    
                           Case 8   ' 1,000,000 repetitions of the letter 'a'
                                ReDim abytData(ONE_MIL - 1)   ' Adjust array to exact size
                                
                                ' format path\filename to be hashed
                                DoEvents
                                strSource = mstrPath & TEST_FILE4
                                
                                ' Build an array of data
                                For lngIndex = 0 To ONE_MIL - 1
                                    abytData(lngIndex) = &H61  ' Lowercase 'a' in hex
                                Next lngIndex
                                
                                ' Verify test file is empty
                                hFile = FreeFile
                                Open strSource For Output As #hFile
                                Close #hFile
                                
                                ' Load test file with appropriate data
                                hFile = FreeFile
                                Open strSource For Binary Access Write As #hFile
                                Put #hFile, , abytData()     ' Write data to file
                                Close #hFile                 ' Close test file
                                
                                DoEvents
                                Erase abytData()                           ' Empty array
                                abytData() = StringToByteArray(strSource)  ' Convert path\filename string to byte array
                                
                                ' Perform hash file function.
                                abytData() = mobjAPI_Hash.HashFile(abytData())
    
                                ' Display hashed results
                                .txtOutput.Text = ByteArrayToString(abytData())
                    
                                ' Empty test file.  No longer needed.
                                hFile = FreeFile
                                Open strSource For Output As #hFile
                                Close #hFile
                                
                                ' Remove source file
                                On Error Resume Next
                                Kill strSource
                                On Error GoTo 0
                    
                           Case 9   ' 1,000,000 repetitions of binary zeroes
                                ReDim abytData(ONE_MIL - 1)   ' Adjust array to exact size
                                
                                ' format path\filename to be hashed
                                DoEvents
                                strSource = mstrPath & TEST_FILE4
                                
                                ' Build an array of data
                                For lngIndex = 0 To ONE_MIL - 1
                                    abytData(lngIndex) = &H0   ' Binary zero in hex
                                Next lngIndex
                                       
                                ' Verify test file is empty
                                hFile = FreeFile
                                Open strSource For Output As #hFile
                                Close #hFile
                                
                                ' Load test file with appropriate data
                                hFile = FreeFile
                                Open strSource For Binary Access Write As #hFile
                                Put #hFile, , abytData()     ' Write data to file
                                Close #hFile                 ' Close test file
                                
                                DoEvents
                                Erase abytData()                           ' Empty array
                                abytData() = StringToByteArray(strSource)  ' Convert path\filename string to byte array
                                
                                ' Perform hash file function.
                                abytData() = mobjAPI_Hash.HashFile(abytData())
    
                                ' Display hashed results
                                .txtOutput.Text = ByteArrayToString(abytData())
                    
                                ' Empty test file.  No longer needed.
                                hFile = FreeFile
                                Open strSource For Output As #hFile
                                Close #hFile
                                
                                ' Remove source file
                                On Error Resume Next
                                Kill strSource
                                On Error GoTo 0
                    End Select
                End With
                
                ' Compare expected output with actual
                ' output. Results should be the same.
                If StrComp(txtExpected.Text, txtOutput.Text, vbBinaryCompare) = 0 Then
                    cmdCopy.Enabled = True   ' Enable copy to clipboard button
                Else
                    cboInput_Click
                    
                    ' See if used opted to stop processing
                    If gblnStopProcessing Then
                        InfoMsg "User cancelled processing."
                    Else
                        InfoMsg "Expected results do not match the actual results." & _
                                 vbNewLine & "Did you make any changes to the code?"
                    End If
                End If
                       
                DoEvents
                SetupCmdButtons                  ' Reset command buttons
                Erase abytData()                 ' Always empty array when not needed
                Screen.MousePointer = vbDefault  ' Reset mouse pointer
                
           Case 1    ' Stop processing
                gblnStopProcessing = True
                mobjAPI_Hash.StopProcessing = True
                    
                DoEvents
                SetupCmdButtons               ' Reset command buttons
                Erase abytData()              ' Always empty array when not needed
                cboInput_Click
                Screen.MousePointer = vbDefault
                
           Case 2   ' Show About screen
                frmMain.Hide
                frmAbout.DisplayAbout
       
           Case Else ' Termlnate application
                Erase abytData()                    ' Always empty array when not needed
                gblnStopProcessing = True           ' Set flag to stop processing
                mobjAPI_Hash.StopProcessing = True  ' Verify all class processing is stopped
                Screen.MousePointer = vbDefault     ' Reset mouse pointer
                TerminateProgram                    ' Terminate this application
    End Select
    
End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetText txtOutput.Text
End Sub

Private Sub Form_Load()
      
    Set mobjKeyEdit = New cKeyEdit
    Set mobjAPI_Hash = New cAPI_Hash
    
    gblnStopProcessing = False           ' Preset processing flag to FALSE
    mobjAPI_Hash.StopProcessing = False  ' Preset processing flag for class object
    mstrPath = QualifyPath(App.Path)     ' Append backslash to application path
  
    With frmMain
         
        .Caption = gstrVersion
        mobjKeyEdit.CenterCaption frmMain   ' Center form window caption
        .txtOutput.Text = vbNullString
        .lblDisclaimer.Caption = "This is a freeware product." & vbNewLine & _
                                 "No warranties or guarantees implied or intended."
        
        ' load combo box
        If mobjAPI_Hash.AES_Ready Then
        
            ' Advanced hashing is available
            With .cboHashType
                .Clear
                .AddItem "MD2 (32-bit)"
                .AddItem "MD4 (32-bit)"
                .AddItem "MD5 (32-bit)"
                .AddItem "SHA-1 (32-bit)"
                .AddItem "SHA-256 (32-bit)"
                .AddItem "SHA-384 (64-bit)"
                .AddItem "SHA-512 (64-bit)"
                .ListIndex = 0
            End With
            
            .lblAES_Msg.Visible = False  ' Hide message at bottom of form
            
        Else
            ' Advanced hashing NOT available
            With .cboHashType
                .Clear
                .AddItem "MD2 (32-bit)"
                .AddItem "MD4 (32-bit)"
                .AddItem "MD5 (32-bit)"
                .AddItem "SHA-1 (32-bit)"
                .ListIndex = 0
            End With
            
            .lblAES_Msg.Visible = True   ' Show message at bottom of form
            
        End If
        
        SetupCmdButtons   ' Display command buttons correctly
        
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2  ' center form on screen
        .Show vbModeless   ' reduce flicker
        .Refresh
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Screen.MousePointer = vbDefault  ' Reset mouse pointer
    Set mobjKeyEdit = Nothing        ' Free class objects from memory
    Set mobjAPI_Hash = Nothing
    
    On Error Resume Next
    ' Remove test file with 1,000,000 characters
    ' if it exists
    If IsPathValid(mstrPath & TEST_FILE4) Then
        Kill mstrPath & TEST_FILE4
    End If
    On Error GoTo 0
    
    ' See if "x" in upper right corner was selected
    If UnloadMode = 0 Then
        TerminateProgram
    End If
    
End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub

Private Sub mobjAPI_Hash_HashProgress(ByVal lngProgress As Long)
    
    ProgressBar picProgressBar, lngProgress
    DoEvents
    
End Sub

Private Sub ResetProgressBar()

    ' Resets progressbar to zero
    ' with all white background
    ProgressBar picProgressBar, 0, vbWhite
    
End Sub

' ***************************************************************************
' Routine:       ProgressBar
'
' Description:   Fill a picturebox as if it were a horizontal progress bar.
'
' Parameters:    objProgBar - name of picture box control
'                lngPercent - Current percentage value
'                lngForeColor - Optional-The progression color. Default = Black.
'                           can use standard VB colors or long Integer
'                           values representing a color.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-NOV-2001  Randy Birch  http://vbnet.mvps.org/index.html
'              Routine created
' 14-FEB-2005  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 01-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for the speed enhancement.
'              This way the progress bar is only initialized once.
' 05-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation
' ***************************************************************************
Private Sub ProgressBar(ByRef objProgBar As PictureBox, _
                        ByVal lngPercent As Long, _
               Optional ByVal lngForeColor As Long = vbBlue)

    Dim strPercent As String
    
    Const MAX_PERCENT As Long = 100
    
    ' Called by ResetProgressBar() routine
    ' to reinitialize progress bar properties.
    ' If forecolor is white then progressbar
    ' is being reset to a starting position.
    If lngForeColor = vbWhite Then
        
        With objProgBar
            .AutoRedraw = True      ' Required to prevent flicker
            .BackColor = &HFFFFFF   ' White
            .DrawMode = 10          ' Not Xor Pen
            .FillStyle = 0          ' Solid fill
            .FontName = "Arial"     ' Name of font
            .FontSize = 11          ' Font point size
            .FontBold = True        ' Font is bold.  Easier to see.
            Exit Sub                ' Exit this routine
        End With
    
    End If
        
    ' If no progress then leave
    If lngPercent < 1 Then
        Exit Sub
    End If
    
    ' Verify flood display has not exceeded 100%
    If lngPercent <= MAX_PERCENT Then

        With objProgBar
        
            ' Error trap in case code attempts to set
            ' scalewidth greater than the max allowable
            If lngPercent > .ScaleWidth Then
                lngPercent = .ScaleWidth
            End If
               
            .Cls                        ' Empty picture box
            .ForeColor = lngForeColor   ' Reset forecolor
         
            ' set picture box ScaleWidth equal to maximum percentage
            .ScaleWidth = MAX_PERCENT
            
            ' format percent into a displayable value (ex: 25%)
            strPercent = Format$(CLng((lngPercent / .ScaleWidth) * 100)) & "%"
            
            ' Calculate X and Y coordinates within
            ' picture box and and center data
            .CurrentX = (.ScaleWidth - .TextWidth(strPercent)) \ 2
            .CurrentY = (.ScaleHeight - .TextHeight(strPercent)) \ 2
                
            objProgBar.Print strPercent   ' print percentage string in picture box
            
            ' Print flood bar up to new percent position in picture box
            objProgBar.Line (0, 0)-(lngPercent, .ScaleHeight), .ForeColor, BF
        
        End With
                
        DoEvents   ' allow flood to complete drawing
    
    End If

End Sub

Private Sub lblURL_Click()

    Dim strURL As String
    
    strURL = Trim$(lblURL.Caption)  ' remove trailing spaces
    
    ' Make hyperlink call
    ShellExecute GetDesktopWindow(), "open", strURL, 0&, 0&, SW_SHOWNORMAL
    
End Sub

Private Sub txtExpected_KeyDown(KeyCode As Integer, Shift As Integer)
    
    ' Protect data from being copied
    mobjKeyEdit.NoCopyText txtExpected, KeyCode, Shift

End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    
    ' Protect data from being copied
    mobjKeyEdit.NoCopyText txtInput, KeyCode, Shift

End Sub

Private Sub ResetCmdButtons()

    With frmMain
        .cboInput.Enabled = False        ' Disable Input combobox
        .cboHashType.Enabled = False     ' Disable hash method combobox
        .cmdChoice(0).Visible = False    ' Hide Go button
        .cmdChoice(0).Enabled = False    ' Disable Go button
        .cmdChoice(1).Enabled = True     ' Enable Stop button
        .cmdChoice(1).Visible = True     ' Show Stop button
        .cmdChoice(2).Enabled = False    ' Disable Help button
    End With
                
End Sub

Private Sub SetupCmdButtons()

    With frmMain
        .cboInput.Enabled = True         ' Enable Input combobox
        .cboHashType.Enabled = True      ' Enable hash method combobox
        .cmdChoice(0).Enabled = True     ' Enable Go button
        .cmdChoice(0).Visible = True     ' Show Go button
        .cmdChoice(1).Visible = False    ' Hide Stop button
        .cmdChoice(1).Enabled = False    ' Disable Stop button
        .cmdChoice(2).Enabled = True     ' Enable Help button
    End With
                
End Sub

