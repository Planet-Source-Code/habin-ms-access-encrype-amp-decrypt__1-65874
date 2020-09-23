VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAccessDB 
   Caption         =   "Encrypt/Decrypt Access Database"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   8490
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "&Decrypt and Open DB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      ToolTipText     =   "Decrypts DB and DB remains  Decrypted"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "&Encrypt "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Encrypts DB"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "E&nd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   1
      ToolTipText     =   "Ends Program and Encrypts DB"
      Top             =   4080
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5318
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   7515
   End
End
Attribute VB_Name = "frmAccessDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
' Original Program Developed by Habin and can be downloaded from:
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=65874&lngWId=1
'
'Habin's email address:  yellow_river_boy@hotmail.com
'
'Thanks:
'John Cunningham
'Heriberto Mantilla Santamaria

'*************************************************************************************************
                           '**** Form Level Declarations ****

'*************************************************************************************************
'  Be sure to add a Reference to Ms ActiveX Data Objects 2.x Library to Project
'*************************************************************************************************
'
'     ********** Data Grid Option has been selected *********
'     Add a Microsoft DataGrid Control to Project.
'
'********************************************************************************************

Dim DbFile As String                              'Name of DataBase
Dim cn As ADODB.Connection                        'Connect to the ADO Data Type
Dim rs As ADODB.Recordset                         'Record Source Name
Dim SQLstmt As String                             'SQL Statement String(s)
Dim RetVal As Variant                             'MsgBox Return Value

Private Sub cmdDecrypt_Click()

DecryptMDB App.Path & "\MTest1.mdb"

'The following demonstrates the ShellExecute API method.
    ShellExecute Me.hwnd, "Open", App.Path & _
        "\MTest1.mdb", "", "C:\", vbNormalFocus
        
 RetVal = MsgBox("Do you want to exit - DB remains Decrypted?", 36, "MS Access Encrypt/Decrypt Database")
 Select Case RetVal
     Case 6     'Yes
          CloseAll
     Case 7     'No
          'Your Code Goes Here
End Select

End Sub

Private Sub cmdEncrypt_Click()

EncryptMDB App.Path & "\MTest1.mdb"
 
RetVal = MsgBox("Do you want to exit - DB remains Encrypted?", 36, "MS Access Encrypt/Decrypt Database")
Select Case RetVal
     Case 6     'Yes
          CloseAll
     Case 7     'No
          'Your Code Goes Here
End Select

End Sub

Private Sub cmdEnd_Click()

Set rs = Nothing
Close_cn
Set cn = Nothing

 EncryptMDB App.Path & "\MTest1.mdb"
 
 CloseAll
   
End Sub

Private Sub Form_Load()

StartEncDec

DecryptMDB App.Path & "\MTest1.mdb"

Label1.Caption = "Database Name - " & App.Path & "\MTest1.mdb"

     Open_cn
     Set DataGrid1.DataSource = rs
     
End Sub

Private Sub Open_cn()

'     Set the Database Applicable Path
      DbFile = App.Path & "\MTest1.mdb"

'      Establish the Connection
       Set cn = New ADODB.Connection
       cn.CursorLocation = adUseClient
       cn.ConnectionString = _
             "Provider=Microsoft.Jet.OLEDB.4.0;" & _
             "Data Source=" & DbFile & ";" & _
             "Persist Security Info=False"

'      Open the Connection
      cn.Open

'      Once this Connection is opened, it can
'      be used throughout the application

      SQLstmt = "SELECT * FROM [tblCvKVTW/KVXW]"

'      Get the Records
      Set rs = New ADODB.Recordset
      rs.Open SQLstmt, cn, adOpenStatic, adLockOptimistic, _
           adCmdText

End Sub

Private Sub Close_cn()

     cn.Close
     Set cn = Nothing
     
End Sub



