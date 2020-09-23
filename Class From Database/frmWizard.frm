VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create a VB 6 Class from a Microsoft Access Database"
   ClientHeight    =   10935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10665
   ControlBox      =   0   'False
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   729
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   711
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pictStepContainer 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   3765
      Index           =   4
      Left            =   480
      ScaleHeight     =   3765
      ScaleWidth      =   7575
      TabIndex        =   36
      Top             =   6000
      Visible         =   0   'False
      Width           =   7575
      Begin VB.PictureBox pictProgress 
         Height          =   375
         Left            =   840
         ScaleHeight     =   315
         ScaleWidth      =   5955
         TabIndex        =   38
         Top             =   1080
         Width           =   6015
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait. This could take a few minutes."
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   37
         Top             =   600
         Width           =   3060
      End
   End
   Begin VB.PictureBox pictStepContainer 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3765
      Index           =   3
      Left            =   360
      ScaleHeight     =   3765
      ScaleWidth      =   7575
      TabIndex        =   20
      Top             =   5880
      Visible         =   0   'False
      Width           =   7575
      Begin VB.Frame frameFriend 
         Caption         =   "Friend"
         Height          =   735
         Left            =   4080
         TabIndex        =   32
         Top             =   2640
         Width           =   2895
         Begin VB.CheckBox chkFriendGet 
            Caption         =   "Get"
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox chkFriendLet 
            Caption         =   "Let"
            Height          =   195
            Left            =   960
            TabIndex        =   34
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox chkFriendSet 
            Caption         =   "Set"
            Height          =   195
            Left            =   1680
            TabIndex        =   33
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame framePrivate 
         Caption         =   "Private"
         Height          =   735
         Left            =   4080
         TabIndex        =   28
         Top             =   1800
         Width           =   2895
         Begin VB.CheckBox chkPrivateGet 
            Caption         =   "Get"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox chkPrivateLet 
            Caption         =   "Let"
            Height          =   195
            Left            =   960
            TabIndex        =   30
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox chkPrivateSet 
            Caption         =   "Set"
            Height          =   195
            Left            =   1680
            TabIndex        =   29
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Un&select All"
         Height          =   375
         Left            =   720
         TabIndex        =   24
         Tag             =   "0"
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Frame framePublic 
         Caption         =   "Public"
         Height          =   735
         Left            =   4080
         TabIndex        =   23
         Top             =   960
         Width           =   2895
         Begin VB.CheckBox chkPublicSet 
            Caption         =   "Set"
            Height          =   195
            Left            =   1680
            TabIndex        =   27
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox chkPublicLet 
            Caption         =   "Let"
            Height          =   195
            Left            =   960
            TabIndex        =   26
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox chkPublicGet 
            Caption         =   "Get"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   615
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2055
         Left            =   720
         TabIndex        =   22
         Top             =   960
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select the field(s) you wish to use, then click Finish."
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   21
         Top             =   600
         Width           =   3630
      End
   End
   Begin VB.PictureBox pictStepContainer 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3765
      Index           =   2
      Left            =   240
      ScaleHeight     =   3765
      ScaleWidth      =   7575
      TabIndex        =   13
      Top             =   5760
      Visible         =   0   'False
      Width           =   7575
      Begin MSComctlLib.ListView ListView1 
         Height          =   2055
         Left            =   720
         TabIndex        =   18
         Top             =   960
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select the table you wish to create classes for, then click Next."
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   19
         Top             =   600
         Width           =   4440
      End
   End
   Begin VB.PictureBox pictStepContainer 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   3765
      Index           =   1
      Left            =   120
      ScaleHeight     =   3765
      ScaleWidth      =   7575
      TabIndex        =   11
      Top             =   5640
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton cmdGetDBPath 
         Height          =   285
         Left            =   5400
         Picture         =   "frmWizard.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtDBPath 
         Height          =   285
         Left            =   720
         TabIndex        =   15
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Locate the database you wish to use, then click Next."
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   17
         Top             =   600
         Width           =   3810
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Path to database:"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   14
         Top             =   1080
         Width           =   1260
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pictStepContainer 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4725
      Index           =   0
      Left            =   0
      ScaleHeight     =   315
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   503
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   7545
      Begin VB.PictureBox pictWelcome 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   4725
         Left            =   0
         ScaleHeight     =   315
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   165
         TabIndex        =   7
         Top             =   0
         Width           =   2475
      End
      Begin VB.Label lblWelcomeDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "This wizard will create a VB 6 class from a Microsoft Access 2000 or newer database."
         Height          =   3015
         Left            =   2640
         TabIndex        =   10
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label lblWelcomeTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to the VB 6 Class Builder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   9
         Top             =   120
         Width           =   4215
      End
      Begin VB.Label lblWelcomeNext 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To continue, click Next."
         Height          =   195
         Left            =   2640
         TabIndex        =   8
         Top             =   4320
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "Finish"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   5040
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":2AAC
            Key             =   "STEP1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":2C06
            Key             =   "STEP2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":2D60
            Key             =   "STEP3"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frameHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7575
      Begin VB.Image imgHeader 
         Height          =   735
         Left            =   6720
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblHeaderText 
         BackStyle       =   0  'Transparent
         Caption         =   "Wizard step description."
         Height          =   435
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Width           =   5640
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblHeaderTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wizard Step Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   120
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Next >"
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "< Back"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   533.333
      Y1              =   62
      Y2              =   62
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   533.333
      Y1              =   61
      Y2              =   61
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   533.333
      Y1              =   317
      Y2              =   317
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   533.333
      Y1              =   316
      Y2              =   316
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CONNECTION_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=<DATA_SOURCE/>;Persist Security Info=False"

Private Const STEP_WELCOME = 0
Private Const STEP_LOAD_DB = 1
Private Const STEP_PICK_TABLE = 2
Private Const STEP_SET_PROPERTIES = 3
Private Const STEP_CONVERT = 4

Private miSteps As Integer
Private miCurrentStep As Integer
Private mcolStepTitles As New Collection
Private mcolStepDescriptions As New Collection

Private Function DoesImageExist(Key As Variant) As Boolean

On Error GoTo ErrorCode

Dim objTMP As ListImage
    
    DoesImageExist = False
        Set objTMP = Me.ImageList1.ListImages(Key)
    DoesImageExist = True
    
ClearVariables:
    Set objTMP = Nothing

    Exit Function
ErrorCode:
    DoesImageExist = False
    GoTo ClearVariables

End Function


Private Sub LoadSteps()
    
    ' Set Titles
    Set mcolStepTitles = Nothing
    Set mcolStepTitles = New Collection
    
    mcolStepTitles.Add "This is not used", CStr("STEP" & STEP_WELCOME)
    mcolStepTitles.Add "Where is the Microsoft Access Database?", CStr("STEP" & STEP_LOAD_DB)
    mcolStepTitles.Add "What table do you wish to export?", CStr("STEP" & STEP_PICK_TABLE)
    mcolStepTitles.Add "What field(s) do you wish to use?", CStr("STEP" & STEP_SET_PROPERTIES)
    mcolStepTitles.Add "Converting database.", CStr("STEP" & STEP_CONVERT)
    
    
    ' Set Descriptions
    Set mcolStepDescriptions = Nothing
    Set mcolStepDescriptions = New Collection
    
    mcolStepDescriptions.Add "This is not used", CStr("STEP" & STEP_WELCOME)
    mcolStepDescriptions.Add "Locate the Microsoft Access database you wish to use to create your VB 6 class.", CStr("STEP" & STEP_LOAD_DB)
    mcolStepDescriptions.Add "Select the table you wish to create class for.", CStr("STEP" & STEP_PICK_TABLE)
    mcolStepDescriptions.Add "Select the field(s) you wish to export.", CStr("STEP" & STEP_SET_PROPERTIES)
    mcolStepDescriptions.Add "Please wait while the wizard creates your VB class.", CStr("STEP" & STEP_CONVERT)
    

End Sub


Private Sub mClearListView(ByVal oListView As ListView)

On Error GoTo ErrorHandle

    With oListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Description"
        .View = lvwReport
        .Checkboxes = False
        .MultiSelect = False
        .HideSelection = False
        .Arrange = lvwAutoLeft
        .FullRowSelect = True
        .HotTracking = False
        .HoverSelection = False
        .LabelEdit = lvwManual
    End With

ClearVariables:
    Exit Sub
    
ErrorHandle:
    Err.Raise Err.Number, Err.source, Err.Description
    GoTo ClearVariables
            

End Sub

Private Function mGetFile() As String

On Error GoTo ErrorHandle

    With Me.CommonDialog1
        .CancelError = False
        .DialogTitle = "Location of file."
        .Filter = "Microsoft Access (*.mdb)|*.mdb"
        .Flags = cdlOFNFileMustExist Or cdlOFNExtensionDifferent Or cdlOFNHideReadOnly
        .ShowOpen
        If DoesFileExist(.FileName) = True Then
            mGetFile = .FileName
        End If
    End With

ClearVariables:
    Exit Function
    
ErrorHandle:
    Err.Raise Err.Number, Err.source, Err.Description
    GoTo ClearVariables

End Function
Private Function mGetType(ByVal lType As DataTypeEnum) As String

    Select Case lType
        Case DataTypeEnum.adBigInt, DataTypeEnum.adInteger, DataTypeEnum.adSingle, DataTypeEnum.adSmallInt, DataTypeEnum.adTinyInt
            mGetType = "Long"
        
        Case DataTypeEnum.adChar, DataTypeEnum.adLongVarChar, DataTypeEnum.adVarChar, DataTypeEnum.adVarWChar, DataTypeEnum.adWChar
            mGetType = "String"
        
        Case DataTypeEnum.adCurrency, DataTypeEnum.adDecimal, DataTypeEnum.adDouble
            mGetType = "Double'"
        
        Case DataTypeEnum.adDate, DataTypeEnum.adDBDate, DataTypeEnum.adDBTime, DataTypeEnum.adDBTimeStamp
            mGetType = "Date"
        
        Case DataTypeEnum.adBoolean
            mGetType = "Boolean"
                        
        Case Else
            mGetType = "Variant"
            
    End Select

End Function

Private Function mLoadFields() As Boolean

On Error GoTo ErrorHandle

Dim sConnectionString As String
Dim oConnection As ADODB.Connection
Dim oCatalog As ADOX.Catalog
Dim oTable As ADOX.Table
Dim oColumn As ADOX.Column
Dim oListItem As ListItem
        
    sConnectionString = CONNECTION_STRING
        sConnectionString = Replace(sConnectionString, "<DATA_SOURCE/>", Me.txtDBPath.Text)
    
    Set oConnection = New ADODB.Connection
        With oConnection
            .ConnectionString = sConnectionString
            .Open
        End With
    
    Set oCatalog = New ADOX.Catalog
        With oCatalog
            Set .ActiveConnection = oConnection
            If .Tables.Count = 0 Then GoTo ClearVariables
            Set oTable = .Tables(Me.ListView1.SelectedItem.Key)
        End With
    
    mClearListView Me.ListView2
        Me.ListView2.ColumnHeaders.Add , , "Type"
        Me.ListView2.Checkboxes = True
    
    For Each oColumn In oTable.Columns
        Set oListItem = Me.ListView2.ListItems.Add(, oColumn.Name, oColumn.Name)
            oListItem.ListSubItems.Add , , mGetType(oColumn.Type)
            oListItem.Checked = True
            oListItem.Tag = "1|1|0|0|0|0|0|0|0"
    Next
    
    AutoSizeColumns Me.ListView2
    Me.ListView2.ListItems(1).Selected = True
    ListView2_ItemClick Me.ListView2.SelectedItem

    mLoadFields = True
    
ClearVariables:
    Set oListItem = Nothing
    Set oColumn = Nothing
    Set oTable = Nothing
    Set oCatalog = Nothing
    
    If Not oConnection Is Nothing Then
        If oConnection.State = adStateOpen Then oConnection.Close
        Set oConnection = Nothing
    End If
    
    Exit Function
    
ErrorHandle:
    mLoadFields = False
    Err.Raise Err.Number, Err.source, Err.Description
    GoTo ClearVariables
            

End Function
Private Function mLoadTables() As Boolean

On Error GoTo ErrorHandle
    
Dim sConnectionString As String
Dim oConnection As ADODB.Connection
Dim oCatalog As ADOX.Catalog
Dim oTable As ADOX.Table

    If DoesFileExist(Me.txtDBPath.Text) = False Then GoTo ClearVariables
    
    sConnectionString = CONNECTION_STRING
        sConnectionString = Replace(sConnectionString, "<DATA_SOURCE/>", Me.txtDBPath.Text)
    
    Set oConnection = New ADODB.Connection
        With oConnection
            .ConnectionString = sConnectionString
            .Open
        End With
    
    Set oCatalog = New ADOX.Catalog
        With oCatalog
            Set .ActiveConnection = oConnection
            If .Tables.Count = 0 Then GoTo ClearVariables
        End With
    
    mClearListView Me.ListView1
    
    For Each oTable In oCatalog.Tables
        If oTable.Type = "TABLE" Then
            Me.ListView1.ListItems.Add , oTable.Name, oTable.Name
        End If
    Next
    
    AutoSizeColumns Me.ListView1

    mLoadTables = True
    
ClearVariables:
    Set oTable = Nothing
    Set oCatalog = Nothing
    
    If Not oConnection Is Nothing Then
        If oConnection.State = adStateOpen Then oConnection.Close
        Set oConnection = Nothing
    End If
    
    Exit Function
    
ErrorHandle:
    mLoadTables = False
    Err.Raise Err.Number, Err.source, Err.Description
    GoTo ClearVariables
            


End Function

Private Sub mSetField()

Dim sTag As String

    sTag = CStr(Me.chkPublicGet.Value)
    sTag = sTag & "|" & CStr(Me.chkPublicLet.Value)
    sTag = sTag & "|" & CStr(Me.chkPublicSet.Value)
    
    sTag = sTag & "|" & CStr(Me.chkPrivateGet.Value)
    sTag = sTag & "|" & CStr(Me.chkPrivateLet.Value)
    sTag = sTag & "|" & CStr(Me.chkPrivateSet.Value)
    
    sTag = sTag & "|" & CStr(Me.chkFriendGet.Value)
    sTag = sTag & "|" & CStr(Me.chkFriendLet.Value)
    sTag = sTag & "|" & CStr(Me.chkFriendSet.Value)
    
    Me.ListView2.SelectedItem.Tag = sTag

End Sub

Public Sub SetSlides()

Dim pictTMP As PictureBox

    Me.Width = 7635
    Me.Height = 5910
    
    For Each pictTMP In Me.pictStepContainer
        If pictTMP.Index <> 0 Then
            pictTMP.Move 0, 64
            pictTMP.BackColor = Me.BackColor
        End If
    Next
    

End Sub

Public Sub StepGotFocus(Step As Integer)

End Sub

Public Sub StepLostFocus(Step As Integer)

End Sub

Public Property Let WizardSteps(iNewValue As Integer)
    miSteps = iNewValue
End Property

Public Property Get WizardSteps() As Integer
    WizardSteps = miSteps
End Property



Private Sub chkFriendGet_Click()

'    If chkFriendGet.Value = 1 Then
'        Me.chkPublicGet.Value = 0
'        Me.chkPrivateGet.Value = 0
'
'        Me.chkPublicGet.Enabled = False
'        Me.chkPrivateGet.Enabled = False
'    Else
'        Me.chkPublicGet.Enabled = True
'        Me.chkPrivateGet.Enabled = True
'    End If
    
    mSetField

End Sub

Private Sub chkFriendLet_Click()

'    If chkFriendLet.Value = 1 Then
'        Me.chkPublicLet.Value = 0
'        Me.chkPrivateLet.Value = 0
'
'        Me.chkPublicLet.Enabled = False
'        Me.chkPrivateLet.Enabled = False
'    Else
'        Me.chkPublicLet.Enabled = True
'        Me.chkPrivateLet.Enabled = True
'    End If
'
    mSetField


End Sub


Private Sub chkFriendSet_Click()

'    If chkFriendSet.Value = 1 Then
'        Me.chkPublicSet.Value = 0
'        Me.chkPrivateSet.Value = 0
'
'        Me.chkPublicSet.Enabled = False
'        Me.chkPrivateSet.Enabled = False
'    Else
'        Me.chkPublicSet.Enabled = True
'        Me.chkPrivateSet.Enabled = True
'    End If
    
    mSetField

End Sub
Private Sub chkPrivateGet_Click()

    If chkPrivateGet.Value = 1 Then
        Me.chkPublicGet.Value = 0
        'Me.chkFriendGet.Value = 0
    
        Me.chkPublicGet.Enabled = False
        'Me.chkFriendGet.Enabled = False
    Else
        Me.chkPublicGet.Enabled = True
        'Me.chkFriendGet.Enabled = True
    End If

    mSetField

End Sub

Private Sub chkPrivateLet_Click()

    If chkPrivateLet.Value = 1 Then
        Me.chkPublicLet.Value = 0
        'Me.chkFriendLet.Value = 0
    
        Me.chkPublicLet.Enabled = False
        'Me.chkFriendLet.Enabled = False
    Else
        Me.chkPublicLet.Enabled = True
        'Me.chkFriendLet.Enabled = True
    End If

    mSetField
End Sub


Private Sub chkPrivateSet_Click()

    If chkPrivateSet.Value = 1 Then
        Me.chkPublicSet.Value = 0
        'Me.chkFriendSet.Value = 0
    
        Me.chkPublicSet.Enabled = False
        'Me.chkFriendSet.Enabled = False
    Else
        Me.chkPublicSet.Enabled = True
        'Me.chkFriendSet.Enabled = True
    End If

    mSetField
    
End Sub


Private Sub chkPublicGet_Click()

    If Me.chkPublicGet.Value = 1 Then
        Me.chkPrivateGet.Value = 0
        'Me.chkFriendGet.Value = 0
    
        Me.chkPrivateGet.Enabled = False
        'Me.chkFriendGet.Enabled = False
    Else
        Me.chkPrivateGet.Enabled = True
        'Me.chkFriendGet.Enabled = True
    End If
    
    mSetField
    
End Sub

Private Sub chkPublicLet_Click()

    If Me.chkPublicLet.Value = 1 Then
        Me.chkPrivateLet.Value = 0
        'Me.chkFriendLet.Value = 0
    
        Me.chkPrivateLet.Enabled = False
        'Me.chkFriendLet.Enabled = False
    Else
        Me.chkPrivateLet.Enabled = True
        'Me.chkFriendLet.Enabled = True
    End If
    
    mSetField

End Sub


Private Sub chkPublicSet_Click()

    If Me.chkPublicSet.Value = 1 Then
        Me.chkPrivateSet.Value = 0
        'Me.chkFriendSet.Value = 0
    
        Me.chkPrivateSet.Enabled = False
        'Me.chkFriendSet.Enabled = False
    Else
        Me.chkPrivateSet.Enabled = True
        'Me.chkFriendSet.Enabled = True
    End If
    
    mSetField


End Sub

Private Sub cmdCancel_Click()

    If MsgBox("Are you sure you wish to cancel this operation?", vbQuestion + vbYesNo + vbDefaultButton2, "Cancel Operation?") = vbNo Then Exit Sub

    Unload Me

End Sub

Private Sub cmdFinish_Click()
    
On Error GoTo ErrorHandle

Dim sTempFile As String
Dim iFreeFile As Integer
Dim oListItem As ListItem
Dim vSettings As Variant
Dim sType As String
Dim sLine As String
Dim sVarName As String


    Me.pictStepContainer(STEP_CONVERT).Visible = True
    Me.pictStepContainer(STEP_CONVERT).ZOrder 0
    Me.cmdMove(0).Enabled = False
    Me.cmdMove(1).Enabled = False
    Me.cmdFinish.Enabled = False
    
    sTempFile = TempFile("~tmp")
    If DoesFileExist(sTempFile) = True Then Kill sTempFile
    
    iFreeFile = FreeFile
    
    Open sTempFile For Output As #iFreeFile
    
        Print #iFreeFile, "Option Explicit" & vbCrLf & vbCrLf
        
        For Each oListItem In Me.ListView2.ListItems
            If oListItem.Checked = True Then
            
                sLine = ""
                sType = ""
                sType = LCase(oListItem.ListSubItems(1).Text)
                Select Case sType
                    Case "string"
                        sLine = "Private ms" & oListItem.Text & " as String"
                    Case "date"
                        sLine = "Private md" & oListItem.Text & " as Date"
                    Case "long"
                        sLine = "Private ml" & oListItem.Text & " as Long"
                    Case "boolean"
                        sLine = "Private mb" & oListItem.Text & " as Boolean"
                    Case "variant"
                        sLine = "Private mv" & oListItem.Text & " as Variant"
                End Select
                
                Print #iFreeFile, sLine
            
            End If
        Next
        
        Print #iFreeFile, ""
        Print #iFreeFile, ""
        
        For Each oListItem In Me.ListView2.ListItems
            If oListItem.Checked = True Then
                
                vSettings = ""
                sLine = ""
                sVarName = ""
                sType = ""
                
                vSettings = Split(oListItem.Tag, "|")
                sType = LCase(oListItem.ListSubItems(1).Text)
                
                Select Case sType
                    Case "string"
                        sVarName = "ms" & oListItem.Text
                    Case "date"
                        sVarName = "md" & oListItem.Text
                    Case "long"
                        sVarName = "ml" & oListItem.Text
                    Case "boolean"
                        sVarName = "mb" & oListItem.Text
                    Case "variant"
                        sVarName = "mv" & oListItem.Text
                End Select
                
                ' ------------ PUBLIC -----------
                If CLng(vSettings(0)) = 1 Then
                    sLine = sLine & "Public Property Get " & oListItem.Text & "() as " & sType & vbCrLf & ""
                        sLine = sLine & vbTab & oListItem.Text & " = " & sVarName & vbCrLf
                    sLine = sLine & "End Property" & vbCrLf
                End If
                
                If CLng(vSettings(1)) = 1 Then
                    sLine = sLine & "Public Property Let " & oListItem.Text & "(ByVal RHS as " & sType & " )" & vbCrLf
                        sLine = sLine & vbTab & sVarName & " = RHS" & vbCrLf
                    sLine = sLine & "End Property" & vbCrLf
                End If
                
                If CLng(vSettings(2)) = 1 Then
                    sLine = sLine & "Public Property Set " & oListItem.Text & "(ByVal RHS as Variant)" & vbCrLf
                        sLine = sLine & vbTab & "Set " & sVarName & " = RHS" & vbCrLf
                    sLine = sLine & "End Property" & vbCrLf
                End If
                
                ' ------------ PRIVATE -----------
                If CLng(vSettings(3)) = 1 Then
                    sLine = sLine & "Private Property Get " & oListItem.Text & "() as " & sType & vbCrLf & ""
                        sLine = sLine & vbTab & oListItem.Text & " = " & sVarName & vbCrLf
                    sLine = sLine & "End Property" & vbCrLf
                End If
                
                If CLng(vSettings(4)) = 1 Then
                    sLine = sLine & "Private Property Let " & oListItem.Text & "(ByVal RHS as " & sType & " )" & vbCrLf
                        sLine = sLine & vbTab & sVarName & " = RHS" & vbCrLf
                    sLine = sLine & "End Property" & vbCrLf
                End If
                
                If CLng(vSettings(5)) = 1 Then
                    sLine = sLine & "Private Property Set " & oListItem.Text & "(ByVal RHS as Variant)" & vbCrLf
                        sLine = sLine & vbTab & "Set " & sVarName & " = RHS" & vbCrLf
                    sLine = sLine & "End Property" & vbCrLf
                End If
                
                ' ------------ FRIEND -----------
                If CLng(vSettings(6)) = 1 Then
                    sLine = sLine & "Friend Property Get friend_" & oListItem.Text & "() as " & sType & vbCrLf & ""
                        sLine = sLine & vbTab & oListItem.Text & " = " & sVarName & vbCrLf
                    sLine = sLine & "End Property" & vbCrLf
                End If
                
                If CLng(vSettings(7)) = 1 Then
                    sLine = sLine & "Friend Property Let friend_" & oListItem.Text & "(ByVal RHS as " & sType & " )" & vbCrLf
                        sLine = sLine & vbTab & sVarName & " = RHS" & vbCrLf
                    sLine = sLine & "End Property" & vbCrLf
                End If
                
                If CLng(vSettings(8)) = 1 Then
                    sLine = sLine & "Friend Property Set friend_" & oListItem.Text & "(ByVal RHS as Variant)" & vbCrLf
                        sLine = sLine & vbTab & "Set " & sVarName & " = RHS" & vbCrLf
                    sLine = sLine & "End Property" & vbCrLf
                End If
                
                
                Print #iFreeFile, sLine
                
            End If
        Next
        
    Close #iFreeFile
    
    Shell "notepad.exe " & sTempFile, vbMaximizedFocus
    Kill sTempFile

ClearVariables:
    Set oListItem = Nothing
    End
    
ErrorHandle:
    Err.Raise Err.Number, Err.source, Err.Description
    GoTo ClearVariables
            
End Sub
Private Sub cmdGetDBPath_Click()

On Error GoTo ErrorHandle

Dim sPath As String

    sPath = mGetFile
    If DoesFileExist(sPath) = True Then
        Me.txtDBPath.Text = sPath
    End If

ClearVariables:
    Exit Sub
    
ErrorHandle:
    ' err.raise err.number, err.source, err.description
    GoTo ClearVariables
            

End Sub

Private Sub cmdMove_Click(Index As Integer)

Dim iNext As Integer
Dim iOldStep As Integer

    iNext = Me.CurrentStep
    iOldStep = iNext

    Select Case Index
    
        Case 0  ' back
            iNext = iNext - 1
            
        Case 1  ' forward
            iNext = iNext + 1
            
    End Select
    
    
    Select Case iNext
    
        Case STEP_WELCOME
        Case STEP_LOAD_DB
            'MsgBox "1"
            
        Case STEP_PICK_TABLE
            If mLoadTables = False Then Exit Sub
            
        Case STEP_SET_PROPERTIES
            If mLoadFields = False Then Exit Sub
        
        
    End Select
    
    
    StepLostFocus iOldStep
        Me.CurrentStep = iNext
    StepGotFocus iNext

End Sub

Private Sub cmdSelect_Click()

Dim oListItem As ListItem

    For Each oListItem In Me.ListView2.ListItems
        oListItem.Checked = CBool(Me.cmdSelect.Tag)
    Next
    
    Me.cmdSelect.Tag = CLng(Not CBool(Me.cmdSelect.Tag))
    If CBool(Me.cmdSelect.Tag) = True Then
        Me.cmdSelect.Caption = "&Select All"
    Else
        Me.cmdSelect.Caption = "U&select All"
    End If

End Sub

Private Sub Form_Load()

    'MsgBox "Don't forget to set the wizard properties"
    Me.WizardSteps = 3
    Me.CurrentStep = 0
    
    LoadSteps
    SetSlides

End Sub



Public Property Get CurrentStep() As Integer
    CurrentStep = miCurrentStep
End Property

Public Property Let CurrentStep(ByVal vNewValue As Integer)

    If vNewValue < 0 Then vNewValue = 0
    If vNewValue > miSteps Then vNewValue = miSteps
    
    
    If vNewValue = 0 Then Me.cmdMove(0).Enabled = False Else Me.cmdMove(0).Enabled = True
    If vNewValue = miSteps Then Me.cmdMove(1).Enabled = False Else Me.cmdMove(1).Enabled = True
    
    If vNewValue <> 0 Then
        Me.pictStepContainer(0).Visible = False
        Me.lblHeaderTitle = mcolStepTitles(CStr("STEP" & vNewValue))
        Me.lblHeaderText = mcolStepDescriptions(CStr("STEP" & vNewValue))
        
        If DoesImageExist(CStr("STEP" & vNewValue)) = True Then
            Set Me.imgHeader.Picture = Me.ImageList1.ListImages(CStr("STEP" & vNewValue)).Picture
        End If
    End If
    
    If vNewValue = miSteps Then
        Me.cmdFinish.Enabled = True
    Else
        Me.cmdFinish.Enabled = False
    End If
        
    Me.pictStepContainer(vNewValue).Visible = True
    Me.pictStepContainer(vNewValue).ZOrder 0
    
    miCurrentStep = vNewValue
    
End Property

Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandle

Dim vSettings As Variant

    vSettings = Split(Item.Tag, "|")
    
    Me.chkPublicGet.Value = vSettings(0)
    Me.chkPublicLet.Value = vSettings(1)
    Me.chkPublicSet.Value = vSettings(2)
    
    Me.chkPrivateGet.Value = vSettings(3)
    Me.chkPrivateLet.Value = vSettings(4)
    Me.chkPrivateSet.Value = vSettings(5)
    
    Me.chkFriendGet.Value = vSettings(6)
    Me.chkFriendLet.Value = vSettings(7)
    Me.chkFriendSet.Value = vSettings(8)

ClearVariables:
    Exit Sub
    
ErrorHandle:
    Err.Raise Err.Number, Err.source, Err.Description
    GoTo ClearVariables
            

End Sub


