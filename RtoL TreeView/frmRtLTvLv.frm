VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Treeview 
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRtLTvLv.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRtLTvLv.frx":0393
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRtLTvLv.frx":072C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRtLTvLv.frx":0B7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRtLTvLv.frx":0FD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRtLTvLv.frx":1422
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7260
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   12806
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7260
      Left            =   5760
      TabIndex        =   0
      Top             =   480
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   12806
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Tree View and List View Right To Left Display"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11655
   End
End
Attribute VB_Name = "Treeview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  
  Dim nd As Node, i As Integer, OldLong As Long
  Dim StrText As String, bkMark As String, FoldType As Integer, strNode As Integer
  
  Call ConnectAccessDb
  Set rs = dbs.OpenRecordset("SELECT * FROM [Tree Table]", dbOpenDynaset)
  rs.FindFirst "GroupName Is Null"

  Do Until rs.NoMatch
    StrText = rs![NodeName]
    strNode = rs![nodeid]
    FoldType = rs!FoldType
    Set nd = TreeView1.Nodes.Add(, , StrText, StrText, FoldType)
    nd.Tag = strNode
    
    bkMark = rs.Bookmark
    AddChildren nd, rs
    rs.Bookmark = bkMark
    rs.FindNext "GroupName Is Null"
    
  Loop
    
CloseAccessDb
    
    For Each nd In TreeView1.Nodes
        nd.Expanded = True
        i = i + 1
        FoldType = TreeView1.Nodes(i).Image
        TreeView1.Nodes(i).Image = IIf(FoldType = 1, 2, FoldType)
        
    Next nd
    
  
    AddList TreeView1.Nodes(1)
    
    'SubClassTreeView TreeView1, RGB(255, 255, 234)
    SetWindowLong TreeView1.hWnd, GWL_EXSTYLE, WS_EX_LAYOUTRTL
    SetWindowLong ListView1.hWnd, GWL_EXSTYLE, WS_EX_LAYOUTRTL
    OldLong = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    
End Sub
Private Sub AddChildren(nodBoss As Node, rst As DAO.Recordset)
 Dim nd As Node, StrText As String, bkMark As String, FoldType As Integer, strNode As Integer

 rst.FindFirst "[GroupName] ='" & nodBoss & "'"

 Do Until rst.NoMatch
    strNode = rst![nodeid]
    StrText = rst![NodeName]
    FoldType = rst!FoldImage
    Set nd = TreeView1.Nodes.Add(nodBoss, tvwChild, StrText, StrText, FoldType)
    nd.Tag = strNode
    bkMark = rst.Bookmark
 
    AddChildren nd, rst
    rst.Bookmark = bkMark
    rst.FindNext "[GroupName] ='" & nodBoss & "'"

 Loop

End Sub
Private Sub AddList(nodBoss As String)
 
 Dim nd As Node, StrText As String, bkMark As String, FoldType As Integer, DocType As String
 Dim LItem As ListItem, TempText As String, TempFold As Integer
 
 ConnectAccessDb

If nodBoss = "ÞÇÜÆÜãÉ ÇáÜÐßí" Then
 Set rs = dbs.OpenRecordset("SELECT * FROM [Tree Table] WHERE FoldType=3 OR FoldType=4 OR FoldType=5 ", dbOpenDynaset, dbReadOnly)
Else
 Set rs = dbs.OpenRecordset("SELECT * FROM [Tree Table] WHERE (FoldType=3 OR FoldType=4 OR FoldType=5) And GroupName='" & nodBoss & "'", dbOpenDynaset, dbReadOnly)
End If

Me.ListView1.ListItems.Clear

Do Until rs.EOF
    
    StrText = rs![NodeName]
    FoldType = rs!FoldType
   
   If FoldType = 3 Then
    Set LItem = ListView1.ListItems.Add(, , StrText, FoldType)
   Else
    Set LItem = ListView1.ListItems.Add(, , StrText, 4)
   End If
   
    rs.MoveNext
Loop
 
    CloseAccessDb
   
End Sub

