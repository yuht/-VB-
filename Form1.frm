VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ַ"
   ClientHeight    =   11610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   21675
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11610
   ScaleWidth      =   21675
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command2 
      Caption         =   "������ҵ���ƣ�������һ��"
      Height          =   360
      Left            =   15750
      TabIndex        =   8
      Top             =   90
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Height          =   11055
      Index           =   2
      Left            =   14580
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   495
      Width           =   7065
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Height          =   11055
      Index           =   1
      Left            =   7740
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Form1.frx":08CA
      Top             =   495
      Width           =   6795
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      CausesValidation=   0   'False
      Height          =   10410
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   7680
      ExtentX         =   13547
      ExtentY         =   18362
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����һ��"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   50
      Width           =   1485
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Height          =   345
      Index           =   0
      Left            =   945
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   11070
      Width           =   6585
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   270
      Left            =   990
      TabIndex        =   0
      Text            =   "��ɽ����ҵ���޹�˾"
      Top             =   90
      Width           =   3615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "��������һ����ҵ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8865
      TabIndex        =   9
      Top             =   90
      Width           =   3300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���Խ����"
      Height          =   180
      Left            =   15
      TabIndex        =   4
      Top             =   11160
      Width           =   900
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "��˾����"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim zzzzzzzzPause  As Boolean

Private Sub Command1_Click()
    If Text1.Text = "" Then
        MsgBox "��������ȷ����ַ", , "����"
        Text1.SetFocus
    Else
        WebBrowser1.Navigate "http://www.qixin.com/search?key=" & UTF8Encode(Text1.Text)
    End If
End Sub
 

Private Sub Command2_Click()
    Dim kk
    Dim i
    kk = Split(Text2(1), vbCrLf)
    If UBound(kk) Then
        For i = 0 To UBound(kk)
            If Len(Trim(kk(i))) Then
            
                Text2(0) = ""
                Text1 = kk(i)
                Call Command1_Click
                
                zzzzzzzzPause = False
                
                Do
                    DoEvents
                Loop While WebBrowser1.ReadyState <> 4
                
                Do
                    DoEvents
                Loop While (zzzzzzzzPause = False)

                If Len(Trim(Text2(0))) Then
                    Text2(2) = Text2(2) & Text2(0) & vbCrLf
                End If
                
            End If
        Next
    End If
End Sub

Private Sub Form_Load()
    zzzzzzzzPause = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub WebBrowser1_DownloadComplete()
   On Error GoTo exitsub
    
    Static lastPos

    Dim WebCont
    Dim pos
    Dim posr
    Err.Clear
    WebCont = WebBrowser1.Document.documentElement.outerHTML
    
    pos = InStr(1, WebCont, "��ַ��")
    
    If pos Then
        posr = InStr(pos, WebCont, "<")
        Text2(0) = Mid(WebCont, pos + 3, posr - pos - 3)
        If Text2(0) <> lastPos Then
            zzzzzzzzPause = True
            lastPos = Text2(0)
        End If
    End If
exitsub:
    
End Sub

Public Function UTF8Encode(ByVal szInput As String) As String
    Dim wch      As String
    Dim uch      As String
    Dim szRet    As String
    Dim x        As Long
    Dim inputLen As Long
    Dim nAsc     As Long
    Dim nAsc2    As Long
    Dim nAsc3    As Long
     
    If szInput = "" Then
        UTF8Encode = szInput
        Exit Function

    End If

    inputLen = Len(szInput)

    For x = 1 To inputLen
        '�õ�ÿ���ַ�
        wch = Mid(szInput, x, 1)
        '�õ���Ӧ��UNICODE����
        nAsc = AscW(wch)

        '����<0�ı��롡����Ҫ����65536
        If nAsc < 0 Then nAsc = nAsc + 65536

        '����<128λ��ASCII�ı������������
        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else

            If (nAsc And &HF000) = 0 Then
                '�����ĵڶ�����뷶ΧΪ000080 - 0007FF
                'Unicode�ڷ�ΧD800-DFFF�в������κ��ַ�������������ƽ����Լ���������Χ����UTF-16��չ��ʶ����ƽ�棨����UTF-16��ʾһ������ƽ���ַ���.
                '��Ȼ���κα��붼�ǿ��Ա�ת���������Χ������unicode�����ǲ��������κκϷ���ֵ��
     
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
                 
            Else
                '���������00000800 �C 0000FFFF
                '����ȡ��ǰ��λ��11100000���л�ȥ���õ�UTF-8�����ǰ8λ
                '���ȡ��ǰ10λ��111111���в����㣬�������ܵõ���ǰ10�����6λ�������ı��롡����10000000���л��������õ�UTF-8�����м��8λ
                '�������111111���в����㣬�������ܵõ������6λ�������ı��롡����10000000���л��������õ�UTF-8�������8λ����
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch

            End If

        End If

    Next
     
    UTF8Encode = szRet

End Function

