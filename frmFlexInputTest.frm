VERSION 5.00
Begin VB.Form frmFlexInputTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Smart Flex Input Class Test"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   LinkTopic       =   "frmFlexInputTest"
   MaxButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7740
   StartUpPosition =   1  'ËùÓÐÕßÖÐÐÄ
   Begin VB.Frame fraDateTimeValidation 
      Caption         =   "(Date/Time) Filter && verify"
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton optIDType 
         Appearance      =   0  'Flat
         Caption         =   "Time"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optDateTimeType 
         Appearance      =   0  'Flat
         Caption         =   "Date"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   11
         Top             =   840
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtDateTimeFilterAndValidation 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdDateTimeValidation 
         Caption         =   "Validate"
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblDateTimeType 
         BackStyle       =   0  'Transparent
         Caption         =   "DateTime &Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame fraTextValidation 
      Caption         =   "(Text)Filter, verify && correct"
      Height          =   3255
      Left            =   3960
      TabIndex        =   13
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txtTextMaxLength 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1800
         TabIndex        =   22
         Text            =   "20"
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton cmdTextValidation 
         Caption         =   "Validate"
         Height          =   375
         Left            =   2280
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtTextFilterAndValidation 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2055
      End
      Begin VB.CheckBox chkAutReplace 
         Appearance      =   0  'Flat
         Caption         =   "A&uto correct"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtForbiddenChars 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1800
         TabIndex        =   17
         Text            =   "'|""|*|^"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtReplaceChar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1800
         TabIndex        =   21
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtSplitChar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1800
         TabIndex        =   19
         Text            =   "|"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblTextMaxLength 
         BackStyle       =   0  'Transparent
         Caption         =   "&Maximum length:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label lblForbiddenChars 
         BackStyle       =   0  'Transparent
         Caption         =   "Forbidden &Chars:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblTipForbiddenChars 
         BackStyle       =   0  'Transparent
         Caption         =   "&Split Char"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblReplaceChar 
         BackStyle       =   0  'Transparent
         Caption         =   "&Replace char"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   1455
      End
   End
   Begin VB.Frame fraNumberValidation 
      Caption         =   "(Number) Filter && verify"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   3615
      Begin VB.TextBox txtNumberFilterAndValidation 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdNumberValidation 
         Caption         =   "Validate"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtMaxNum 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   1800
         TabIndex        =   6
         Text            =   "9999999999"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtMinNum 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   1800
         TabIndex        =   4
         Text            =   "-9999999999"
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblMaxNum 
         BackStyle       =   0  'Transparent
         Caption         =   "M&aximum number:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblMinNum 
         BackStyle       =   0  'Transparent
         Caption         =   "M&inimum number:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame fraInputFilterAndValidation 
      Caption         =   "General filter && validate"
      Height          =   855
      Left            =   120
      TabIndex        =   25
      Top             =   3600
      Width           =   7455
      Begin VB.CheckBox chkCanBeEmpty 
         Appearance      =   0  'Flat
         Caption         =   "Allow &Empty"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5640
         TabIndex        =   29
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chkShowMsg 
         Appearance      =   0  'Flat
         Caption         =   "Enable &Hint"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   28
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkEnableValidation 
         Appearance      =   0  'Flat
         Caption         =   "Enable &Verify"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkEnableFilter 
         Appearance      =   0  'Flat
         Caption         =   "Enable &Filter"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmFlexInputTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************
'*
'*   Program Name: Smart Flex Input
'*   Description : filter, verify and correct input
'*   Version     : 0.5
'*   Copyright   : Smartsoft 2002
'*
'*******************************************************************

Option Explicit

Private m_udtETRet As enumErrorType

Private m_FlexInput As cFlexInput



Private Sub cmdDateTimeValidation_Click()
    
    If chkEnableValidation.Value Then
        m_FlexInput.ValidateDateTime txtDateTimeFilterAndValidation, _
                                        , _
                                        , _
                                        , _
                                        , _
                                        , _
                                        chkCanBeEmpty.Value, _
                                        chkShowMsg.Value
    End If
    
End Sub


Private Sub cmdNumberValidation_Click()

    If chkEnableValidation.Value Then
        m_FlexInput.ValidateNumber txtNumberFilterAndValidation, _
                                        CCur(txtMinNum.Text), _
                                        CCur(txtMaxNum.Text), _
                                        Len(txtMaxNum.Text), _
                                        chkCanBeEmpty.Value, _
                                        chkShowMsg.Value
    End If
    
End Sub


Private Sub cmdTextValidation_Click()
        
    If chkEnableValidation.Value Then
        m_FlexInput.ValidateText txtTextFilterAndValidation, _
                                Trim(txtForbiddenChars.Text), _
                                Trim(txtSplitChar.Text), _
                                Trim(txtReplaceChar.Text), _
                                chkAutReplace.Value, _
                                CInt(txtTextMaxLength.Text), _
                                chkCanBeEmpty.Value, _
                                chkShowMsg.Value
    End If
    
End Sub


Private Sub Form_Load()
    
    Set m_FlexInput = New cFlexInput
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set m_FlexInput = Nothing
    
    Set frmFlexInputTest = Nothing

End Sub


Private Sub txtMinNum_KeyPress(KeyAscii As Integer)
        
    If chkEnableFilter.Value Then
        KeyAscii = m_FlexInput.FilterInput(txtMinNum, _
                                            KeyAscii, _
                                            itNumeric, _
                                            "0-9", _
                                            , _
                                            14, _
                                            chkShowMsg.Value)
    End If

End Sub


Private Sub txtMaxNum_KeyPress(KeyAscii As Integer)
    
    If chkEnableFilter.Value Then
        KeyAscii = m_FlexInput.FilterInput(txtMaxNum, _
                                            KeyAscii, _
                                            itNumeric, _
                                            "0-9", _
                                            , _
                                            14, _
                                            chkShowMsg.Value)
    End If

End Sub


Private Sub txtNumberFilterAndValidation_KeyPress(KeyAscii As Integer)
    
    If chkEnableFilter.Value Then
        KeyAscii = m_FlexInput.FilterInput(txtNumberFilterAndValidation, _
                                            KeyAscii, _
                                            itNumeric, _
                                            "0-9", _
                                            , _
                                            Len(txtMaxNum.Text), _
                                            chkShowMsg.Value)
    End If
    
End Sub


Private Sub txtDateTimeFilterAndValidation_KeyPress(KeyAscii As Integer)
    
    If chkEnableFilter.Value Then
        KeyAscii = m_FlexInput.FilterInput(txtDateTimeFilterAndValidation, _
                                            KeyAscii, _
                                            itDate, _
                                            "0-9", _
                                            , _
                                            , _
                                            chkShowMsg.Value)
    End If
    
End Sub


Private Sub txtTextMaxLength_KeyPress(KeyAscii As Integer)
    
    If chkEnableFilter.Value Then
        KeyAscii = m_FlexInput.FilterInput(txtTextMaxLength, _
                                            KeyAscii, _
                                            itNumeric, _
                                            "0-9", _
                                            , _
                                            2, _
                                            chkShowMsg.Value)
    End If

End Sub


Private Sub txtTextFilterAndValidation_KeyPress(KeyAscii As Integer)
    
    If chkEnableFilter.Value Then
        KeyAscii = m_FlexInput.FilterInput(txtTextFilterAndValidation, _
                                            KeyAscii, _
                                            itChar, _
                                            "0-9|a-z|A-Z", _
                                            , _
                                            CInt(txtTextMaxLength.Text), _
                                            chkShowMsg.Value)
    End If
    
End Sub


