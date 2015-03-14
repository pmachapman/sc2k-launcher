VERSION 4.00
Begin VB.Form FormMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SimCity 2000"
   ClientHeight    =   1455
   ClientLeft      =   1845
   ClientTop       =   5490
   ClientWidth     =   5175
   Height          =   1965
   Icon            =   "FormMain.frx":0000
   Left            =   1785
   LinkTopic       =   "FormMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   5175
   Top             =   5040
   Width           =   5295
   Begin VB.CommandButton CommandOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox TextCompany 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox TextName 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label LabelCompany 
      Caption         =   "Mayor's &Company"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label LabelName 
      Caption         =   "Mayor's &Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' Make sure all variables are declared
Option Explicit

' Cancel button click event handler
Private Sub CommandCancel_Click()
    ' Unload the form
    Unload Me
End Sub

' Ok button click event handler
Private Sub CommandOk_Click()
    ' Get the program directory
    Dim ProgramDirectory As String
    ProgramDirectory = CurDir
    If Right(ProgramDirectory, 1) = "\" Then
        ProgramDirectory = Left(ProgramDirectory, Len(ProgramDirectory) - 1)
    End If
    
    ' Create the registry keys
    SetKeyValue "Software\Maxis\SimCity 2000\Localize", "Language", "USA", REG_SZ
    SetKeyValue "Software\Maxis\SimCity 2000\Options", "Speed", 1, REG_DWORD
    SetKeyValue "Software\Maxis\SimCity 2000\Options", "Sound", 1, REG_DWORD
    SetKeyValue "Software\Maxis\SimCity 2000\Options", "Music", 1, REG_DWORD
    SetKeyValue "Software\Maxis\SimCity 2000\Options", "AutoGoto", 1, REG_DWORD
    SetKeyValue "Software\Maxis\SimCity 2000\Options", "AutoBudget", 0, REG_DWORD
    SetKeyValue "Software\Maxis\SimCity 2000\Options", "Disasters", 1, REG_DWORD
    SetKeyValue "Software\Maxis\SimCity 2000\Options", "AutoSave", 0, REG_DWORD
    SetKeyValue "Software\Maxis\SimCity 2000\Paths", "Home", ProgramDirectory, REG_SZ
    SetKeyValue "Software\Maxis\SimCity 2000\Paths", "Graphics", ProgramDirectory & "\Bitmaps", REG_SZ
    SetKeyValue "Software\Maxis\SimCity 2000\Paths", "Music", ProgramDirectory & "\Sounds", REG_SZ
    SetKeyValue "Software\Maxis\SimCity 2000\Paths", "Data", ProgramDirectory & "\Data", REG_SZ
    SetKeyValue "Software\Maxis\SimCity 2000\Paths", "Goodies", ProgramDirectory & "\Goodies", REG_SZ
    SetKeyValue "Software\Maxis\SimCity 2000\Paths", "Cities", ProgramDirectory & "\Cities", REG_SZ
    SetKeyValue "Software\Maxis\SimCity 2000\Paths", "SaveGame", ProgramDirectory & "\Cities", REG_SZ
    SetKeyValue "Software\Maxis\SimCity 2000\Paths", "TileSets", ProgramDirectory & "\ScurkArt", REG_SZ
    SetKeyValue "Software\Maxis\SimCity 2000\Paths", "Scenarios", ProgramDirectory & "\Scenario", REG_SZ
    SetKeyValue "Software\Maxis\SimCity 2000\REGISTRATION", "Mayor Name", TextName.Text, REG_SZ
    SetKeyValue "Software\Maxis\SimCity 2000\REGISTRATION", "Company Name", TextCompany.Text, REG_SZ
    SetKeyValue "Software\Maxis\SimCity 2000\Version", "SimCity 2000", 256, REG_DWORD
    
    ' Run SimCity and end the program
    On Error GoTo ProgramNotFound
    Shell ProgramName
    Unload Me
    
    ' Program not found error handler
ProgramNotFound:
    MsgBox "The program " & ProgramName & " could not be found.", vbOKOnly + vbCritical, "Error"
    Resume Next
End Sub

' Form load event handler
Private Sub Form_Load()
    ' Set the mayor's name to the user's name
    TextName.Text = UserName()
    TextName.SelStart = 0
    TextName.SelLength = Len(TextName.Text)
End Sub

' Form Unload event handler
Private Sub Form_Unload(Cancel As Integer)
    ' End the program
    End
End Sub
