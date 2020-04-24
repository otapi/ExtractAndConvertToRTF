VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim ScriptDoc
Dim fso
Private Sub UnzupAndConvert_Click()
Set ScriptDoc = ActiveDocument
    Dim directory As String
    directory = InputBox("What is the directory to process?", "Unzip and Convert")
    
    Dim folder, files
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    doDir directory
    PrintxOut "Everything done!"
End Sub

Private Sub doDir(directory)
    Set folder = fso.GetFolder(directory)
    Set files = folder.files
    
    PrintxOut directory
    For Each file In files
        Dim newName As String
        newName = Replace(file.Path, ".docx", ".rtf")
        newName = Replace(newName, ".doc", ".rtf")
        
        On Error Resume Next
        If file.Path <> newName Then
            Dim doc
            PrintxOut file.Path & " ## Convert to rtf"
            Documents.Open FileName:=file.Path, _
            ConfirmConversions:=False, ReadOnly:=True, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", Revert:=False, _
            WritePasswordDocument:="", WritePasswordTemplate:="", Format:= _
            wdOpenFormatAuto, XMLTransform:=""
            
            Set doc = ActiveDocument
            ScriptDoc.Activate
            
            doc.SaveAs2 FileName:=newName, _
            FileFormat:=wdFormatRTF, LockComments:=False, Password:="", _
            AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
            EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
            :=False, SaveAsAOCELetter:=False, CompatibilityMode:=0
            
            doc.Close
            Kill file.Path
        ElseIf InStr(LCase(file.Path), ".zip") > 0 Or InStr(LCase(file.Path), ".rar") > 0 Then
            Dim NewFolderName As Variant
            NewFolderName = Left(file.Path, InStrRev(file.Path, ".") - 1)
            PrintxOut file.Path & " ## Extract"
            Call UnzipAFile(file.Path, NewFolderName)
            Kill file.Path
        End If

    Next
    
    Set folder = fso.GetFolder(directory)
    Set dirs = folder.subfolders
    
    For Each dirss In dirs
        doDir dirss
    Next
    
End Sub

Sub UnzipAFile(zippedFileFullName As Variant, unzipToPath As Variant)
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 7

    Dim RarIt, sh As String
    Dim SevenZipPath As String

    SevenZipPath = "c:\Program Files\7-Zip\"

    sh = Chr(34) & SevenZipPath & "7z.exe" & Chr(34) & " x " & Chr(34) & zippedFileFullName & Chr(34) & " -o" & Chr(34) & unzipToPath & Chr(34) & " -y"
    wsh.Run sh, windowStyle, waitOnReturn
End Sub

Sub PrintxOut(Message As Variant)
    ScriptDoc.Content.InsertAfter Text:=Message & vbCrLf
    Selection.EndKey Unit:=wdStory
End Sub


