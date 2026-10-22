Attribute VB_Name = "CertificateGenerator"
Option Explicit

'==============================================================================
' Certificate Generator
' Creates one PDF certificate per participant via Word Mail Merge.
'
' Requirements (same folder as this workbook):
'   CertificateTemplate.docx  - Word template with MERGEFIELDs
' Output:
'   Subfolder "Certificates" with one file per participant: Lastname_Firstname.pdf
'
' Data flow:
'   Sheet1 (input, with dropdown validation)
'     -> Sheet2 (merge source: grade -> feedback text via VLOOKUP)
'       -> Word template -> PDF per participant
'==============================================================================

Private Const TEMPLATE_NAME As String = "CertificateTemplate.docx"
Private Const OUTPUT_FOLDER As String = "Certificates"
Private Const INPUT_SHEET As String = "Sheet1"
Private Const MERGE_SHEET As String = "Sheet2"

' Word constants (late binding, no reference to the Word library needed)
Private Const wdFormLetters As Long = 0
Private Const wdSendToNewDocument As Long = 0
Private Const wdLastRecord As Long = -5          ' WdMailMergeActiveRecord.wdLastRecord
Private Const wdExportFormatPDF As Long = 17

Public Sub GenerateCertificates()
    Dim wordApp As Object
    Dim templateDoc As Object
    Dim outDoc As Object
    Dim templatePath As String
    Dim outputPath As String
    Dim pdfPath As String
    Dim firstName As String
    Dim lastName As String
    Dim lastRec As Long
    Dim i As Long
    Dim created As Long
    Dim failed As Boolean
    Dim wbFolder As String

    On Error GoTo ErrHandler

    ' Local folder of this workbook (also works if it is synced via OneDrive)
    wbFolder = GetLocalFolder()
    templatePath = wbFolder & Application.PathSeparator & TEMPLATE_NAME
    outputPath = wbFolder & Application.PathSeparator & OUTPUT_FOLDER

    ' --- Checks before starting Word ---
    If Dir(templatePath) = "" Then
        MsgBox "Word template not found:" & vbNewLine & templatePath & vbNewLine & vbNewLine & _
               "Place " & TEMPLATE_NAME & " in the same folder as this workbook.", _
               vbExclamation, "Certificate Generator"
        Exit Sub
    End If

    If CountParticipants() = 0 Then
        MsgBox "No participants entered in " & INPUT_SHEET & ".", vbExclamation, "Certificate Generator"
        Exit Sub
    End If

    If Dir(outputPath, vbDirectory) = "" Then MkDir outputPath

    ' Mail Merge reads the saved file, so save first
    ThisWorkbook.Save

    ' --- Start Word in the background ---
    Application.StatusBar = "Starting Word..."
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0

    Set templateDoc = wordApp.Documents.Open(Filename:=templatePath, ReadOnly:=True, AddToRecentFiles:=False)

    With templateDoc.MailMerge
        .MainDocumentType = wdFormLetters
        .OpenDataSource Name:=wbFolder & Application.PathSeparator & ThisWorkbook.Name, _
                        ConfirmConversions:=False, _
                        ReadOnly:=True, _
                        AddToRecentFiles:=False, _
                        Revert:=False, _
                        Format:=0, _
                        SQLStatement:="SELECT * FROM [" & MERGE_SHEET & "$] WHERE [First Name] <> 'No'"
        .Destination = wdSendToNewDocument
        .SuppressBlankLines = True

        ' Number of records after filtering
        .DataSource.ActiveRecord = wdLastRecord
        lastRec = .DataSource.ActiveRecord

        ' --- One merge + PDF export per participant ---
        For i = 1 To lastRec
            .DataSource.ActiveRecord = i
            firstName = CleanFileName(GetField(.DataSource, "First Name"))
            lastName = CleanFileName(GetField(.DataSource, "Last Name"))
            Application.StatusBar = "Creating certificate " & i & " of " & lastRec & ": " & firstName & " " & lastName

            .DataSource.FirstRecord = i
            .DataSource.LastRecord = i
            .Execute Pause:=False

            Set outDoc = wordApp.ActiveDocument
            pdfPath = outputPath & Application.PathSeparator & lastName & "_" & firstName & ".pdf"
            outDoc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdExportFormatPDF
            outDoc.Close SaveChanges:=False
            Set outDoc = Nothing

            created = created + 1
        Next i
    End With

CleanUp:
    On Error Resume Next
    If Not outDoc Is Nothing Then outDoc.Close SaveChanges:=False
    If Not templateDoc Is Nothing Then templateDoc.Close SaveChanges:=False
    If Not wordApp Is Nothing Then wordApp.Quit SaveChanges:=False
    Set outDoc = Nothing
    Set templateDoc = Nothing
    Set wordApp = Nothing
    Application.StatusBar = False
    On Error GoTo 0

    If Not failed Then
        MsgBox created & " certificate(s) created in:" & vbNewLine & outputPath, vbInformation, "Certificate Generator"
    End If
    Exit Sub

ErrHandler:
    failed = True
    MsgBox "The certificates could not be created." & vbNewLine & vbNewLine & _
           "Error " & Err.Number & ": " & Err.Description & vbNewLine & vbNewLine & _
           created & " certificate(s) were created before the error.", _
           vbCritical, "Certificate Generator"
    Resume CleanUp
End Sub

' Returns the local folder of this workbook.
' In OneDrive folders, ThisWorkbook.Path is a web address (https://...),
' which Dir, MkDir and Word cannot use. This converts it back to the local path.
Private Function GetLocalFolder() As String
    Dim p As String
    Dim root As String
    Dim rest As String
    Dim pos As Long

    p = ThisWorkbook.Path
    If LCase$(Left$(p, 4)) <> "http" Then
        GetLocalFolder = p
        Exit Function
    End If

    If InStr(1, p, "d.docs.live.net", vbTextCompare) > 0 Then
        ' Personal OneDrive: https://d.docs.live.net/<id>/Folder/Subfolder
        root = Environ$("OneDriveConsumer")
        If root = "" Then root = Environ$("OneDrive")
        pos = InStr(Len("https://d.docs.live.net/") + 1, p, "/")
        If pos > 0 Then rest = Mid$(p, pos + 1)
    Else
        ' OneDrive for Business: https://<tenant>-my.sharepoint.com/personal/<user>/Documents/Folder
        root = Environ$("OneDriveCommercial")
        If root = "" Then root = Environ$("OneDrive")
        pos = InStr(1, p, "/Documents/", vbTextCompare)
        If pos > 0 Then rest = Mid$(p, pos + Len("/Documents/"))
    End If

    If root = "" Then
        Err.Raise vbObjectError + 1, , "This workbook is stored in OneDrive and its local folder could not be found." & _
                  vbNewLine & "Move the folder to a local drive (e.g. C:\Projects) and try again."
    End If

    If rest = "" Then
        GetLocalFolder = root
    Else
        GetLocalFolder = root & "\" & Replace(rest, "/", "\")
    End If
End Function

' Counts filled-in first names on the input sheet
Private Function CountParticipants() As Long
    CountParticipants = Application.WorksheetFunction.CountA( _
        ThisWorkbook.Worksheets(INPUT_SHEET).Range("B2:B1000"))
End Function

' Reads a field from the current Mail Merge record.
' Word may show column headers with spaces ("First Name") or underscores ("First_Name").
Private Function GetField(ByVal dataSource As Object, ByVal fieldName As String) As String
    Dim fld As Object
    Dim wanted As String
    wanted = LCase$(Replace(fieldName, " ", "_"))
    For Each fld In dataSource.DataFields
        If LCase$(Replace(fld.Name, " ", "_")) = wanted Then
            GetField = Trim$(CStr(fld.Value))
            Exit Function
        End If
    Next fld
    GetField = "Unknown"
End Function

' Removes characters that are not allowed in Windows file names
Private Function CleanFileName(ByVal s As String) As String
    Dim badChars As Variant
    Dim c As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each c In badChars
        s = Replace(s, c, "")
    Next c
    s = Trim$(Replace(s, " ", "-"))
    If s = "" Then s = "Unknown"
    CleanFileName = s
End Function
