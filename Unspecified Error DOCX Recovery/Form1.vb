Option Explicit On
Option Strict Off
Option Compare Binary

Imports System
Imports Microsoft
Imports Microsoft.VisualBasic
Imports System.Runtime.InteropServices
Imports System.Xml
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Reflection
Imports Microsoft.Office.Interop.Word
Imports Word = Microsoft.Office.Interop.Word
Imports System.Threading
Imports System.Text
Imports Microsoft.VisualBasic.Strings
Imports System.Xml.Schema

Public Class Form1

    Inherits System.Windows.Forms.Form
    Dim firstTimeThrough As Boolean = False
    Dim selectedFile As String
    Dim tempPath As String = System.IO.Path.GetTempPath() + "Savvy"

    Private Property ret As Integer

    Private Shared Sub ValidationCallBack(sender As Object, args As ValidationEventArgs)
        If (args.Severity = XmlSeverityType.Warning) Then
            MsgBox("   Warning: Matching schema not found.  No validation occurred." + args.Message)
        Else
            MsgBox("   Validation error: " + args.Message)
        End If
    End Sub
    Public Sub SelectFile()

        Dim exists As Boolean = System.IO.Directory.Exists(tempPath)

        If Not exists Then
            System.IO.Directory.CreateDirectory(tempPath)
        End If

        'Makes the progress bar disappear until recovery starts.
        'ProgressBar1.Hide()
        'Opens file dialogu box.
        Dim dlgOpenFile As New OpenFileDialog
        'These are all the modern Office Open XML Word 2007 - 
        '2013 formats, replacing the old .doc and .dot.
        dlgOpenFile.Filter = "Microsoft 2007-2013 Format Word Documents " _
            & "(*.docx;*.docm;*.dotx;*.dotm)|*.docx;*.docm;*.dotx;*.dotm|All files (*.*)|*.*"
        dlgOpenFile.RestoreDirectory = True

        If dlgOpenFile.ShowDialog() = DialogResult.OK Then

            'Writes selected file full path to text box next to browse button.
            TextBox1.Text = dlgOpenFile.FileName

        End If

    End Sub

    Public Function DelFromLeft(ByVal sChars As String, _
      ByVal sLine As String) As String

        ' Removes unwanted characters from left of given string
        '  EXAMPLE
        '      'MsgBox DelFromLeft("THIS", "THIS IS A TEST")
        '        displays  "IS A TEST"

        Dim iCount As Integer
        Dim sChar As String

        DelFromLeft = ""
        ' Remove unwanted characters to left of folder name
        If InStr(sLine, sChars) > 0 Then
            For iCount = 1 To Len(sChars)
                ' Retrieve character from start string to 
                'look for in folder string (sLine)
                sChar = Mid$(sChars, iCount, 1)
                ' Remove all characters to left of found string
                sLine = Mid$(sLine, InStr(sLine, sChar) + 1)

            Next iCount
        End If
        DelFromLeft = sLine
        Exit Function

    End Function

    Declare Auto Function GetShortPathName Lib "kernel32.dll" _
(ByVal lpszLongPath As String, ByVal lpszShortPath As StringBuilder, _
ByVal cchBuffer As Integer) As Integer

    Private Sub Impl_Automatic_Recovery()

        selectedFile = TextBox1.Text
        Dim rand1 As New Random()
        Dim number9Places As Integer = rand1.Next(1, 1000000000)
        Dim selectedFileInfo As FileInfo = New FileInfo(selectedFile)
        Dim selectedFileName As String = LCase(selectedFileInfo.Name)
        Dim selectedFileExtension As String = LCase(selectedFileInfo.Extension)
        Dim sFileFullNameLC As String = LCase(selectedFileInfo.FullName.ToString)

        'Dim selectedFileBasePath As String = LCase(selectedFileInfo.Directory.ToString)
        Dim selectedFileBasePath As String = tempPath
        Dim selectedFileLC As String = LCase(selectedFile)
        Dim zipRepairedsFileNameNonZipExt As String = "zipRepaired_2_" & selectedFileName
        Dim zipRepairedsFileNameNonZipExtNoSpace As String = zipRepairedsFileNameNonZipExt.Replace(" ", "_")
        Dim zipRepairedsFileNameNonZipExtNoSpaceFullName As String = selectedFileBasePath & _
            "\" & zipRepairedsFileNameNonZipExtNoSpace
        Dim salvagedsFileName As String = "salvaged_" & selectedFileName
        Dim salvagedsFileName2 As String = "salvaged_2_" & selectedFileName
        Dim salvagedsFileNameNoSpace As String = salvagedsFileName.Replace(" ", "_") & ".txt"
        Dim salvagedsFileNameNoSpace2 As String = salvagedsFileName2.Replace(" ", "_") & ".txt"
        Dim salvagedsFileNameAndBasePathNoSpace As String = selectedFileBasePath & "\" & salvagedsFileNameNoSpace
        Dim salvagedsFileNameAndBasePathNoSpace2 As String = selectedFileBasePath & "\" & salvagedsFileNameNoSpace2
        Dim salvagedsFileNameAndBasePathNoSpaceOld As String = selectedFileBasePath & "\old_" & salvagedsFileNameNoSpace
        Dim salvagedsFileNameAndBasePathNoSpaceOld2 As String = selectedFileBasePath & "\old_" & salvagedsFileNameNoSpace2
        Dim sFile As String = selectedFileBasePath & "\" & number9Places.ToString & selectedFileExtension
        Dim zipRepairedsFileWithZipExtension As String = selectedFileBasePath & _
            "\zipRepaired_" & number9Places.ToString & ".docx.zip"
        Dim zipRepairedDirectoryName As String = "zipRepaired_" & number9Places.ToString
        Dim zipRepairedModDirectoryName As String = "zipRepaired_" & number9Places.ToString & "_mod"
        Dim sFileInfo As FileInfo = New FileInfo(sFile)
        Dim sFileExtension As String = LCase(sFileInfo.Extension)
        Dim x As Integer = 0
        Dim sFileName As String = LCase(sFileInfo.Name)
        'Dim sFileBasePath As String = LCase(sFileInfo.Directory.ToString)
        Dim sFileBasePath As String = tempPath
        Dim sFileFullName As String = LCase(sFileInfo.FullName.ToString)
        Dim sFileNameWithoutExtension As String = sFileName.Substring(0, sFileInfo.Name.Length - 5)
        Dim sFileNameChangedtoDocExtension = sFileNameWithoutExtension & ".doc"
        Dim sFileBasePathAndNameChangedtoDocExtension = sFileBasePath & "\" & sFileNameChangedtoDocExtension
        Dim sFileNameChangedtoRtfExtension = sFileNameWithoutExtension & ".rtf"
        Dim sFileBasePathAndNameChangedtoRtfExtension = sFileBasePath & "\" & sFileNameChangedtoRtfExtension
        Dim sFileNameChangedtoWpsExtension = sFileNameWithoutExtension & ".wps"
        Dim sFileBasePathAndNameChangedtoWpsExtension = sFileBasePath & "\" & sFileNameChangedtoWpsExtension
        Dim prettyPrintedDocumentXmlsFile As String = "pretty_printed_" & sFileName
        Dim prettyPrintedDocumentXmlFullPath As String = sFileBasePath & "\" & prettyPrintedDocumentXmlsFile
        Dim objWord As Word.Application
        Dim objDoc As Word.Document

        RichTextBox1.Clear()

        If String.IsNullOrEmpty(TextBox1.Text) = True Then

            SelectFile()

        Else

            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            'ProgressBar1.Style = ProgressBarStyle.Marquee
            ProgressBar1.MarqueeAnimationSpeed = 100
            ProgressBar1.Show()
            ProgressBar1.Value = 10

            File.Copy(selectedFile, sFile, True)

            ' initializes and sets up parameters of progress bar.

            Dim byteMatchDoc As Match = Regex.Match(Trid_Identify(sFileFullName), "\(.DOC\)")
            Dim byteMatchRTF As Match = Regex.Match(Trid_Identify(sFileFullName), "\(.RTF\)")
            Dim byteMatchWPS As Match = Regex.Match(Trid_Identify(sFileFullName), "\(.WPS\)")

            If String.IsNullOrEmpty(byteMatchDoc.ToString) = False Then

                MsgBox("Your file may not have the right extension. We are changing the " _
                    & ".docx to .doc on a copy of the document and then will attempt to open it. " _
                    & "If this is unsuccessful, Savvy DOCX recovery will try to use the command " _
                    & "line program DocToText to recover the text.", MsgBoxStyle.Exclamation)

                'The next few lines bypass the section where the corrupt file is copied
                'to one whose name is a random number.

                If File.Exists(sFileBasePathAndNameChangedtoDocExtension) Then

                    File.Delete(sFileBasePathAndNameChangedtoDocExtension)

                End If

                File.Copy(selectedFile, sFileBasePathAndNameChangedtoDocExtension, True)

                objWord = New Word.Application
                objDoc = New Word.Document

                Try
                    ProgressBar1.Value = 20
                    'Will try to open Word file with the Open and Repair feature.
                    'If successful there's no need to do more and it makes Word visible
                    'even though it was started invisibly. If unsuccessful the program 
                    'moves to the first catch section.

                    objDoc = objWord.Documents.OpenNoRepairDialog(sFileBasePathAndNameChangedtoDocExtension, OpenAndRepair:=True)
                    objWord.Visible = True
                    objWord.WindowState = Word.WdWindowState.wdWindowStateMaximize
                    objWord.Activate()
                    Me.Cursor = System.Windows.Forms.Cursors.Default

                    'ProgressBar1.Hide()

                Catch

                    Impl_S2RecoverText_From_DOC(sFileBasePathAndNameChangedtoDocExtension)

                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    'ProgressBar1.Hide()

                Finally

                    If File.Exists(sFile) Then

                        File.Delete(sFile)

                    End If

                    CloseWordApplicationWithDocument(objWord, objDoc)
                    ProgressBar1.Value = 30
                End Try



            ElseIf String.IsNullOrEmpty(byteMatchRTF.ToString) = False Then

                MsgBox("Your file may not have the right extension. We are changing the .docx " _
                    & "to .rtf on a copy of the document and then will attempt to open it. " _
                    & "If this is unsuccessful, Savvy DOCX recovery will try to use the command " _
                    & "line program DocToText to recover the text.", MsgBoxStyle.Exclamation)

                'The next few lines bypass the section where the corrupt file is copied
                'to one whose name is a random number.

                If File.Exists(sFileBasePathAndNameChangedtoRtfExtension) Then

                    File.Delete(sFileBasePathAndNameChangedtoRtfExtension)

                End If

                File.Copy(selectedFile, sFileBasePathAndNameChangedtoRtfExtension, True)

                objWord = New Word.Application
                objDoc = New Word.Document

                Try
                    ProgressBar1.Value = 40
                    'Will try to open Word file with the Open and Repair feature.
                    'If successful there's no need to do more and it makes Word visible
                    'even though it was started invisibly. If unsuccessful the program 
                    'moves to the first catch section.

                    objDoc = objWord.Documents.OpenNoRepairDialog(sFileBasePathAndNameChangedtoRtfExtension, OpenAndRepair:=True)
                    objWord.Visible = True
                    objWord.WindowState = Word.WdWindowState.wdWindowStateMaximize
                    objWord.Activate()
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    'ProgressBar1.Hide()

                Catch

                    Impl_S2RecoverText_From_DOC(sFileBasePathAndNameChangedtoRtfExtension)

                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    'ProgressBar1.Hide()

                Finally

                    If File.Exists(sFile) Then

                        File.Delete(sFile)

                    End If

                    CloseWordApplicationWithDocument(objWord, objDoc)
                    ProgressBar1.Value = 50
                End Try

            ElseIf String.IsNullOrEmpty(byteMatchWPS.ToString) = False Then

                MsgBox("Your file may not have the right extension. We are changing the " _
                    & ".docx to the Microsoft Works Writer .wps on a copy of ithe document " _
                    & "and then will attempt to open it. If this is unsuccessful, Savvy " _
                    & "DOCX recovery will try to use the command line program DocToText " _
                    & "to recover the text.", MsgBoxStyle.Exclamation)

                If File.Exists(sFileBasePathAndNameChangedtoWpsExtension) Then

                    File.Delete(sFileBasePathAndNameChangedtoWpsExtension)

                End If

                'The next few lines bypass the section where the corrupt file is copied
                'to one whose name is a random number.

                File.Copy(selectedFile, sFileBasePathAndNameChangedtoWpsExtension, True)

                objWord = New Word.Application
                objDoc = New Word.Document

                Try
                    ProgressBar1.Value = 60
                    'Will try to open Word file with the Open and Repair feature.
                    'If successful there's no need to do more and it makes Word visible
                    'even though it was started invisibly. If unsuccessful the program 
                    'moves to the first catch section.

                    objDoc = objWord.Documents.OpenNoRepairDialog(sFileBasePathAndNameChangedtoWpsExtension, OpenAndRepair:=True)
                    objWord.Visible = True
                    objWord.WindowState = Word.WdWindowState.wdWindowStateMaximize
                    objWord.Activate()
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    'ProgressBar1.Hide()

                Catch

                    Impl_S2RecoverText_From_DOC(sFileFullName)

                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    'ProgressBar1.Hide()

                Finally

                    If File.Exists(sFile) Then

                        File.Delete(sFile)

                    End If

                    CloseWordApplicationWithDocument(objWord, objDoc)
                    ProgressBar1.Value = 70
                End Try
            Else

                objWord = New Word.Application
                objDoc = New Word.Document

                Try
                    ProgressBar1.Value = 80
                    'Will try to open Word file with the Open and Repair feature.
                    'If successful there's no need to do more and it makes Word visible
                    'even though it was started invisibly. If unsuccessful the program 
                    'moves to the first catch section.

                    objDoc = objWord.Documents.OpenNoRepairDialog(selectedFile, OpenAndRepair:=True)
                    objWord.Visible = True
                    objWord.WindowState = Word.WdWindowState.wdWindowStateMaximize
                    objWord.Activate()

                    MsgBox("When opened, at least with open and repair function agreed to, this " _
                           & "file does not appear to cause an error", MsgBoxStyle.Information)
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    'ProgressBar1.Hide()

                    Exit Try

                Catch ex As Exception

                    CloseWordApplicationWithDocument(objWord, objDoc)

                    'There are three kinds off corruption as organized in this program.
                    'The Unpsecified errors are fixed by removing math tags and the in 
                    'between code. The end tag not matching error are dealt with by
                    'rearranging the order of a math tag and another one. Finally all other
                    'errors are dealt with by going through all the XML sub files and
                    'repairing them with the program xmllint. If that doesn't work all
                    'XML subfiles with errors are truncated at the error and then the
                    'correct end tags are added by the same program used in the previous 
                    'step, xmllint. Finally if that doesn't work SilverCoder's DocToText
                    'is used to try to salvage the text alone without formatting. An added
                    'Feature will probably be implimented to always provide salvaged text 
                    'with any corrupt office open xml word file in case extra text can saved.

                    If ex.Message.Contains("Unspecified error") Then

                        If Not Deal_With_Unspecified_Error(sFileFullName) = "OK" Then

                            Dim result As Integer = MessageBox.Show("Recovery was unsuccessful, " _
                                & "would you like try to salvage the text without or with limited " _
                                & "formatting?", "caption", MessageBoxButtons.YesNoCancel)
                            If result = DialogResult.Cancel Then
                                MessageBox.Show("Cancel pressed")
                            ElseIf result = DialogResult.No Then
                                MessageBox.Show("No pressed")
                            ElseIf result = DialogResult.Yes Then
                                MessageBox.Show("Two versions of text salvaging will try " _
                                & "to be opened. One with no formatting and a second with " _
                                & "formatting limited to that possible with text files.")

                                               File.Copy(sFileFullNameLC, zipRepairedsFileNameNonZipExtNoSpaceFullName, True)

                                If File.Exists(salvagedsFileNameAndBasePathNoSpace2) = True Then

                                    File.Delete(salvagedsFileNameAndBasePathNoSpace2)

                                End If

                                If File.Exists(salvagedsFileNameAndBasePathNoSpace) = True Then

                                    File.Delete(salvagedsFileNameAndBasePathNoSpace)

                                End If
                                'MsgBox("Check that zipRepairedsFileNameNonZipExtNoSpaceFullName exists called: " & zipRepairedsFileNameNonZipExtNoSpaceFullName)
                                DocToText_Process(zipRepairedsFileNameNonZipExtNoSpaceFullName, salvagedsFileNameNoSpace, sFileBasePath, True)

                                Process.Start(salvagedsFileNameAndBasePathNoSpace)
                                'ProgressBar1.Value= 50

                                DocToText_Process(zipRepairedsFileNameNonZipExtNoSpaceFullName, salvagedsFileNameNoSpace2, sFileBasePath, False)
                                Process.Start(salvagedsFileNameAndBasePathNoSpace2)

                                Me.Cursor = System.Windows.Forms.Cursors.Default
                                'ProgressBar1.Hide()

                                If File.Exists(zipRepairedsFileNameNonZipExtNoSpaceFullName) Then

                                    File.Delete(zipRepairedsFileNameNonZipExtNoSpaceFullName)

                                End If

                            End If

                        Else
                            Me.Invoke(DirectCast(Sub() ProgressBar1.Value = 100, MethodInvoker))
                            ProgressBar1.Refresh()
                            System.Windows.Forms.Application.DoEvents()
                            Exit Sub

                        End If

                    ElseIf ex.Message.Contains("The name in the end tag of the element must match the element type in the start tag.") Then

                        If Not Deal_With_Name_Tag_Error(sFileFullName) = "OK" Then

                            Dim result As Integer = MessageBox.Show("Recovery was unsuccessful, " _
                                & "would you like try to salvage the text without or with limited " _
                                & "formatting?", "caption", MessageBoxButtons.YesNoCancel)
                            If result = DialogResult.Cancel Then
                                MessageBox.Show("Cancel pressed")
                            ElseIf result = DialogResult.No Then
                                MessageBox.Show("No pressed")
                            ElseIf result = DialogResult.Yes Then
                                MessageBox.Show("Two versions of text salvaging will try " _
                                & "to be opened. One with no formatting and a second with " _
                                & "formatting limited to that possible with text files.")

                                                File.Copy(sFileFullNameLC, zipRepairedsFileNameNonZipExtNoSpaceFullName, True)

                                If File.Exists(salvagedsFileNameAndBasePathNoSpace2) = True Then

                                    File.Delete(salvagedsFileNameAndBasePathNoSpace2)

                                End If

                                If File.Exists(salvagedsFileNameAndBasePathNoSpace) = True Then

                                    File.Delete(salvagedsFileNameAndBasePathNoSpace)

                                End If
                                'MsgBox("Check that zipRepairedsFileNameNonZipExtNoSpaceFullName exists called: " & zipRepairedsFileNameNonZipExtNoSpaceFullName)
                                DocToText_Process(zipRepairedsFileNameNonZipExtNoSpaceFullName, salvagedsFileNameNoSpace, sFileBasePath, True)

                                Process.Start(salvagedsFileNameAndBasePathNoSpace)
                                'ProgressBar1.Value= 50

                                DocToText_Process(zipRepairedsFileNameNonZipExtNoSpaceFullName, salvagedsFileNameNoSpace2, sFileBasePath, False)
                                Process.Start(salvagedsFileNameAndBasePathNoSpace2)

                                Me.Cursor = System.Windows.Forms.Cursors.Default
                                'ProgressBar1.Hide()

                                If File.Exists(zipRepairedsFileNameNonZipExtNoSpaceFullName) Then

                                    File.Delete(zipRepairedsFileNameNonZipExtNoSpaceFullName)

                                End If

                            End If

                        Else

                            Me.Invoke(DirectCast(Sub() ProgressBar1.Value = 100, MethodInvoker))
                            ProgressBar1.Refresh()
                            System.Windows.Forms.Application.DoEvents()
                            Exit Sub

                        End If

                    Else
                        ProgressBar1.Value = 90
                        If Not Deal_With_Everything_Else(sFileFullName) = "OK" Then

                            Dim result As Integer = MessageBox.Show("Recovery was unsuccessful, " _
                                & "would you like try to salvage the text without or with limited " _
                                & "formatting?", "caption", MessageBoxButtons.YesNoCancel)
                            If result = DialogResult.Cancel Then
                                MessageBox.Show("Cancel pressed")
                            ElseIf result = DialogResult.No Then
                                MessageBox.Show("No pressed")
                            ElseIf result = DialogResult.Yes Then
                                MessageBox.Show("Two versions of text salvaging will try " _
                                & "to be opened. One with no formatting and a second with " _
                                & "formatting limited to that possible with text files.")

                                File.Copy(sFileFullNameLC, zipRepairedsFileNameNonZipExtNoSpaceFullName, True)

                                If File.Exists(salvagedsFileNameAndBasePathNoSpace2) = True Then

                                    File.Delete(salvagedsFileNameAndBasePathNoSpace2)

                                End If

                                If File.Exists(salvagedsFileNameAndBasePathNoSpace) = True Then

                                    File.Delete(salvagedsFileNameAndBasePathNoSpace)

                                End If
                                'MsgBox("Check that zipRepairedsFileNameNonZipExtNoSpaceFullName exists called: " & zipRepairedsFileNameNonZipExtNoSpaceFullName)
                                DocToText_Process(zipRepairedsFileNameNonZipExtNoSpaceFullName, salvagedsFileNameNoSpace, sFileBasePath, True)

                                Process.Start(salvagedsFileNameAndBasePathNoSpace)
                                'ProgressBar1.Value= 50

                                DocToText_Process(zipRepairedsFileNameNonZipExtNoSpaceFullName, salvagedsFileNameNoSpace2, sFileBasePath, False)
                                Process.Start(salvagedsFileNameAndBasePathNoSpace2)

                                Me.Cursor = System.Windows.Forms.Cursors.Default
                                'ProgressBar1.Hide()


                                If File.Exists(zipRepairedsFileNameNonZipExtNoSpaceFullName) Then

                                    File.Delete(zipRepairedsFileNameNonZipExtNoSpaceFullName)

                                End If

                            End If

                        Else

                            Me.Invoke(DirectCast(Sub() ProgressBar1.Value = 100, MethodInvoker))
                            ProgressBar1.Refresh()
                            System.Windows.Forms.Application.DoEvents()
                            Exit Sub

                        End If

                    End If

                Finally

                    Me.Cursor = System.Windows.Forms.Cursors.Default

                    If Directory.Exists(zipRepairedDirectoryName) Then

                        Directory.Delete(zipRepairedDirectoryName, True)

                    End If

                    If Directory.Exists(zipRepairedModDirectoryName) Then

                        Directory.Delete(zipRepairedModDirectoryName, True)

                    End If

                    If File.Exists(sFile) Then

                        File.Delete(sFile)

                    End If

                    If File.Exists(zipRepairedsFileWithZipExtension) Then

                        File.Delete(zipRepairedsFileWithZipExtension)

                    End If

                    CloseWordApplicationWithDocument(objWord, objDoc)


                End Try

            End If

        End If
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Invoke(DirectCast(Sub() ProgressBar1.Value = 100, MethodInvoker))
        ProgressBar1.Refresh()
        System.Windows.Forms.Application.DoEvents()
    End Sub

    Private Function Deal_With_Zip_Repair_AndPretty_Print_Routine(sFileFullName As String)

        Dim sFile As String = sFileFullName
        Dim byteMatch As Match = Nothing
        Dim sFileInfo As FileInfo = New FileInfo(sFile)
        Dim corruptFileDirRetrievedFileName As String = Nothing
        Dim corruptFileXMLSubFileFullName As String = Nothing
        Dim byteMatchString As String = Nothing
        Dim extractedCorruptFileDirPath As String = Nothing
        Dim modExtractedCorruptFileDirPath As String = Nothing
        Dim modDocumentXmlFullPath As String = Nothing
        Dim sFileExtension As String = sFileInfo.Extension
        Dim sFileName As String = LCase(sFileInfo.Name)
        Dim sFileZip As String = sFile & ".zip"
        Dim zipRepairedsFileName As String = "zipRepaired_" & sFileName & ".zip"
        'Dim sFileBasePath As String = LCase(sFileInfo.Directory.ToString)
        Dim sFileBasePath As String = tempPath
        Dim zipRepairedBasePathAndFileName As String = sFileBasePath & "\" & zipRepairedsFileName
        Dim zipRepairedBasePathAndFileNameIndexLastPeriod As Integer = zipRepairedBasePathAndFileName.LastIndexOf(".")
        Dim extractedRepairedZipOutputDirectory As String = _
            zipRepairedBasePathAndFileName.Remove(zipRepairedBasePathAndFileNameIndexLastPeriod - 5)
        Dim modifiedextractedRepairedZipOutputDirectory As String = extractedRepairedZipOutputDirectory & "_mod"
        Dim wordDocumentXmlFullExtractedPath = extractedRepairedZipOutputDirectory & "\word\document.xml"
        Dim preliminaryXmlRepairedFileName As String = "preliminary_unspecified_error_repaired_" & sFileName
        Dim preliminaryXmlRepairedFullPath As String = sFileBasePath & "\" & preliminaryXmlRepairedFileName
        Dim xmlRepairedFileName As String = "unspecified_error_repaired_" & sFileName
        Dim xmlRepairedFullPath As String = sFileBasePath & "\" & xmlRepairedFileName
        Dim zipExtensionXMLRepairedFullPath As String = xmlRepairedFullPath & ".zip"
        Dim prettyPrintedDocumentXmlsFile As String = "pretty_printed_" & sFileName
        Dim prettyPrintedDocumentXmlFullPath As String = sFileBasePath & "\" & prettyPrintedDocumentXmlsFile

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        'Deletes previous versions of files from previous recoveries if they
        'are still hanging around.

        If File.Exists(sFileZip) Then

            File.Delete(sFileZip)

        End If

        'Makes target file have zip extension so zip repair will work.

        File.Copy(sFile, sFileZip, True)

        'Deletes previous versions of files from previous recoveries if they
        'are still hanging around.

        If File.Exists(zipRepairedBasePathAndFileName) Then

            File.Delete(zipRepairedBasePathAndFileName)

        End If

        'We repair the zip file in case the zip part is corrupt using a subroutine, see below.

        If Not Zip_Repair(sFileZip, zipRepairedBasePathAndFileName) = "OK" Then

            MsgBox("Your file's zip structure is unrepairable. There is apparently nothing " _
                   & "recoverable but you can try the ""Try Salvaging Just Text"", " _
                   & "choice on the Recover Menu to be sure.")

            Deal_With_Zip_Repair_AndPretty_Print_Routine = "Failed"

            Exit Function

        End If

        'MsgBox("The processing returned to the zip repair and pretty print thread after zip repair.")

        'Don't need sFileZip anymore for the moment.

        If File.Exists(sFileZip) Then

            File.Delete(sFileZip)

        End If

        'Deletes previous versions of files from previous recoveries if they
        'are still hanging around.

        If Directory.Exists(extractedRepairedZipOutputDirectory) Then

            Directory.Delete(extractedRepairedZipOutputDirectory, True)

        End If

        'Now we extract the repaired file using a subroutine, see below.

        If Not Extract_Zip(zipRepairedBasePathAndFileName, extractedRepairedZipOutputDirectory) = "OK" Then

            MsgBox("Savvy DOCX Recovery could not extract the zip structure " _
                   & "of your target file. There is apparently nothing " _
                   & "recoverable but you can try the, ""Try Salvaging Just Text" _
                   & """ choice on the Recover menu to be sure.")

            Deal_With_Zip_Repair_AndPretty_Print_Routine = "Failed"

            Exit Function

        End If

        'MsgBox("The processing returned to the zip repair and pretty print thread after zip extraction.")

        'We don't need zipRepairedBasePathAndFileName for the moment.

        If File.Exists(zipRepairedBasePathAndFileName) Then

            File.Delete(zipRepairedBasePathAndFileName)

        End If

        'Now we make the document.xml file into an normal indented XML file with many
        'lines rather than the 2 lines which is normal for document.xml. We do this
        'especially for the Unpsecified Error files as they often indicate no location data
        'other than "Line 2, coulmn 0", which is useless. Pretty printing with one tag for 
        'line does not harm the document.xml file and forces Word to indicate what line the 
        'first unspecified error now is at.

        If Not Make_Document_XML_Pretty_Print(extractedRepairedZipOutputDirectory) = "OK" Then

            MsgBox("You docx file does not contains a word\document.xml sub-file where the " _
                   & "document text is stored. Unfortunately full recovery is not possible. " _
                   & "However you can try the text extraction choice #4 on the Recover " _
                   & "menu to try extract other XML sub-files such as the footnotes.")

            Deal_With_Zip_Repair_AndPretty_Print_Routine = "Failed"

            Exit Function

        End If

        'Now we will rezip our directory with our pretty printed document.xml. See the subroutine below.                    

        If File.Exists(zipExtensionXMLRepairedFullPath) Then

            File.Delete(zipExtensionXMLRepairedFullPath)

        End If

        If Not Re_Zip(extractedRepairedZipOutputDirectory, zipExtensionXMLRepairedFullPath) = "OK" Then

            MsgBox("You docx file's zip structure was presumably repaired and the file's " _
                   & "word/document.xml subfile was prepared for providing the location " _
                   & "of the error in case it was not specified in ther opening error " _
                   & "message. However the file could not be rezipped. This is unusual. " _
                   & "Try again or try individual recovery methods from the Recover menu.")

            Deal_With_Zip_Repair_AndPretty_Print_Routine = "Failed"

            Exit Function

        End If

        'For the moment, we don't need the extracted directory. Although
        'perhap deleting and extracting again is wasteful, it helps lower 
        'the chance of left over files and folders whould a process fail.

        'MsgBox("extractedRepairedZipOutputDirectory is: " & extractedRepairedZipOutputDirectory)

        If Directory.Exists(extractedRepairedZipOutputDirectory) Then

            Directory.Delete(extractedRepairedZipOutputDirectory, True)

        End If

        'The rezipped file is copied to its pretty printed name.

        If File.Exists(prettyPrintedDocumentXmlFullPath) Then

            File.Delete(prettyPrintedDocumentXmlFullPath)

        End If

        If File.Exists(zipExtensionXMLRepairedFullPath) Then

            File.Copy(zipExtensionXMLRepairedFullPath, prettyPrintedDocumentXmlFullPath, True)

        Else

            MsgBox("Something went wrong, Savvy DOCX recovery failed to find the repaired " _
                   & "pretty printed rezipped version of your file. But you can try the, " _
                   & """Salvage Text or Data If Possible"" to recover your text.")

            Deal_With_Zip_Repair_AndPretty_Print_Routine = "Failed"

            Exit Function

        End If

        'The zip extension was necessary for I think the repair step. 7zip I think doesn't
        'care if the new file has a docx or zip extension, but anyways we don't need the 
        'the file with a zip extension anymore. I guess we could have renamed it but I'm a noob.

        If File.Exists(zipExtensionXMLRepairedFullPath) Then

            File.Delete(zipExtensionXMLRepairedFullPath)

        End If

        Deal_With_Zip_Repair_AndPretty_Print_Routine = "OK"

    End Function

    Private Function Deal_With_Unspecified_Error(sFileFullName As String)

        Dim sFile As String = sFileFullName
        Dim byteMatch As Match = Nothing
        Dim extractedCorruptFileDirInfo As DirectoryInfo
        Dim modExtractedCorruptFileDirInfo As DirectoryInfo
        Dim corruptFileDirRetrievedFilesInfoArray As FileInfo()
        Dim modCorruptFileDirRetrievedFilesInfoArray As FileInfo()
        Dim sFileInfo As FileInfo = New FileInfo(sFile)
        Dim corruptFileDirRetrievedFileInfo As FileInfo
        Dim modCorruptFileDirRetrievedFileInfo As FileInfo
        Dim corruptFileDirRetrievedFileName As String = Nothing
        Dim corruptFileXMLSubFileFullName As String = Nothing
        Dim byteMatchString As String = Nothing
        Dim extractedCorruptFileDirPath As String = Nothing
        Dim modExtractedCorruptFileDirPath As String = Nothing
        Dim modDocumentXmlFullPath As String = Nothing
        Dim sFileExtension As String = sFileInfo.Extension
        Dim sFileName As String = LCase(sFileInfo.Name)
        Dim sFileZip As String = sFile & ".zip"
        Dim zipRepairedsFileName As String = "zipRepaired_" & sFileName & ".zip"
        'Dim sFileBasePath As String = LCase(sFileInfo.Directory.ToString)
        Dim sFileBasePath As String = tempPath
        Dim zipRepairedBasePathAndFileName As String = sFileBasePath & "\" & zipRepairedsFileName
        Dim zipRepairedBasePathAndFileNameIndexLastPeriod As Integer = zipRepairedBasePathAndFileName.LastIndexOf(".")
        Dim extractedRepairedZipOutputDirectory As String = _
            zipRepairedBasePathAndFileName.Remove(zipRepairedBasePathAndFileNameIndexLastPeriod - 5)
        Dim modifiedExtractedRepairedZipOutputDirectory As String = extractedRepairedZipOutputDirectory & "_mod"
        Dim wordDocumentXmlFullExtractedPath = extractedRepairedZipOutputDirectory & "\word\document.xml"
        Dim preliminaryXmlRepairedFileName As String = "preliminary_unspecified_error_repaired_" & sFileName
        Dim preliminaryXmlRepairedFullPath As String = sFileBasePath & "\" & preliminaryXmlRepairedFileName
        Dim xmlRepairedFileName As String = "unspecified_error_repaired_" & sFileName
        Dim xmlRepairedAndTruncTreatedFileName As String = "unspecified_error__and_trunc_treated_repaired_" & sFileName
        Dim xmlRepairedFullPath As String = sFileBasePath & "\" & xmlRepairedFileName
        Dim xmlRepairedTruncTreatedFullPath As String = sFileBasePath & "\" & xmlRepairedAndTruncTreatedFileName
        Dim zipExtensionXMLRepairedFullPath As String = xmlRepairedFullPath & ".zip"
        Dim zipExtensionXMLRepairedTruncTreatedFullPath As String = xmlRepairedTruncTreatedFullPath & ".zip"
        Dim prettyPrintedDocumentXmlsFile As String = "pretty_printed_" & sFileName
        Dim prettyPrintedDocumentXmlFullPath As String = sFileBasePath & "\" & prettyPrintedDocumentXmlsFile
        Dim originalFileFullPath As String = TextBox1.Text
        Dim originalFileFullPathRepairedName As String = originalFileFullPath & ".savvy_fix.docx"
        Dim lineListFirstPass As New List(Of String)
        Dim lineListMath As New List(Of String)
        Dim lineListTable As New List(Of String)
        Dim lineListfldChar As New List(Of String)
        Dim lineListEndNote As New List(Of String)
        Dim secondLineColumnErrorLocationInteger As Integer = 0
        Dim indexBadMathTagOpenLine As Integer = 0
        Dim indexBadMathTagCloseLine As Integer = 0
        Dim indexBadTableTagOpenLine As Integer = 0
        Dim indexBadTableTagCloseLine As Integer = 0
        Dim indexBadfldCharTagOpenLine As Integer = 0
        Dim indexBadfldCharTagCloseLine As Integer = 0
        Dim indexBadEndNoteTagOpenLine As Integer = 0
        Dim indexBadEndNoteTagCloseLine As Integer = 0
        Dim numberOfCharactersInBadTags As Long = 0
        Dim numberOfCharactersFromErrorToMathTag As Integer = 0
        Dim numberOfCharactersFromErrorToTableTag As Integer = 0
        Dim numberOfCharactersFromErrorTofldCharTag As Integer = 0
        Dim numberOfCharactersFromErrorToEndNoteTag As Integer = 0
        Dim secondLineColumnErrorLocationString As String = Nothing
        Dim excisedString As String = Nothing
        Dim byteErrorLocationInteger As Integer = 0
        Dim truncatedLength As Integer = 0
        Dim intTruncationAmount As Integer = 0
        Dim truncatedLengthAsString As String = Nothing
        Dim byteErrorLocation As String = Nothing
        Dim sFileFullNameLC As String = LCase(sFileInfo.FullName)
        Dim salvagedsFileName As String = "salvaged_" & sFileName
        Dim salvagedsFileName2 As String = "salvaged_2_" & sFileName
        Dim salvagedsFileNameNoSpace As String = salvagedsFileName.Replace(" ", "_") & ".txt"
        Dim salvagedsFileNameNoSpace2 As String = salvagedsFileName2.Replace(" ", "_") & ".txt"
        Dim salvagedsFileNameAndBasePathNoSpace As String = sFileBasePath & "\" & salvagedsFileNameNoSpace
        Dim salvagedsFileNameAndBasePathNoSpace2 As String = sFileBasePath & "\" & salvagedsFileNameNoSpace2
        Dim zipRepairedsFileNameNonZipExt As String = "zipRepaired_2_" & sFileName
        Dim zipRepairedsFileNameNonZipExtNoSpace As String = zipRepairedsFileNameNonZipExt.Replace(" ", "_")
        Dim zipRepairedsFileNameNonZipExtNoSpaceFullName As String = sFileBasePath & _
            "\" & zipRepairedsFileNameNonZipExtNoSpace
        Dim normalUnspecifiedError As Boolean = True
        Dim oMathFound As Boolean
        Dim wordFileOpendOK As Boolean = False
        Dim objWordPrettyPrintedOpenAttempt As Word.Application
        Dim objDocPrettyPrintedOpenAttempt As Word.Document
        Dim objWordMath As Word.Application
        Dim objDocMath As Word.Document
        Dim objWordTable As Word.Application
        Dim objDocTable As Word.Document
        Dim objWordfldChar As Word.Application
        Dim objDocfldChar As Word.Document
        Dim objWordXMLProcessingCompleted As Word.Application
        Dim objDocXMLProcessingCompleted As Word.Document
        Dim objWordTruncationSalvagingCompleted As Word.Application
        Dim objDocTruncationSalvagingCompleted As Word.Document
        Dim objWordEndNote As Word.Application
        Dim objDocEndNote As Word.Document
        Dim xT As Integer = 0

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        If Not Deal_With_Zip_Repair_AndPretty_Print_Routine(sFile) = "OK" Then

            Deal_With_Unspecified_Error = "Failed"
            Exit Function

        End If

        'With the pretty printed document.xml rezipped Word file now we are going to use
        'Word to tell us which XML line is where the first bad tag is.

        objWordPrettyPrintedOpenAttempt = New Word.Application
        objDocPrettyPrintedOpenAttempt = New Word.Document

        Try

            objDocPrettyPrintedOpenAttempt = objWordPrettyPrintedOpenAttempt.Documents.OpenNoRepairDialog(prettyPrintedDocumentXmlFullPath, _
                                                                                                          OpenAndRepair:=True)
            objWordPrettyPrintedOpenAttempt.Visible = True
            objWordPrettyPrintedOpenAttempt.WindowState = Word.WdWindowState.wdWindowStateMaximize
            objWordPrettyPrintedOpenAttempt.Activate()

            'MsgBox("That's weird, the file shouldn't open at this point. We have only altered it to show what line an error is at.")

            Me.Cursor = System.Windows.Forms.Cursors.Default
            'ProgressBar1.Hide()

            If File.Exists(prettyPrintedDocumentXmlFullPath) Then

                File.Delete(prettyPrintedDocumentXmlFullPath)

            End If

            Deal_With_Unspecified_Error = "OK"

            'CloseWordApplicationWithDocument(objWordPrettyPrintedOpenAttempt, objDocPrettyPrintedOpenAttempt)

            Exit Function

        Catch ex_Pretty_Printed_First_Pass_Error As Exception

            CloseWordApplicationWithDocument(objWordPrettyPrintedOpenAttempt, objDocPrettyPrintedOpenAttempt)

            byteMatchString = ex_Pretty_Printed_First_Pass_Error.ToString

            If File.Exists(zipRepairedBasePathAndFileName) Then

                File.Delete(zipRepairedBasePathAndFileName)

            End If

            File.Copy(prettyPrintedDocumentXmlFullPath, zipRepairedBasePathAndFileName, True)

            'If File.Exists(prettyPrintedDocumentXmlFullPath) Then

            '    File.Delete(prettyPrintedDocumentXmlFullPath)

            'End If

            'ProgressBar1.Value

            'Now we extract the repaired file again, now that we know which line the error 
            'is on. This is for unspecified error without error character or line numbers

            If Directory.Exists(extractedRepairedZipOutputDirectory) Then

                Directory.Delete(extractedRepairedZipOutputDirectory, True)

            End If

            If Not Extract_Zip(zipRepairedBasePathAndFileName, extractedRepairedZipOutputDirectory) = "OK" Then

                MsgBox("Savvy DOCX Recovery could not extract the zip structure " _
                       & "of your target file. There is apparently nothing " _
                       & "recoverable but you can try Method 4, ""Salvage Text " _
                       & "or Data If Possible"", on the Recover Menu to be sure.")

                Deal_With_Unspecified_Error = "Failed"

                Exit Function

            End If

            If Directory.Exists(modifiedExtractedRepairedZipOutputDirectory) Then

                Directory.Delete(modifiedExtractedRepairedZipOutputDirectory, True)

            End If

            'We need a second extracted version I think because I found errors
            'were generated about files being in use and such. So we will work
            'on virgin subfile in one directory and copy the fixed version over
            'to the "modified" folder from where we will rezip our final fixed product.

            If Not Extract_Zip(zipRepairedBasePathAndFileName, modifiedExtractedRepairedZipOutputDirectory) = "OK" Then

                MsgBox("Savvy DOCX Recovery could not extract the zip structure " _
                       & "of your target file. There is apparently nothing " _
                       & "recoverable but you can try Method 4, ""Salvage Text " _
                       & "or Data If Possible"", on the Recover Menu to be sure.")

                Deal_With_Unspecified_Error = "Failed"

                Exit Function

            End If

            'ProgressBar1.Value= 30

            extractedCorruptFileDirPath = extractedRepairedZipOutputDirectory & "\"
            extractedCorruptFileDirInfo = New DirectoryInfo(extractedCorruptFileDirPath)

            'Below gives us an array of all the files in our extracted corrupt target.

            corruptFileDirRetrievedFilesInfoArray = _
                extractedCorruptFileDirInfo.GetFiles("*.*", SearchOption.AllDirectories)

            'ProgressBar1.Value= 40

            'These lines help us move the progress bar over its next 50 increments divided
            'by the number of XML subfiles we will process. If an XML file is found corrupt,
            'we will try to repair it.

            Dim extractedCorruptFileDirInfoArrayCount As Integer = _
                corruptFileDirRetrievedFilesInfoArray.GetLength(0)
            Dim progressBarIncrement As Integer = 50 \ (extractedCorruptFileDirInfoArrayCount + 1)

            'If there are more than 100 subfiles, the increment will be less that 0.5 
            'which will be rounded to the integer 0, so we need to make the increment
            'artificially equal to 1 in those cases.

            If progressBarIncrement = 0 Then

                progressBarIncrement = 1

            End If

            'Now we will process each subfile in unzipped collection. If the file is an 
            'XML .xml or XML .rels file we will check if it is well formed XML and try to
            'fix it with xmllint if it is not.

            For Each corruptFileDirRetrievedFileInfo In corruptFileDirRetrievedFilesInfoArray

                If InStr(corruptFileDirRetrievedFileInfo.Extension, ".xml") Or _
                    InStr(corruptFileDirRetrievedFileInfo.Extension, ".rels") Then

                    corruptFileDirRetrievedFileName = corruptFileDirRetrievedFileInfo.Name()
                    corruptFileXMLSubFileFullName = corruptFileDirRetrievedFileInfo.FullName
                    modDocumentXmlFullPath = modifiedExtractedRepairedZipOutputDirectory & _
                        corruptFileXMLSubFileFullName.Substring(extractedRepairedZipOutputDirectory.Length)

                    'MsgBox("corruptFileXMLSubFileFullName is: " & corruptFileXMLSubFileFullName)
                    'If we artificially gave the progress bar an increment of 1 and there
                    'are more than 50 sub files to check, we will need to keep the
                    'progress bar at 90.

                    If ProgressBar1.Value + progressBarIncrement > 90 Then

                        'ProgressBar1.Value= 90

                    Else

                        'ProgressBar1.Value=ProgressBar1.Value+ progressBarIncrement

                    End If

                    'Code which just deals with the all important doument.xml file where the
                    'errors are found in unpsecified error instances. We actually don't
                    'even process any other XML file if there an Unspecified error as the 
                    'problem is overwhelmingly just with the document.xml in these instances.
                    'With other errors will check all the xml sub-files, but not here for the
                    'moment. At least that is my thinking now. No need to risk messing up the recovery.
                    'MsgBox("corruptFileXMLSubFileFullName is: " & corruptFileXMLSubFileFullName _
                    '& " wordDocumentXmlFullExtractedPath is: " & wordDocumentXmlFullExtractedPath)

                    'MsgBox("ex_Pretty_Printed_First_Pass_Error.ToString is: " & _
                    'ex_Pretty_Printed_First_Pass_Error.ToString)

                    If corruptFileXMLSubFileFullName = wordDocumentXmlFullExtractedPath Then

                        'We use regex to isolate the line from the ex_Pretty_Printed_First_Pass_Error string error.

                        byteMatch = Regex.Match(ex_Pretty_Printed_First_Pass_Error.Message, ("Line: [0-9]+"))
                        byteMatchString = byteMatch.ToString
                        secondLineColumnErrorLocationString = DelFromLeft("Line: ", byteMatchString)

                        Integer.TryParse(secondLineColumnErrorLocationString, secondLineColumnErrorLocationInteger)

                        Using sr As StreamReader = New StreamReader(corruptFileXMLSubFileFullName)

                            Do While sr.Peek() >= 0
                                lineListFirstPass.Add(sr.ReadLine())
                            Loop

                            sr.Close()

                        End Using

                        'If Word tells us the problem is centered around math tags, that's good because it's
                        'easy to excise the offending math tag section and display what has been removed
                        'for recreation of what is usually a math formula.

                        indexBadMathTagOpenLine = lineListFirstPass.LastIndexOf("<m:oMath>", _
                                                secondLineColumnErrorLocationInteger) + 0
                        indexBadMathTagCloseLine = lineListFirstPass.IndexOf("</m:oMath>", _
                                                secondLineColumnErrorLocationInteger) + 1
                        indexBadTableTagOpenLine = lineListFirstPass.LastIndexOf("<w:tbl>", _
                                                secondLineColumnErrorLocationInteger) + 0
                        indexBadTableTagCloseLine = lineListFirstPass.IndexOf("</w:tbl>", _
                                                secondLineColumnErrorLocationInteger) + 1
                        indexBadfldCharTagOpenLine = lineListFirstPass.LastIndexOf("<w:fldChar w:fldCharType=""begin""/>", _
                                                secondLineColumnErrorLocationInteger) + 0
                        indexBadfldCharTagCloseLine = lineListFirstPass.IndexOf("<w:fldChar w:fldCharType=""end""/>", _
                                                secondLineColumnErrorLocationInteger) + 1
                        indexBadEndNoteTagOpenLine = lineListFirstPass.LastIndexOf("<w:instrText xml:space=""preserve""> ADDIN EN.CITE <EndNote>", _
                                                secondLineColumnErrorLocationInteger) + 0
                        indexBadEndNoteTagCloseLine = lineListFirstPass.IndexOf("</EndNote>", _
                                                secondLineColumnErrorLocationInteger) + 2
                        numberOfCharactersFromErrorToMathTag = secondLineColumnErrorLocationInteger - _
                                                indexBadMathTagOpenLine
                        numberOfCharactersFromErrorToTableTag = secondLineColumnErrorLocationInteger - _
                                                indexBadTableTagOpenLine
                        numberOfCharactersFromErrorTofldCharTag = secondLineColumnErrorLocationInteger - _
                                                indexBadfldCharTagOpenLine
                        numberOfCharactersFromErrorToEndNoteTag = secondLineColumnErrorLocationInteger - _
                                                indexBadEndNoteTagOpenLine

                        'MsgBox("indexBadEndNoteTagOpenLine is: " & indexBadEndNoteTagOpenLine.ToString & _
                        '      " and indexBadEndNoteTagCloseLine is: " & indexBadEndNoteTagCloseLine.ToString & _
                        '      " and numberOfCharactersFromErrorToEndNoteTag is: " & numberOfCharactersFromErrorToEndNoteTag.ToString)

                        If indexBadMathTagOpenLine = -1 Then

                            numberOfCharactersFromErrorToMathTag = 999999995

                        End If
                        'MsgBox("indexBadMathTagOpenLine is: " & indexBadMathTagOpenLine)
                        'MsgBox("indexBadMathTagCloseLine is: " & indexBadMathTagCloseLine)
                        If indexBadTableTagOpenLine = -1 Then

                            numberOfCharactersFromErrorToTableTag = 999999996

                        End If

                        If indexBadfldCharTagOpenLine = -1 Then

                            numberOfCharactersFromErrorTofldCharTag = 999999997

                        End If

                        If indexBadEndNoteTagOpenLine = -1 Then

                            numberOfCharactersFromErrorToEndNoteTag = 999999998

                        End If

                        Dim vals As Integer() = {numberOfCharactersFromErrorToMathTag, _
                                                 numberOfCharactersFromErrorToTableTag, _
                                                 numberOfCharactersFromErrorTofldCharTag, _
                                                 numberOfCharactersFromErrorToEndNoteTag}

                        'I appeared to have copied this code for the next section:
                        'http://www.dotnetperls.com/math-max-min-vbnet

                        Dim largest As Integer = Integer.MinValue
                        Dim smallest As Integer = Integer.MaxValue

                        For Each element As Integer In vals
                            smallest = Math.Min(smallest, element)
                        Next

                        If numberOfCharactersFromErrorToMathTag = smallest Then

                            oMathFound = True

                            'Clears our text box from possible previous resultss.

                            RichTextBox1.Clear()

                            'Hopefully from the code above we know the line numbers of the bad closing tag
                            'and the bad opening one. The difference indicating the number of lines needing 
                            'to be excised starting at the opening bad tag.

                            numberOfCharactersInBadTags = indexBadMathTagCloseLine - indexBadMathTagOpenLine
                            'MsgBox("numberOfCharactersInBadTags is: " & numberOfCharactersInBadTags)
                            'Here we write the soon to be excised lines to our text box.

                            For Each excisedString In lineListFirstPass.GetRange(indexBadMathTagOpenLine, numberOfCharactersInBadTags)

                                RichTextBox1.AppendText(excisedString)

                            Next

                            'Here is where our bad lines are actually removed from our list of tags
                            'that makes up the document.xml file. The second line of code just rewrites
                            'the now hopefully good XML back to replace the bad in the file. However,
                            'note we are replacing the XML in the modified extracted directory, not the original.

                            lineListFirstPass.RemoveRange(indexBadMathTagOpenLine, numberOfCharactersInBadTags)
                            File.WriteAllLines(modDocumentXmlFullPath, lineListFirstPass.ToArray)
                            normalUnspecifiedError = True
                            lineListFirstPass.Clear()

                            'ProgressBar1.Value= 95

                            Do Until wordFileOpendOK = True

                                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                                    File.Delete(zipExtensionXMLRepairedFullPath)

                                End If

                                Re_Zip(modifiedExtractedRepairedZipOutputDirectory, zipExtensionXMLRepairedFullPath)

                                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                                    File.Copy(zipExtensionXMLRepairedFullPath, preliminaryXmlRepairedFullPath, True)

                                End If

                                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                                    File.Delete(zipExtensionXMLRepairedFullPath)

                                End If

                                objWordMath = New Word.Application
                                objDocMath = New Word.Document

                                Try

                                    'The next few lines bypass the section where the corrupt 
                                    'file is copied to one whose name is a random number.

                                    'If File.Exists(originalFileFullPathRepairedName) Then

                                    '    File.Delete(originalFileFullPathRepairedName)

                                    'End If

                                    'File.Copy(preliminaryXmlRepairedFullPath, originalFileFullPathRepairedName, True)
                                    objWordMath.Visible = False
                                    objDocMath = objWordMath.Documents.OpenNoRepairDialog(preliminaryXmlRepairedFullPath, OpenAndRepair:=True)
                                    objWordMath.Visible = False
                                    objWordMath.Documents(preliminaryXmlRepairedFullPath).Close(Word.WdSaveOptions.wdDoNotSaveChanges)
                                    'objWordMath.WindowState = Word.WdWindowState.wdWindowStateMaximize
                                    'objWordMath.Activate()

                                    'Me.Cursor = System.Windows.Forms.Cursors.Default

                                    'If File.Exists(preliminaryXmlRepairedFullPath) Then

                                    '    File.Delete(preliminaryXmlRepairedFullPath)

                                    'End If

                                    'If ProgressBar1.Visible = True Then

                                    '    Me.Cursor = System.Windows.Forms.Cursors.Default
                                    '    'ProgressBar1.Hide()

                                    'End If

                                    'MsgBox("Although this program may have been able to recover some of your document with " _
                                    '    & "formatting, there may be additional text which is salvageable without formatting " _
                                    '    & "or with the formatting limited to the types of which text files are capable. This " _
                                    '    & "salvaging process uses SilverCoder's command line program, DocToText and can be" _
                                    '    & "started by choosing the ""Try Salvaging Just Text"" choice  on the Recover menu.")

                                    'Deal_With_Unspecified_Error = "OK"
                                    wordFileOpendOK = True

                                    'Exit Function

                                Catch ex_math As Exception

                                    CloseWordApplicationWithDocument(objWordMath, objDocMath)

                                    'We need to delete zipExtensionXMLRepairedFullPath 
                                    'because we will be getting new versions soon

                                    If ex_math.Message.Contains("Unspecified error") Then

                                        byteMatch = Regex.Match(ex_math.Message, ("Line: [0-9]+"))
                                        byteMatchString = byteMatch.ToString
                                        secondLineColumnErrorLocationString = DelFromLeft("Line: ", byteMatchString)

                                        Integer.TryParse(secondLineColumnErrorLocationString, secondLineColumnErrorLocationInteger)

                                        Using sr As StreamReader = New StreamReader(modDocumentXmlFullPath)

                                            Do While sr.Peek() >= 0
                                                lineListMath.Add(sr.ReadLine())
                                            Loop

                                            sr.Close()

                                        End Using

                                        'So this is the 2nd round of excising Math tags

                                        indexBadMathTagOpenLine = lineListMath.LastIndexOf("<m:oMath>", secondLineColumnErrorLocationInteger) + 0
                                        indexBadMathTagCloseLine = lineListMath.IndexOf("</m:oMath>", secondLineColumnErrorLocationInteger) + 1

                                        numberOfCharactersInBadTags = indexBadMathTagCloseLine - indexBadMathTagOpenLine

                                        If (indexBadMathTagOpenLine < 0) Then
                                            indexBadMathTagOpenLine = 0
                                        End If

                                        If (numberOfCharactersInBadTags < 0) Then
                                            numberOfCharactersInBadTags = 1
                                        End If

                                        'MsgBox("indexBadMathTagOpenLine is: " & indexBadMathTagOpenLine.ToString _
                                        '& " indexBadMathTagCloseLine is: " & indexBadMathTagCloseLine.ToString _
                                        '& " numberOfCharactersInBadTags is: " & numberOfCharactersInBadTags.ToString _
                                        '& " secondLineColumnErrorLocationInteger is " & secondLineColumnErrorLocationInteger.ToString)
                                        'Here we write the soon to be excised lines to our text box.

                                        For Each excisedString In lineListMath.GetRange(indexBadMathTagOpenLine, numberOfCharactersInBadTags)

                                            RichTextBox1.AppendText(excisedString)

                                        Next

                                        lineListMath.RemoveRange(indexBadMathTagOpenLine, numberOfCharactersInBadTags)
                                        File.WriteAllLines(modDocumentXmlFullPath, lineListMath.ToArray)
                                        lineListMath.Clear()

                                    ElseIf ex_math.Message.Contains("The name in the end tag of the " _
                                            & "element must match the element type in the start tag.") Then

                                        'Subroutine, see below.

                                        'MsgBox("We seem to have corrected the ""Unspecified"" error(s) but there are " _
                                        '& "or is additional Name_Tag ending error(s) we now need to address for the file to open.")
                                        'MsgBox("extractedRepairedZipOutputDirectory is: " & extractedRepairedZipOutputDirectory)

                                        Deal_With_Name_Tag_Error_Detail(preliminaryXmlRepairedFullPath, _
                                                extractedRepairedZipOutputDirectory, modifiedExtractedRepairedZipOutputDirectory)

                                    Else

                                        'Subroutine, see below.

                                        'MsgBox("We seem to have corrected the ""Unspecified"" error(s) but there are " _
                                        '& "additional kinds we now need to address for the file to open.")
                                        If (Deal_With_Everything_Else_After_Unspecified_Error_Fix(preliminaryXmlRepairedFullPath, _
                                                    extractedRepairedZipOutputDirectory, modifiedExtractedRepairedZipOutputDirectory) = "OK") Then
                                            wordFileOpendOK = True
                                            Deal_With_Unspecified_Error = "OK"
                                            Exit Function
                                        Else
                                            Deal_With_Unspecified_Error = "Failed"
                                            Exit Function
                                        End If

                                    End If

                                Finally

                                    GC.Collect()
                                    GC.WaitForPendingFinalizers()

                                    ' GC needs to be called twice in order to get the Finalizers called  
                                    ' - the first time in, it simply makes a list of what is to be  
                                    ' finalized, the second time in, it actually is finalizing. Only  
                                    ' then will the object do its automatic ReleaseComObject. 

                                    GC.Collect()
                                    GC.WaitForPendingFinalizers()

                                End Try

                            Loop

                        ElseIf numberOfCharactersFromErrorToTableTag = smallest Then

                            oMathFound = False

                            'Clears our text box from possible previous resultss.

                            RichTextBox1.Clear()

                            'Hopefully from the code above we know the line numbers of the bad closing tag
                            'and the bad opening one. The difference indicating the number of lines needing 
                            'to be excised starting at the opening bad tag.

                            numberOfCharactersInBadTags = indexBadTableTagCloseLine - indexBadTableTagOpenLine

                            'Here we write the soon to be excised lines to our text box.

                            For Each excisedString In lineListFirstPass.GetRange(indexBadTableTagOpenLine, numberOfCharactersInBadTags)

                                RichTextBox1.AppendText(excisedString)

                            Next

                            'Here is where our bad lines are actually removed from our list of tags
                            'that makes up the document.xml file. The second line of code just rewrites
                            'the now hopefully good XML back to replace the bad in the file. However,
                            'note we are replacing the XML in the modified extracted directory, not the original.

                            lineListFirstPass.RemoveRange(indexBadTableTagOpenLine, numberOfCharactersInBadTags)
                            File.WriteAllLines(modDocumentXmlFullPath, lineListFirstPass.ToArray)
                            normalUnspecifiedError = True
                            lineListFirstPass.Clear()

                            'ProgressBar1.Value= 95

                            Do Until wordFileOpendOK = True

                                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                                    File.Delete(zipExtensionXMLRepairedFullPath)

                                End If

                                Re_Zip(modifiedExtractedRepairedZipOutputDirectory, zipExtensionXMLRepairedFullPath)

                                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                                    File.Copy(zipExtensionXMLRepairedFullPath, preliminaryXmlRepairedFullPath, True)

                                End If

                                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                                    File.Delete(zipExtensionXMLRepairedFullPath)

                                End If

                                objWordTable = New Word.Application
                                objDocTable = New Word.Document

                                Try

                                    'The next few lines bypass the section where the corrupt 
                                    'file is copied to one whose name is a random number.

                                    'If File.Exists(originalFileFullPathRepairedName) Then

                                    '    File.Delete(originalFileFullPathRepairedName)

                                    'End If

                                    'File.Copy(preliminaryXmlRepairedFullPath, originalFileFullPathRepairedName, True)
                                    objWordTable.Visible = False
                                    objDocTable = objWordTable.Documents.OpenNoRepairDialog(preliminaryXmlRepairedFullPath, OpenAndRepair:=True)
                                    objDocTable.Visible = False
                                    objWordTable.Documents(preliminaryXmlRepairedFullPath).Close(Word.WdSaveOptions.wdDoNotSaveChanges)
                                    'objWordTable.WindowState = Word.WdWindowState.wdWindowStateMaximize
                                    'objWordTable.Activate()

                                    'If File.Exists(preliminaryXmlRepairedFullPath) Then

                                    '    File.Delete(preliminaryXmlRepairedFullPath)

                                    'End If

                                    'Me.Cursor = System.Windows.Forms.Cursors.Default

                                    'If ProgressBar1.Visible = True Then

                                    '    Me.Cursor = System.Windows.Forms.Cursors.Default
                                    '    'ProgressBar1.Hide()

                                    'End If

                                    'MsgBox("Although this program may have been able to recover some of your document " _
                                    '    & "there may be additional text which is salvageable without formatting using " _
                                    '    & "SilverCoder's command line program, DocToText. Please choose ""Method 4 - " _
                                    '    & "Salvage Text or Data If Possible"" choice if desired on the menu")

                                    'Deal_With_Unspecified_Error = "OK"

                                    wordFileOpendOK = True

                                    'Exit Function

                                Catch ex_table As Exception

                                    CloseWordApplicationWithDocument(objWordTable, objDocTable)

                                    'We need to delete zipExtensionXMLRepairedFullPath and preliminaryXmlRepairedFullPath 
                                    'because we will be getting new versions soon.

                                    If ex_table.Message.Contains("Unspecified error") Then

                                        byteMatch = Regex.Match(ex_table.Message, ("Line: [0-9]+"))
                                        byteMatchString = byteMatch.ToString
                                        secondLineColumnErrorLocationString = DelFromLeft("Line: ", byteMatchString)

                                        Integer.TryParse(secondLineColumnErrorLocationString, secondLineColumnErrorLocationInteger)

                                        Using sr As StreamReader = New StreamReader(modDocumentXmlFullPath)

                                            Do While sr.Peek() >= 0
                                                lineListTable.Add(sr.ReadLine())
                                            Loop

                                            sr.Close()

                                        End Using

                                        'So this is the 2nd round of excising Math tags

                                        indexBadTableTagOpenLine = lineListTable.LastIndexOf("<w:tbl>", secondLineColumnErrorLocationInteger) + 0
                                        indexBadTableTagCloseLine = lineListTable.IndexOf("</w:tbl>", secondLineColumnErrorLocationInteger) + 1

                                        numberOfCharactersInBadTags = indexBadTableTagCloseLine - indexBadTableTagOpenLine

                                        'Here we write the soon to be excised lines to our text box.

                                        For Each excisedString In lineListTable.GetRange(indexBadTableTagOpenLine, numberOfCharactersInBadTags)

                                            RichTextBox1.AppendText(excisedString)

                                        Next

                                        lineListTable.RemoveRange(indexBadTableTagOpenLine, numberOfCharactersInBadTags)
                                        File.WriteAllLines(modDocumentXmlFullPath, lineListTable.ToArray)
                                        lineListTable.Clear()

                                    ElseIf ex_table.Message.Contains("The name in the end tag of the element must match the element type in the start tag.") Then

                                        'Subroutine, see below.

                                        'MsgBox("We seem to have corrected the ""Unspecified"" error(s) but there are " _
                                        '& "additional kinds we now need to address for the file to open.")

                                        Deal_With_Name_Tag_Error_Detail(preliminaryXmlRepairedFullPath, _
                                                    extractedRepairedZipOutputDirectory, modifiedExtractedRepairedZipOutputDirectory)

                                    Else

                                        'Subroutine, see below.

                                        'MsgBox("We seem to have corrected the ""Unspecified"" error(s) but there are " _
                                        '& "additional kinds we now need to address for the file to open.")
                                        If (Deal_With_Everything_Else_After_Unspecified_Error_Fix(preliminaryXmlRepairedFullPath, _
                                                    extractedRepairedZipOutputDirectory, modifiedExtractedRepairedZipOutputDirectory) = "OK") Then
                                            wordFileOpendOK = True
                                            Deal_With_Unspecified_Error = "OK"
                                            Exit Function
                                        End If
                                    End If

                                Finally

                                    GC.Collect()
                                    GC.WaitForPendingFinalizers()

                                    ' GC needs to be called twice in order to get the Finalizers called  
                                    ' - the first time in, it simply makes a list of what is to be  
                                    ' finalized, the second time in, it actually is finalizing. Only  
                                    ' then will the object do its automatic ReleaseComObject. 

                                    GC.Collect()
                                    GC.WaitForPendingFinalizers()

                                End Try

                            Loop

                        ElseIf numberOfCharactersFromErrorTofldCharTag = smallest Then

                            oMathFound = False

                            'Clears our text box from possible previous resultss.

                            RichTextBox1.Clear()

                            'Hopefully from the code above we know the line numbers of the bad closing tag
                            'and the bad opening one. The difference indicating the number of lines needing 
                            'to be excised starting at the opening bad tag.

                            numberOfCharactersInBadTags = indexBadfldCharTagCloseLine - indexBadfldCharTagOpenLine

                            'Here we write the soon to be excised lines to our text box.

                            For Each excisedString In lineListFirstPass.GetRange(indexBadfldCharTagOpenLine, numberOfCharactersInBadTags)

                                RichTextBox1.AppendText(excisedString)

                            Next

                            'Here is where our bad lines are actually removed from our list of tags
                            'that makes up the document.xml file. The second line of code just rewrites
                            'the now hopefully good XML back to replace the bad in the file. However,
                            'note we are replacing the XML in the modified extracted directory, not the original.

                            lineListFirstPass.RemoveRange(indexBadfldCharTagOpenLine, numberOfCharactersInBadTags)
                            File.WriteAllLines(modDocumentXmlFullPath, lineListFirstPass.ToArray)
                            normalUnspecifiedError = True
                            lineListFirstPass.Clear()

                            'If File.Exists(zipExtensionXMLRepairedFullPath) Then

                            '    File.Delete(zipExtensionXMLRepairedFullPath)

                            'End If

                            'Re_Zip(modifiedExtractedRepairedZipOutputDirectory, zipExtensionXMLRepairedFullPath)

                            'If File.Exists(zipExtensionXMLRepairedFullPath) Then

                            '    File.Copy(zipExtensionXMLRepairedFullPath, preliminaryXmlRepairedFullPath, True)

                            'End If

                            'If File.Exists(zipExtensionXMLRepairedFullPath) Then

                            '    File.Delete(zipExtensionXMLRepairedFullPath)

                            'End If

                            'ProgressBar1.Value= 95

                            Do Until wordFileOpendOK = True

                                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                                    File.Delete(zipExtensionXMLRepairedFullPath)

                                End If

                                Re_Zip(modifiedExtractedRepairedZipOutputDirectory, zipExtensionXMLRepairedFullPath)

                                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                                    File.Copy(zipExtensionXMLRepairedFullPath, preliminaryXmlRepairedFullPath, True)

                                End If

                                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                                    File.Delete(zipExtensionXMLRepairedFullPath)

                                End If

                                objWordfldChar = New Word.Application
                                objDocfldChar = New Word.Document

                                Try

                                    'The next few lines bypass the section where the corrupt 
                                    'file is copied to one whose name is a random number.

                                    'If File.Exists(originalFileFullPathRepairedName) Then

                                    '    File.Delete(originalFileFullPathRepairedName)

                                    'End If

                                    'File.Copy(preliminaryXmlRepairedFullPath, originalFileFullPathRepairedName, True)
                                    objWordfldChar.Visible = False
                                    objDocfldChar = objWordfldChar.Documents.OpenNoRepairDialog(preliminaryXmlRepairedFullPath, OpenAndRepair:=True)
                                    'objWordfldChar.WindowState = Word.WdWindowState.wdWindowStateMaximize
                                    objWordfldChar.Visible = False
                                    objWordfldChar.Documents(preliminaryXmlRepairedFullPath).Close(Word.WdSaveOptions.wdDoNotSaveChanges)

                                    'objWordfldChar.Activate()

                                    'Me.Cursor = System.Windows.Forms.Cursors.Default
                                    'ProgressBar1.Hide()

                                    'If File.Exists(preliminaryXmlRepairedFullPath) Then

                                    '    File.Delete(preliminaryXmlRepairedFullPath)

                                    'End If

                                    'MsgBox("Although this program may have been able to recover some of your document " _
                                    '    & "there may be additional text which is salvageable without formatting using " _
                                    '    & "SilverCoder's command line program, DocToText. Please choose ""Method 4 - " _
                                    '    & "Salvage Text or Data If Possible"" choice if desired on the menu")

                                    'Deal_With_Unspecified_Error = "OK"

                                    wordFileOpendOK = True

                                    'Exit Function

                                Catch ex_fldChar As Exception

                                    CloseWordApplicationWithDocument(objWordfldChar, objDocfldChar)

                                    'We need to delete zipExtensionXMLRepairedFullPath 
                                    'because we will be getting new versions soon

                                    If ex_fldChar.Message.Contains("Unspecified error") Then

                                        byteMatch = Regex.Match(ex_fldChar.Message, ("Line: [0-9]+"))
                                        byteMatchString = byteMatch.ToString
                                        secondLineColumnErrorLocationString = DelFromLeft("Line: ", byteMatchString)

                                        Integer.TryParse(secondLineColumnErrorLocationString, secondLineColumnErrorLocationInteger)

                                        Using sr As StreamReader = New StreamReader(modDocumentXmlFullPath)

                                            Do While sr.Peek() >= 0
                                                lineListfldChar.Add(sr.ReadLine())
                                            Loop

                                            sr.Close()

                                        End Using

                                        'So this is the 2nd round of excising Math tags

                                        indexBadfldCharTagOpenLine = lineListfldChar.LastIndexOf("<w:fldChar w:fldCharType=""begin""/>", _
                                                                                              secondLineColumnErrorLocationInteger) + 0
                                        indexBadfldCharTagCloseLine = lineListfldChar.IndexOf("<w:fldChar w:fldCharType=""end""/>", _
                                                                                           secondLineColumnErrorLocationInteger) + 1

                                        numberOfCharactersInBadTags = indexBadfldCharTagCloseLine - indexBadfldCharTagOpenLine

                                        'Here we write the soon to be excised lines to our text box.

                                        For Each excisedString In lineListfldChar.GetRange(indexBadfldCharTagOpenLine, numberOfCharactersInBadTags)

                                            RichTextBox1.AppendText(excisedString)

                                        Next

                                        lineListfldChar.RemoveRange(indexBadfldCharTagOpenLine, numberOfCharactersInBadTags)
                                        File.WriteAllLines(modDocumentXmlFullPath, lineListfldChar.ToArray)
                                        lineListfldChar.Clear()

                                        'If File.Exists(zipExtensionXMLRepairedFullPath) Then

                                        '    File.Delete(zipExtensionXMLRepairedFullPath)

                                        'End If

                                        'Re_Zip(modifiedExtractedRepairedZipOutputDirectory, zipExtensionXMLRepairedFullPath)

                                        'If File.Exists(zipExtensionXMLRepairedFullPath) Then

                                        '    File.Copy(zipExtensionXMLRepairedFullPath, preliminaryXmlRepairedFullPath, True)

                                        'End If

                                        'If File.Exists(zipExtensionXMLRepairedFullPath) Then

                                        '    File.Delete(zipExtensionXMLRepairedFullPath)

                                        'End If

                                    ElseIf ex_fldChar.Message.Contains("The name in the end tag of the element must match the element type in the start tag.") Then

                                        'Subroutine, see below.

                                        'MsgBox("We seem to have corrected the ""Unspecified"" error(s) but there are " _
                                        '& "additional kinds we now need to address for the file to open.")

                                        Deal_With_Name_Tag_Error_Detail(preliminaryXmlRepairedFullPath, _
                                                    extractedRepairedZipOutputDirectory, modifiedExtractedRepairedZipOutputDirectory)

                                    Else

                                        'Subroutine, see below.

                                        'MsgBox("We seem to have corrected the ""Unspecified"" error(s) but there are " _
                                        '& "additional kinds we now need to address for the file to open.")

                                        If (Deal_With_Everything_Else_After_Unspecified_Error_Fix(preliminaryXmlRepairedFullPath, _
                                                    extractedRepairedZipOutputDirectory, modifiedExtractedRepairedZipOutputDirectory) = "OK") Then
                                            wordFileOpendOK = True
                                            Deal_With_Unspecified_Error = "OK"
                                            Exit Function
                                        End If

                                    End If

                                Finally

                                    GC.Collect()
                                    GC.WaitForPendingFinalizers()

                                    ' GC needs to be called twice in order to get the Finalizers called  
                                    ' - the first time in, it simply makes a list of what is to be  
                                    ' finalized, the second time in, it actually is finalizing. Only  
                                    ' then will the object do its automatic ReleaseComObject. 

                                    GC.Collect()
                                    GC.WaitForPendingFinalizers()

                                End Try

                            Loop

                        ElseIf numberOfCharactersFromErrorToEndNoteTag = smallest Then

                            oMathFound = False

                            'Clears our text box from possible previous resultss.

                            RichTextBox1.Clear()

                            'Hopefully from the code above we know the line numbers of the bad closing tag
                            'and the bad opening one. The difference indicating the number of lines needing 
                            'to be excised starting at the opening bad tag.

                            numberOfCharactersInBadTags = indexBadEndNoteTagCloseLine - indexBadEndNoteTagOpenLine

                            'Here we write the soon to be excised lines to our text box.

                            For Each excisedString In lineListFirstPass.GetRange(indexBadEndNoteTagOpenLine, numberOfCharactersInBadTags)

                                RichTextBox1.AppendText(excisedString)

                            Next

                            'Here is where our bad lines are actually removed from our list of tags
                            'that makes up the document.xml file. The second line of code just rewrites
                            'the now hopefully good XML back to replace the bad in the file. However,
                            'note we are replacing the XML in the modified extracted directory, not the original.

                            lineListFirstPass.RemoveRange(indexBadEndNoteTagOpenLine, numberOfCharactersInBadTags)
                            File.WriteAllLines(modDocumentXmlFullPath, lineListFirstPass.ToArray)
                            normalUnspecifiedError = True
                            lineListFirstPass.Clear()

                            'ProgressBar1.Value= 95

                            Do Until wordFileOpendOK = True

                                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                                    File.Delete(zipExtensionXMLRepairedFullPath)

                                End If

                                Re_Zip(modifiedExtractedRepairedZipOutputDirectory, zipExtensionXMLRepairedFullPath)

                                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                                    File.Copy(zipExtensionXMLRepairedFullPath, preliminaryXmlRepairedFullPath, True)

                                End If

                                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                                    File.Delete(zipExtensionXMLRepairedFullPath)

                                End If

                                objWordEndNote = New Word.Application
                                objDocEndNote = New Word.Document

                                Try

                                    'The next few lines bypass the section where the corrupt 
                                    'file is copied to one whose name is a random number.

                                    'If File.Exists(originalFileFullPathRepairedName) Then

                                    '    File.Delete(originalFileFullPathRepairedName)

                                    'End If

                                    'File.Copy(preliminaryXmlRepairedFullPath, originalFileFullPathRepairedName, True)
                                    objWordEndNote.Visible = False
                                    objDocEndNote = objWordEndNote.Documents.OpenNoRepairDialog(preliminaryXmlRepairedFullPath, OpenAndRepair:=True)
                                    'objWordEndNote.WindowState = Word.WdWindowState.wdWindowStateMaximize
                                    objWordEndNote.Visible = False
                                    objWordEndNote.Documents(preliminaryXmlRepairedFullPath).Close(Word.WdSaveOptions.wdDoNotSaveChanges)
                                    'objWordEndNote.Activate()
                                    'Me.Cursor = System.Windows.Forms.Cursors.Default
                                    ''ProgressBar1.Hide()

                                    'If File.Exists(preliminaryXmlRepairedFullPath) Then

                                    '    File.Delete(preliminaryXmlRepairedFullPath)

                                    'End If

                                    'MsgBox("Although this program may have been able to recover some of your document " _
                                    '    & "there may be additional text which is salvageable without formatting using " _
                                    '    & "SilverCoder's command line program, DocToText. Please choose ""Method 4 - " _
                                    '    & "Salvage Text or Data If Possible"" choice if desired on the menu")

                                    'Deal_With_Unspecified_Error = "OK"
                                    wordFileOpendOK = True

                                    'Exit Function

                                Catch ex_EndNote As Exception

                                    CloseWordApplicationWithDocument(objWordEndNote, objDocEndNote)

                                    'We need to delete zipExtensionXMLRepairedFullPath and preliminaryXmlRepairedFullPath 
                                    'because we will be getting new versions soon

                                    If ex_EndNote.Message.Contains("Unspecified error") Then

                                        byteMatch = Regex.Match(ex_EndNote.Message, ("Line: [0-9]+"))
                                        byteMatchString = byteMatch.ToString
                                        secondLineColumnErrorLocationString = DelFromLeft("Line: ", byteMatchString)

                                        Integer.TryParse(secondLineColumnErrorLocationString, secondLineColumnErrorLocationInteger)

                                        Using sr As StreamReader = New StreamReader(modDocumentXmlFullPath)

                                            Do While sr.Peek() >= 0
                                                lineListEndNote.Add(sr.ReadLine())
                                            Loop

                                            sr.Close()

                                        End Using

                                        'So this is the 2nd round of excising EndNote tags

                                        indexBadEndNoteTagOpenLine = lineListEndNote.LastIndexOf("<w:instrText xml:space=""preserve""> ADDIN EN.CITE <EndNote>", secondLineColumnErrorLocationInteger) + 0
                                        indexBadEndNoteTagCloseLine = lineListEndNote.IndexOf("</EndNote>", secondLineColumnErrorLocationInteger) + 2

                                        numberOfCharactersInBadTags = indexBadEndNoteTagCloseLine - indexBadEndNoteTagOpenLine

                                        'Here we write the soon to be excised lines to our text box.

                                        For Each excisedString In lineListEndNote.GetRange(indexBadEndNoteTagOpenLine, numberOfCharactersInBadTags)

                                            RichTextBox1.AppendText(excisedString)

                                        Next

                                        lineListEndNote.RemoveRange(indexBadEndNoteTagOpenLine, numberOfCharactersInBadTags)
                                        File.WriteAllLines(modDocumentXmlFullPath, lineListEndNote.ToArray)
                                        lineListEndNote.Clear()

                                    ElseIf ex_EndNote.Message.Contains("The name in the end tag of the element must match the element type in the start tag.") Then

                                        'Subroutine, see below.

                                        'MsgBox("We seem to have corrected the ""Unspecified"" error(s) but there are " _
                                        '& "additional kinds we now need to address for the file to open.")

                                        Deal_With_Name_Tag_Error_Detail(preliminaryXmlRepairedFullPath, _
                                                    extractedRepairedZipOutputDirectory, modifiedExtractedRepairedZipOutputDirectory)

                                        Re_Zip(modifiedExtractedRepairedZipOutputDirectory, zipExtensionXMLRepairedFullPath)

                                    Else

                                        'Subroutine, see below.

                                        'MsgBox("We seem to have corrected the ""Unspecified"" error(s) but there are " _
                                        '& "additional kinds we now need to address for the file to open.")

                                        If Deal_With_Everything_Else_After_Unspecified_Error_Fix(preliminaryXmlRepairedFullPath, _
                                                    extractedRepairedZipOutputDirectory, modifiedExtractedRepairedZipOutputDirectory) = "OK" Then

                                        End If

                                        Re_Zip(modifiedExtractedRepairedZipOutputDirectory, zipExtensionXMLRepairedFullPath)

                                    End If

                                Finally

                                    CloseWordApplicationWithDocument(objWordEndNote, objDocEndNote)

                                End Try

                            Loop

                            'MsgBox("Check the document.xml file in the EndNoteCorrupt extracted folder")

                        Else

                            MsgBox("Your unspecified error appears to be of an unusual type. " _
                                    & "This program is reverting to another form of recovery that " _
                                    & "uses the command line program xmllint.", MsgBoxStyle.Exclamation)

                            'So here we just treat with xmllint and hope that fixes the issue.
                            'If not we can try XML truncation and DocToText salvaging.

                            Process_XML_Files_With_Xmllint_First(corruptFileXMLSubFileFullName, modDocumentXmlFullPath)
                            normalUnspecifiedError = False

                        End If

                        Continue For

                    Else

                        'Here we process all the other XML files in case any are corrupt

                        Continue For

                    End If

                Else

                    Continue For

                End If

            Next

            'Now we will rezip our directory.

            If File.Exists(zipExtensionXMLRepairedFullPath) Then

                File.Delete(zipExtensionXMLRepairedFullPath)

            End If

            If Not Re_Zip(modifiedExtractedRepairedZipOutputDirectory, zipExtensionXMLRepairedFullPath) = "OK" Then

                MsgBox("You docx file's zip structure was presumably repaired and the file's " _
                       & "word/document.xml subfile was prepared for providing the location " _
                       & "of the error in case it was not specified in ther opening error " _
                       & "message. The reason for the file's unspecified error was also " _
                       & "apparently fixed. However the file could not be rezipped from its " _
                       & "unzipped state that was necessary for fixing. This is unusual. " _
                       & "Try again or try individual recovery methods from the Recover menu.")

                Deal_With_Unspecified_Error = "Failed"

                Exit Function

            End If

            If File.Exists(zipExtensionXMLRepairedFullPath) Then

                File.Copy(zipExtensionXMLRepairedFullPath, xmlRepairedFullPath, True)

            End If

            'ProgressBar1.Value= 95

            'We copy the results to out final name, which is "unspecified_error_repaired_"
            'plus tha original name of the file in the original directory. In order to be
            'able to write to any directory where are our target file is loaded from, I 
            'abrogated full admin rights to the program whenver it starts.

            objWordXMLProcessingCompleted = New Word.Application
            objDocXMLProcessingCompleted = New Word.Document

            Try

                'The next few lines change the repaired file name from the random number
                'to a more conventionally named file based on the original file name.

                If File.Exists(originalFileFullPathRepairedName) Then

                    File.Delete(originalFileFullPathRepairedName)

                End If

                File.Copy(xmlRepairedFullPath, originalFileFullPathRepairedName, True)

                objDocXMLProcessingCompleted = objWordXMLProcessingCompleted.Documents.OpenNoRepairDialog(originalFileFullPathRepairedName, OpenAndRepair:=True)
                objWordXMLProcessingCompleted.WindowState = Word.WdWindowState.wdWindowStateMaximize
                objWordXMLProcessingCompleted.Visible = True
                objWordXMLProcessingCompleted.Activate()
                Me.Cursor = System.Windows.Forms.Cursors.Default
                'ProgressBar1.Hide()

                If File.Exists(xmlRepairedFullPath) Then

                    File.Delete(xmlRepairedFullPath)

                End If

                MsgBox("Although this program may have been able to recover some of your document " _
                    & "there may be additional text which is salvageable without formatting using " _
                    & "SilverCoder's command line program, DocToText. Please choose the """ _
                    & "Try Salvaging Just Text"" choice on the Recover menu if desired.")

                Deal_With_Unspecified_Error = "OK"

                Exit Function

            Catch ex_TryingTruncatingFirst As Exception

                CloseWordApplicationWithDocument(objWordXMLProcessingCompleted, objDocXMLProcessingCompleted)

                If File.Exists(xmlRepairedFullPath) Then

                    File.Delete(xmlRepairedFullPath)

                End If

                If File.Exists(originalFileFullPathRepairedName) Then

                    File.Delete(originalFileFullPathRepairedName)

                End If

                modExtractedCorruptFileDirPath = modifiedExtractedRepairedZipOutputDirectory & "\"

                'Here we process each .xml or .rels file by validating it first. If it is invalid,
                'we truncate it at the error and repair with xmllint. We try first adding no bits
                'extra to truncate, if the the file after treating with xmllint does not validate,
                'we remove 50 bits, then try 100 and running through the xmlrepai again.

                modExtractedCorruptFileDirInfo = New DirectoryInfo(modExtractedCorruptFileDirPath)
                modCorruptFileDirRetrievedFilesInfoArray = _
                    modExtractedCorruptFileDirInfo.GetFiles("*.*", SearchOption.AllDirectories)

                For Each modCorruptFileDirRetrievedFileInfo In modCorruptFileDirRetrievedFilesInfoArray

                    If InStr(modCorruptFileDirRetrievedFileInfo.Extension, ".xml") Or _
                        InStr(modCorruptFileDirRetrievedFileInfo.Extension, ".rels") Then

                        If ProgressBar1.Value + progressBarIncrement > 100 Then

                            progressBarIncrement = 0

                        End If

                        'ProgressBar1.Value=ProgressBar1.Value+ progressBarIncrement

                        Dim modCorruptFileDirRetrievedXMLFileFullPath As String = modCorruptFileDirRetrievedFileInfo.FullName
                        Dim modExtractedCorruptFileDirPathCharacterCount As Integer = _
                            modifiedExtractedRepairedZipOutputDirectory.Length
                        Dim modCorruptFileDirRetrievedXMLFileFullPathCharacterCount As Integer = _
                            modCorruptFileDirRetrievedXMLFileFullPath.Length
                        Dim modCorruptFileDirRetrievedXMLFileWithinArchivePathCharacterCount As Integer = _
                            modCorruptFileDirRetrievedXMLFileFullPathCharacterCount - _
                            modExtractedCorruptFileDirPathCharacterCount

                        Process_XML_Files_With_Truncator_First(modCorruptFileDirRetrievedXMLFileFullPath)

                    End If

                Next

                'Now we will rezip our directory and open as a docx.

                If File.Exists(zipExtensionXMLRepairedTruncTreatedFullPath) Then

                    File.Delete(zipExtensionXMLRepairedTruncTreatedFullPath)

                End If

                If Not Re_Zip(modifiedExtractedRepairedZipOutputDirectory, zipExtensionXMLRepairedTruncTreatedFullPath) = "OK" Then

                    MsgBox("You docx file's zip structure was presumably repaired and the file's word/" _
                           & "document.xml subfile was prepared for providing the location of the " _
                           & "error in case it was not specified in ther opening error message. The " _
                           & "reason for the file's unspecified error was also apparently fixed and " _
                           & "further issues in the word/document.xml subfile and other XML subfiles " _
                           & "were further apparently repaired. However the file could not be rezipped from " _
                           & "its unzipped state that was necessary for fixing. This is unusual. " _
                           & "Try again or try individual recovery methods from the Recover menu.")

                    Deal_With_Unspecified_Error = "Failed"

                    Exit Function

                End If

                File.Copy(zipExtensionXMLRepairedTruncTreatedFullPath, xmlRepairedTruncTreatedFullPath, True)

                If File.Exists(zipExtensionXMLRepairedTruncTreatedFullPath) Then

                    File.Delete(zipExtensionXMLRepairedTruncTreatedFullPath)

                End If

                objWordTruncationSalvagingCompleted = New Word.Application
                objDocTruncationSalvagingCompleted = New Word.Document

                Try

                    'The next few lines bypass the section where the corrupt 
                    'file is copied to one whose name is a random number.

                    File.Copy(xmlRepairedTruncTreatedFullPath, originalFileFullPathRepairedName, True)

                    objDocTruncationSalvagingCompleted = objWordTruncationSalvagingCompleted.Documents.OpenNoRepairDialog(originalFileFullPathRepairedName, _
                                                        OpenAndRepair:=True)
                    objWordTruncationSalvagingCompleted.Visible = True
                    objWordTruncationSalvagingCompleted.WindowState = Word.WdWindowState.wdWindowStateMaximize
                    objWordTruncationSalvagingCompleted.Activate()
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    'ProgressBar1.Hide()

                    If File.Exists(xmlRepairedTruncTreatedFullPath) Then

                        File.Delete(xmlRepairedTruncTreatedFullPath)

                    End If

                    MsgBox("Although this program may have been able to recover some of your document " _
                        & "there may be additional text which is salvageable without formatting using " _
                        & "SilverCoder's command line program, DocToText. Please choose the """ _
                        & "Try Salvaging Just Text"" choice on the Recover menu if desired.")

                    Deal_With_Unspecified_Error = "OK"

                    Exit Function

                Catch ex_TruncationSalvagingCompleted As Exception

                    CloseWordApplicationWithDocument(objWordTruncationSalvagingCompleted, objDocTruncationSalvagingCompleted)

                    If File.Exists(xmlRepairedTruncTreatedFullPath) Then

                        File.Delete(xmlRepairedTruncTreatedFullPath)

                    End If

                    MsgBox("The program was unable to fix your unspecified or matching " _
                       & "tag error through normal means. Please try another " _
                       & "method on the menu such as XMLLint Treatment, XML " _
                       & "Truncation or Text Salvaging.", MsgBoxStyle.Exclamation)

                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    'ProgressBar1.Hide()

                    Deal_With_Unspecified_Error = "Failed"

                    Exit Function

                Finally

                    GC.Collect()
                    GC.WaitForPendingFinalizers()

                    ' GC needs to be called twice in order to get the Finalizers called  
                    ' - the first time in, it simply makes a list of what is to be  
                    ' finalized, the second time in, it actually is finalizing. Only  
                    ' then will the object do its automatic ReleaseComObject. 

                    GC.Collect()
                    GC.WaitForPendingFinalizers()

                End Try

            Finally

                GC.Collect()
                GC.WaitForPendingFinalizers()

                ' GC needs to be called twice in order to get the Finalizers called  
                ' - the first time in, it simply makes a list of what is to be  
                ' finalized, the second time in, it actually is finalizing. Only  
                ' then will the object do its automatic ReleaseComObject. 

                GC.Collect()
                GC.WaitForPendingFinalizers()

            End Try

        Finally

            'MsgBox("We are going through the PrettyPrinted Final section.")
            Me.Cursor = System.Windows.Forms.Cursors.Default

            CloseWordApplicationWithDocument(objWordPrettyPrintedOpenAttempt, objDocPrettyPrintedOpenAttempt)

            'Now we need to get rid of our preliminary XmlRepairedFullPath if it still exists.

            If Directory.Exists(extractedRepairedZipOutputDirectory) Then

                Directory.Delete(extractedRepairedZipOutputDirectory, True)

            End If

            If Directory.Exists(modifiedExtractedRepairedZipOutputDirectory) Then

                Directory.Delete(modifiedExtractedRepairedZipOutputDirectory, True)

            End If

            If File.Exists(sFile) Then

                File.Delete(sFile)

            End If

            If File.Exists(zipExtensionXMLRepairedFullPath) Then

                File.Delete(zipExtensionXMLRepairedFullPath)

            End If

            If File.Exists(zipRepairedBasePathAndFileName) Then

                File.Delete(zipRepairedBasePathAndFileName)

            End If

            If File.Exists(prettyPrintedDocumentXmlFullPath) Then

                File.Delete(prettyPrintedDocumentXmlFullPath)

            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' GC needs to be called twice in order to get the Finalizers called  
            ' - the first time in, it simply makes a list of what is to be  
            ' finalized, the second time in, it actually is finalizing. Only  
            ' then will the object do its automatic ReleaseComObject. 

            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try

    End Function

    Private Function Deal_With_Name_Tag_Error(sFileFullName As String)

        Dim sFile As String = sFileFullName
        Dim sFileInfo As FileInfo = New FileInfo(sFile)
        Dim corruptFileXMLSubFileFullName As String = Nothing
        Dim extractedCorruptFileDirInfo As DirectoryInfo
        Dim corruptFileDirRetrievedFilesInfoArray As FileInfo()
        Dim corruptFileXMLSubFileName As String = Nothing
        Dim sFileExtension As String = Nothing
        Dim extractedCorruptFileDirPath As String = Nothing
        Dim sFileName As String = LCase(sFileInfo.Name)
        'Dim sFileBasePath As String = LCase(sFileInfo.Directory.ToString)
        Dim sFileBasePath As String = tempPath
        Dim zipRepairedsFileName As String = "zipRepaired_" & sFileName & ".zip"
        Dim xmlRepairedFileName As String = "end_tag_error_repaired_" & sFileName
        Dim originalFileFullPath As String = TextBox1.Text
        Dim originalFileFullPathRepairedName As String = originalFileFullPath & ".savvy_fix.docx"
        Dim zipExtensionxmlRepairedFileName As String = xmlRepairedFileName & ".zip"
        Dim zipExtensionXMLRepairedFullPath As String = sFileBasePath & "\" & zipExtensionxmlRepairedFileName
        Dim xmlRepairedFullPath As String = sFileBasePath & "\" & xmlRepairedFileName
        Dim zipRepairedBasePathAndFileName As String = sFileBasePath & "\" & zipRepairedsFileName
        Dim zipRepairedBasePathAndFileNameIndexLastPeriod As Integer = zipRepairedBasePathAndFileName.LastIndexOf(".")
        Dim extractedRepairedZipOutputDirectory As String = _
            zipRepairedBasePathAndFileName.Remove(zipRepairedBasePathAndFileNameIndexLastPeriod - 5)
        Dim modifiedextractedRepairedZipOutputDirectory As String = extractedRepairedZipOutputDirectory & "_mod"
        Dim modDocumentXmlFullPath As String = modifiedextractedRepairedZipOutputDirectory & "\word\document.xml"
        Dim corruptFileDirRetrievedFileInfo As FileInfo = New FileInfo(modDocumentXmlFullPath)
        Dim prettyPrintedDocumentXmlsFile As String = "pretty_printed_" & sFileName
        Dim prettyPrintedDocumentXmlFullPath As String = sFileBasePath & "\" & prettyPrintedDocumentXmlsFile

        Deal_With_Name_Tag_Error = "Failed"

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        If Not Deal_With_Zip_Repair_AndPretty_Print_Routine(sFile) = "OK" Then

            Deal_With_Name_Tag_Error = "Failed"
            Exit Function

        End If

        Dim objWordPrettyPrintedOpenTry As New Word.Application
        Dim objDocPrettyPrintedOpenTry As New Word.Document

        Try

            objDocPrettyPrintedOpenTry = objWordPrettyPrintedOpenTry.Documents.OpenNoRepairDialog(prettyPrintedDocumentXmlFullPath, OpenAndRepair:=True)
            objWordPrettyPrintedOpenTry.WindowState = Word.WdWindowState.wdWindowStateMaximize
            objWordPrettyPrintedOpenTry.Activate()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            'ProgressBar1.Hide()

            'MsgBox("That's weird, the file shouldn't open at this point. We have only altered it to show what line an error is at.")

            If File.Exists(prettyPrintedDocumentXmlFullPath) Then

                File.Delete(prettyPrintedDocumentXmlFullPath)

            End If

            Deal_With_Name_Tag_Error = "OK"
            Exit Function

        Catch ex2 As Exception

            CloseWordApplicationWithDocument(objWordPrettyPrintedOpenTry, objDocPrettyPrintedOpenTry)

            sFileExtension = sFileInfo.Extension

            If File.Exists(zipRepairedBasePathAndFileName) Then

                File.Delete(zipRepairedBasePathAndFileName)

            End If

            File.Copy(prettyPrintedDocumentXmlFullPath, zipRepairedBasePathAndFileName, True)

            If File.Exists(prettyPrintedDocumentXmlFullPath) Then

                Try
                    File.Delete(prettyPrintedDocumentXmlFullPath)
                Catch ex As Exception

                End Try


            End If

            'ProgressBar1.Value

            'Now we extract the repaired file.

            If Directory.Exists(extractedRepairedZipOutputDirectory) Then

                Directory.Delete(extractedRepairedZipOutputDirectory, True)

            End If

            If Not Extract_Zip(zipRepairedBasePathAndFileName, extractedRepairedZipOutputDirectory) = "OK" Then

                MsgBox("Savvy DOCX Recovery could not extract the zip structure " _
                       & "of your target file. There is apparently nothing " _
                       & "recoverable but you can try Method 4, ""Salvage Text " _
                       & "or Data If Possible"", on the Recover Menu to be sure.")

                Deal_With_Name_Tag_Error = "Failed"

                Exit Function

            End If

            If Directory.Exists(modifiedextractedRepairedZipOutputDirectory) Then

                Directory.Delete(modifiedextractedRepairedZipOutputDirectory, True)

            End If

            If Not Extract_Zip(zipRepairedBasePathAndFileName, modifiedextractedRepairedZipOutputDirectory) = "OK" Then

                MsgBox("Savvy DOCX Recovery could not extract the zip structure " _
                       & "of your target file. There is apparently nothing " _
                       & "recoverable but you can try Method 4, ""Salvage Text " _
                       & "or Data If Possible"", on the Recover Menu to be sure.")

                Deal_With_Name_Tag_Error = "Failed"

                Exit Function

            End If

            'ProgressBar1.Value= 30

            If File.Exists(zipRepairedBasePathAndFileName) Then

                File.Delete(zipRepairedBasePathAndFileName)

            End If

            extractedCorruptFileDirPath = extractedRepairedZipOutputDirectory & "\"
            extractedCorruptFileDirInfo = New DirectoryInfo(extractedCorruptFileDirPath)
            corruptFileDirRetrievedFilesInfoArray = _
                extractedCorruptFileDirInfo.GetFiles("*.*", SearchOption.AllDirectories)

            'ProgressBar1.Value= 40

            Dim extractedCorruptFileDirInfoArrayCount As Integer = _
                corruptFileDirRetrievedFilesInfoArray.GetLength(0)
            Dim progressBarIncrement As Integer = 50 \ (extractedCorruptFileDirInfoArrayCount + 1)

            If progressBarIncrement = 0 Then

                progressBarIncrement = 1

            End If

            For Each corruptFileDirRetrievedFileInfo In corruptFileDirRetrievedFilesInfoArray

                corruptFileXMLSubFileName = corruptFileDirRetrievedFileInfo.Name
                corruptFileXMLSubFileFullName = corruptFileDirRetrievedFileInfo.FullName

                If InStr(corruptFileDirRetrievedFileInfo.Extension, ".xml") Or _
                    InStr(corruptFileDirRetrievedFileInfo.Extension, ".rels") Then

                    Dim wordDocumentXmlFullExtractedPath = extractedRepairedZipOutputDirectory & "\word\document.xml"

                    modDocumentXmlFullPath = modifiedextractedRepairedZipOutputDirectory & _
                        corruptFileXMLSubFileFullName.Substring(extractedRepairedZipOutputDirectory.Length)

                    If corruptFileXMLSubFileFullName = wordDocumentXmlFullExtractedPath Then

                        Deal_With_Name_Tag_Error_Detail(sFileFullName, extractedRepairedZipOutputDirectory, modifiedextractedRepairedZipOutputDirectory)

                    Else

                        'This subroutine processes all other XML based subfiles (other than doucment.xml) 
                        'with xmllint if they show any corruption.

                        Process_XML_Files_With_Xmllint_First(corruptFileXMLSubFileFullName, modDocumentXmlFullPath)

                        Continue For

                    End If

                End If

            Next

            'Now we will rezip our directory and open as a docx file.

            If File.Exists(zipExtensionXMLRepairedFullPath) Then

                File.Delete(zipExtensionXMLRepairedFullPath)

            End If

            Re_Zip(modifiedextractedRepairedZipOutputDirectory, zipExtensionXMLRepairedFullPath)

            'ProgressBar1.Value= 95

            If File.Exists(zipExtensionXMLRepairedFullPath) Then

                File.Copy(zipExtensionXMLRepairedFullPath, xmlRepairedFullPath, True)

            End If

            'Now we try to open the hopefully fixed file.

            Dim objWordAfterNameTagRepair As New Word.Application
            Dim objDocAfterNameTagRepair As New Word.Document

            Try

                If File.Exists(originalFileFullPathRepairedName) Then

                    File.Delete(originalFileFullPathRepairedName)

                End If

                File.Copy(xmlRepairedFullPath, originalFileFullPathRepairedName, True)

                objWordAfterNameTagRepair.Visible = True
                objDocAfterNameTagRepair = objWordAfterNameTagRepair.Documents.OpenNoRepairDialog(originalFileFullPathRepairedName, OpenAndRepair:=True)
                objWordAfterNameTagRepair.WindowState = Word.WdWindowState.wdWindowStateMaximize
                objWordAfterNameTagRepair.Visible = True
                objWordAfterNameTagRepair.Activate()
                Me.Cursor = System.Windows.Forms.Cursors.Default
                'ProgressBar1.Hide()

                MsgBox("Although this program may have been able to recover some of your " _
                    & "document there may be additional text which is salvageable without " _
                    & "formatting using SilverCoder's command line program, DocToText. Please " _
                    & "choose ""Try Salvaging Just Text"" choice on the Recover menu, if desired.")

                Deal_With_Name_Tag_Error = "OK"

                Exit Function

            Catch ex_AfterNameTagRepairException As Exception

                CloseWordApplicationWithDocument(objWordAfterNameTagRepair, objDocAfterNameTagRepair)

                extractedCorruptFileDirPath = extractedRepairedZipOutputDirectory & "\"

                'Here we process each .xml or .rels file by validating it first. If it is invalid,
                'we truncate it at the error and repair with xmllint. We try first adding no bits
                'extra to truncate, if the the file after treating with xmllint does not validate,
                'we remove 50 bits, then try 100 and running through the xmlrepai again.

                extractedCorruptFileDirInfo = New DirectoryInfo(extractedCorruptFileDirPath)
                corruptFileDirRetrievedFilesInfoArray = _
                    extractedCorruptFileDirInfo.GetFiles("*.*", SearchOption.AllDirectories)

                For Each corruptFileDirRetrievedFileInfo In corruptFileDirRetrievedFilesInfoArray

                    If InStr(corruptFileDirRetrievedFileInfo.Extension, ".xml") Or _
                        InStr(corruptFileDirRetrievedFileInfo.Extension, ".rels") Then

                        If ProgressBar1.Value + progressBarIncrement > 100 Then

                            progressBarIncrement = 0

                        End If

                        'ProgressBar1.Value=ProgressBar1.Value+ progressBarIncrement

                        Dim corruptFileDirRetrievedXMLFileFullPath As String = corruptFileDirRetrievedFileInfo.FullName
                        Dim extractedCorruptFileDirPathCharacterCount As Integer = _
                            extractedRepairedZipOutputDirectory.Length
                        Dim corruptFileDirRetrievedXMLFileFullPathCharacterCount As Integer = _
                            corruptFileDirRetrievedXMLFileFullPath.Length
                        Dim corruptFileDirRetrievedXMLFileWithinArchivePathCharacterCount As Integer = _
                            corruptFileDirRetrievedXMLFileFullPathCharacterCount - _
                            extractedCorruptFileDirPathCharacterCount

                        Process_XML_Files_With_Truncator_First(corruptFileDirRetrievedXMLFileFullPath)

                    End If

                Next

                'Now we will rezip our directory and open as a docx.

                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                    File.Delete(zipExtensionXMLRepairedFullPath)

                End If

                Re_Zip(extractedRepairedZipOutputDirectory, zipExtensionXMLRepairedFullPath)

                If File.Exists(xmlRepairedFullPath) Then

                    File.Delete(xmlRepairedFullPath)

                End If

                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                    File.Copy(zipExtensionXMLRepairedFullPath, xmlRepairedFullPath, True)

                End If

                Dim objWordAfterXMLTruncatorFirst As New Word.Application
                Dim objDocAfterXMLTruncatorFirst As New Word.Document

                Try

                    'The next few lines bypass the section where the corrupt 
                    'file is copied to one whose name is a random number.

                    If File.Exists(originalFileFullPathRepairedName) Then

                        File.Delete(originalFileFullPathRepairedName)

                    End If

                    File.Copy(xmlRepairedFullPath, originalFileFullPathRepairedName, True)

                    objWordAfterXMLTruncatorFirst.Visible = True
                    objDocAfterXMLTruncatorFirst = objWordAfterXMLTruncatorFirst.Documents.OpenNoRepairDialog(originalFileFullPathRepairedName, OpenAndRepair:=True)
                    objWordAfterXMLTruncatorFirst.WindowState = Word.WdWindowState.wdWindowStateMaximize
                    objWordAfterXMLTruncatorFirst.Visible = True
                    objWordAfterXMLTruncatorFirst.Activate()
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    'ProgressBar1.Hide()

                    MsgBox("Although this program may have been able to recover some of your document " _
                        & "there may be additional text which is salvageable without formatting using " _
                        & "SilverCoder's command line program, DocToText. Please choose ""Method 4 - " _
                        & "Salvage Text or Data If Possible"" choice if desired on the menu")

                    Deal_With_Name_Tag_Error = "OK"

                    Exit Function

                Catch ex_After_XML_Truncator_First_Treatment As Exception

                    CloseWordApplicationWithDocument(objWordAfterXMLTruncatorFirst, objDocAfterXMLTruncatorFirst)

                    MsgBox("The program was unable to fix your unspecified " _
                       & "or matching tag error through normal means. " _
                       & "Please try another method on the menu such as " _
                       & "XmlLint, XML truncation or text salvaging.", MsgBoxStyle.Exclamation)

                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    'ProgressBar1.Hide()

                    Deal_With_Name_Tag_Error = "Failed"

                    Exit Function

                Finally
                    If Not (Deal_With_Name_Tag_Error = "OK") Then
                        CloseWordApplicationWithDocument(objWordAfterXMLTruncatorFirst, objDocAfterXMLTruncatorFirst)
                    End If
                End Try

            Finally

                CloseWordApplicationWithDocument(objWordAfterNameTagRepair, objDocAfterNameTagRepair)

                GC.Collect()
                GC.WaitForPendingFinalizers()

                ' GC needs to be called twice in order to get the Finalizers called  
                ' - the first time in, it simply makes a list of what is to be  
                ' finalized, the second time in, it actually is finalizing. Only  
                ' then will the object do its automatic ReleaseComObject. 

                GC.Collect()
                GC.WaitForPendingFinalizers()

            End Try

        Finally

            Me.Cursor = System.Windows.Forms.Cursors.Default

            If Directory.Exists(extractedRepairedZipOutputDirectory) Then

                Directory.Delete(extractedRepairedZipOutputDirectory, True)

            End If

            If Directory.Exists(modifiedextractedRepairedZipOutputDirectory) Then

                Directory.Delete(modifiedextractedRepairedZipOutputDirectory, True)

            End If

            If File.Exists(sFile) Then

                File.Delete(sFile)

            End If

            If File.Exists(zipRepairedsFileName) Then

                File.Delete(zipRepairedsFileName)

            End If

            If File.Exists(zipRepairedBasePathAndFileName) Then

                File.Delete(zipRepairedBasePathAndFileName)

            End If

            If File.Exists(zipExtensionXMLRepairedFullPath) Then

                File.Delete(zipExtensionXMLRepairedFullPath)

            End If

            CloseWordApplicationWithDocument(objWordPrettyPrintedOpenTry, objDocPrettyPrintedOpenTry)

        End Try

    End Function

    Private Sub Deal_With_Name_Tag_Error_Detail(sFileFullName As String, _
            extractedRepairedZipOutputDirectory As String, modifiedextractedRepairedZipOutputDirectory As String)

        Dim sFile As String = sFileFullName
        Dim byteMatchStartTag As Match = Nothing
        Dim byteMatchEndTag As Match = Nothing
        Dim sFileInfo As FileInfo = New FileInfo(sFile)
        Dim byteMatchStringStartTag As String = Nothing
        Dim byteMatchStringEndTag As String = Nothing
        Dim corruptFileXMLSubFileFullName As String = extractedRepairedZipOutputDirectory & "\word\document.xml"
        Dim modDocumentXmlFullPath As String = modifiedextractedRepairedZipOutputDirectory & "\word\document.xml"
        Dim lineListOfEntireUnspecifiedErrorFixedXMLDocument As New List(Of String)
        Dim startTagTrimmedString As String
        Dim endTagTrimmedString As String
        Dim endTagTrimmedStringLocationInteger As Integer
        Dim corruptFileDirRetrievedFileInfo As FileInfo = New FileInfo(modDocumentXmlFullPath)
        Dim modDocumentXmlFullPath2 As String = modDocumentXmlFullPath
        Dim reader As XmlReader = XmlReader.Create(modDocumentXmlFullPath)
        Dim wordFileShouldOpendOK As Boolean = False
        Dim ex_NameTagError As Exception

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        corruptFileXMLSubFileFullName = corruptFileDirRetrievedFileInfo.FullName

        Using sr As StreamReader = New StreamReader(modDocumentXmlFullPath2)

            Do While sr.Peek() >= 0

                lineListOfEntireUnspecifiedErrorFixedXMLDocument.Add(sr.ReadLine())

            Loop

            'MsgBox("Finished sr stream reading")

            sr.Close()

        End Using

        Do Until wordFileShouldOpendOK = True

            Try

                While (reader.Read())
                End While
                reader.Close()
                'MsgBox("The word file should Open OK")

                wordFileShouldOpendOK = True

            Catch ex_NameTagError

                reader.Close()

                byteMatchStartTag = Regex.Match(ex_NameTagError.Message, ("The '[a-zA-Z0-9]+:[a-zA-Z0-9]+"))
                byteMatchStringStartTag = byteMatchStartTag.ToString
                startTagTrimmedString = DelFromLeft("The '", byteMatchStringStartTag)

                byteMatchEndTag = Regex.Match(ex_NameTagError.Message, ("\. Line [0-9]+"))
                byteMatchStringEndTag = byteMatchEndTag.ToString
                endTagTrimmedString = DelFromLeft(". Line ", byteMatchStringEndTag)

                Integer.TryParse(endTagTrimmedString, endTagTrimmedStringLocationInteger)

                If startTagTrimmedString = "mc:Fallback" Then

                    Dim mcFallbackError As Exception = Nothing
                    Dim mcFallbackTagsShouldBeFixed As Boolean = False

                    Do Until mcFallbackTagsShouldBeFixed = True

                        Dim modDocumentXmlFullPath3 As String = modDocumentXmlFullPath
                        Dim mcFallbackSearchReader As XmlReader = XmlReader.Create(modDocumentXmlFullPath3)

                        Try

                            While (mcFallbackSearchReader.Read())
                            End While
                            mcFallbackSearchReader.Close()

                            'MsgBox("The mcFallbackSearchReader has completed without " _
                            '& "error, the Word file should now open without issue")
                            mcFallbackTagsShouldBeFixed = True

                        Catch mcFallbackError

                            mcFallbackSearchReader.Close()

                            'MsgBox("FallbackError read through. The mcFallbackError Loop error is: " & _
                            'mcFallbackError.ToString)

                            Dim endTagTrimmedmcFallbackStringLocationInteger As Integer = 0
                            Dim byteMatchStartTagmcFallbackDoWhileLoop As Match = _
                                Regex.Match(mcFallbackError.Message, ("The '[a-zA-Z0-9]+:[a-zA-Z0-9]+"))
                            Dim byteMatchStringStartTagmcFallbackDoWhileLoop As String = _
                                byteMatchStartTagmcFallbackDoWhileLoop.ToString
                            Dim startTagTrimmedStringmcFallbackDoWhileLoop As String = _
                                DelFromLeft("The '", byteMatchStringStartTagmcFallbackDoWhileLoop)
                            Dim byteMatchEndTagmcFallbackDoWhileLoop As Match = _
                                Regex.Match(mcFallbackError.Message, ("\. Line [0-9]+"))
                            Dim byteMatchStringEndTagmcFallbackDoWhileLoop As String = _
                                byteMatchEndTagmcFallbackDoWhileLoop.ToString
                            Dim endTagTrimmedStringmcFallbackDoWhileLoop = _
                                DelFromLeft(". Line ", byteMatchStringEndTagmcFallbackDoWhileLoop)

                            Integer.TryParse(endTagTrimmedStringmcFallbackDoWhileLoop, _
                                             endTagTrimmedmcFallbackStringLocationInteger)


                            '& " and the startTagTrimmedString2 is: " & _
                            'startTagTrimmedStringmcFallbackDoWhileLoop _
                            '& ". The missing closed tag is: " & _
                            'endTagTrimmedStringmcFallbackDoWhileLoop)

                            If startTagTrimmedStringmcFallbackDoWhileLoop = "mc:Fallback" Then

                                lineListOfEntireUnspecifiedErrorFixedXMLDocument.Insert(endTagTrimmedmcFallbackStringLocationInteger - 1, "</mc:Fallback>")
                                lineListOfEntireUnspecifiedErrorFixedXMLDocument.Insert(endTagTrimmedmcFallbackStringLocationInteger, "</mc:AlternateContent>")
                                lineListOfEntireUnspecifiedErrorFixedXMLDocument.Insert(endTagTrimmedmcFallbackStringLocationInteger + 1, "</w:r>")

                                'RichTextBox1.Clear()
                                'MsgBox("endTagTrimmedmcFallbackStringLocationInteger.ToString is: " _
                                '& endTagTrimmedmcFallbackStringLocationInteger.ToString)

                                File.WriteAllLines(modDocumentXmlFullPath, lineListOfEntireUnspecifiedErrorFixedXMLDocument.ToArray)
                                'normalEndTagNotMatchingError = True
                                'MsgBox("The word file should Open OK")
                                'wordFileShouldOpendOK = True

                            ElseIf startTagTrimmedStringmcFallbackDoWhileLoop = "mc:AlternateContent" Then

                                lineListOfEntireUnspecifiedErrorFixedXMLDocument.Insert(endTagTrimmedmcFallbackStringLocationInteger - 1, _
                                                        "</mc:AlternateContent>")

                                'MsgBox("endTagTrimmedmcFallbackStringLocationInteger.ToString is: " _
                                '& endTagTrimmedmcFallbackStringLocationInteger.ToString)
                                File.WriteAllLines(modDocumentXmlFullPath, _
                                                   lineListOfEntireUnspecifiedErrorFixedXMLDocument.ToArray)

                            ElseIf startTagTrimmedStringmcFallbackDoWhileLoop = "m:oMath" Then

                                Dim altContPosLineNumberList As New List(Of Integer)
                                Dim altContentPos As Integer = 0

                                Do While altContentPos <> -1

                                    altContentPos = lineListOfEntireUnspecifiedErrorFixedXMLDocument.IndexOf("<mc:AlternateContent>", _
                                                                                                             altContentPos + 1)
                                    'MsgBox(altContentPos.ToString)

                                    If lineListOfEntireUnspecifiedErrorFixedXMLDocument.Item(altContentPos + 2) = "<m:oMath>" Then

                                        altContPosLineNumberList.Add(altContentPos)

                                        lineListOfEntireUnspecifiedErrorFixedXMLDocument.Insert(altContentPos, "<m:oMath>")

                                        Dim mathTagPos As Integer = _
                                            lineListOfEntireUnspecifiedErrorFixedXMLDocument.IndexOf("<m:oMath>", altContentPos + 2)

                                        'MsgBox(mathTagPos.ToString)

                                        lineListOfEntireUnspecifiedErrorFixedXMLDocument.RemoveRange(mathTagPos, 1)
                                        File.WriteAllLines(modDocumentXmlFullPath, lineListOfEntireUnspecifiedErrorFixedXMLDocument.ToArray)

                                    End If

                                Loop

                            End If

                        Finally

                        End Try

                    Loop

                ElseIf startTagTrimmedString = "mc:AlternateContent" Then

                    Dim AlternateContentError As Exception = Nothing
                    Dim alternateContentShouldBeFixed As Boolean = False

                    Do Until alternateContentShouldBeFixed = True

                        Dim modDocumentXmlFullPath4 As String = modDocumentXmlFullPath
                        Dim AlternateContentSearchReader As XmlReader = XmlReader.Create(modDocumentXmlFullPath4)
                        Dim endTagTrimmedAlternateContentStringLocationInteger As Integer = 0
                        Dim byteMatchStartTagAlternateContentDoWhileLoop As Match = _
                            Regex.Match(AlternateContentError.Message, ("The '[a-zA-Z0-9]+:[a-zA-Z0-9]+"))
                        Dim byteMatchStringStartTagAlternateContentDoWhileLoop As String = _
                            byteMatchStartTagAlternateContentDoWhileLoop.ToString
                        Dim startTagTrimmedStringAlternateContentDoWhileLoop As String = _
                            DelFromLeft("The '", byteMatchStringStartTagAlternateContentDoWhileLoop)
                        Dim byteMatchEndTagAlternateContentDoWhileLoop As Match = _
                            Regex.Match(AlternateContentError.Message, ("\. Line [0-9]+"))
                        Dim byteMatchStringEndTagAlternateContentDoWhileLoop As String = _
                            byteMatchEndTagAlternateContentDoWhileLoop.ToString
                        Dim endTagTrimmedStringAlternateContentDoWhileLoop = _
                            DelFromLeft(". Line ", byteMatchStringEndTagAlternateContentDoWhileLoop)

                        Integer.TryParse(endTagTrimmedStringAlternateContentDoWhileLoop, _
                                         endTagTrimmedAlternateContentStringLocationInteger)
                        'MsgBox("endTagTrimmedAlternateContentStringLocationInteger.ToString is: " _
                        '& endTagTrimmedAlternateContentStringLocationInteger.ToString)

                        Try

                            While (AlternateContentSearchReader.Read())
                            End While

                            'MsgBox("The mcFallbackSearchReader has completed without " _
                            '& "error, the Word file should now open without issue")

                            alternateContentShouldBeFixed = True

                        Catch AlternateContentError

                            AlternateContentSearchReader.Close()

                            'MsgBox("The AlternateContentError Loop error is: " & _
                            ' AlternateContentError.ToString _
                            ' & " and the startTagTrimmedString2 is: " & _
                            ' startTagTrimmedStringAlternateContentDoWhileLoop _
                            ' & ". The missing closed tag is: " & _
                            'endTagTrimmedStringAlternateContentDoWhileLoop)

                            lineListOfEntireUnspecifiedErrorFixedXMLDocument.Insert(endTagTrimmedAlternateContentStringLocationInteger - 1, _
                                                                "</mc:AlternateContent>")
                            File.WriteAllLines(modDocumentXmlFullPath, lineListOfEntireUnspecifiedErrorFixedXMLDocument.ToArray)

                        End Try

                    Loop

                    Dim mcFallbackError As Exception = Nothing
                    Dim mcFallbackTagsShouldBeFixed As Boolean = False

                    Do Until mcFallbackTagsShouldBeFixed = True

                        Dim modDocumentXmlFullPath3 As String = modDocumentXmlFullPath
                        Dim mcFallbackSearchReader As XmlReader = XmlReader.Create(modDocumentXmlFullPath3)
                        Dim endTagTrimmedmcFallbackStringLocationInteger As Integer = 0
                        Dim byteMatchStartTagmcFallbackDoWhileLoop As Match = _
                            Regex.Match(mcFallbackError.Message, ("The '[a-zA-Z0-9]+:[a-zA-Z0-9]+"))
                        Dim byteMatchStringStartTagmcFallbackDoWhileLoop As String = _
                            byteMatchStartTagmcFallbackDoWhileLoop.ToString
                        Dim startTagTrimmedStringmcFallbackDoWhileLoop As String = _
                            DelFromLeft("The '", byteMatchStringStartTagmcFallbackDoWhileLoop)
                        Dim byteMatchEndTagmcFallbackDoWhileLoop As Match = _
                            Regex.Match(mcFallbackError.Message, ("\. Line [0-9]+"))
                        Dim byteMatchStringEndTagmcFallbackDoWhileLoop As String = _
                            byteMatchEndTagmcFallbackDoWhileLoop.ToString
                        Dim endTagTrimmedStringmcFallbackDoWhileLoop = _
                            DelFromLeft(". Line ", byteMatchStringEndTagmcFallbackDoWhileLoop)

                        Integer.TryParse(endTagTrimmedStringmcFallbackDoWhileLoop, _
                                         endTagTrimmedmcFallbackStringLocationInteger)

                        'MsgBox("endTagTrimmedStringLocationInteger.ToString is: " _
                        '& endTagTrimmedmcFallbackStringLocationInteger.ToString)

                        Try

                            While (mcFallbackSearchReader.Read())
                            End While

                            'MsgBox("The mcFallbackSearchReader has completed without " _
                            ' & "error, the Word file should now open without issue")
                            'mcFallbackTagsShouldBeFixed = True

                        Catch mcFallbackError

                            mcFallbackSearchReader.Close()

                            'MsgBox("The mcFallbackError Loop error is: " & _
                            ' mcFallbackError.ToString _
                            ' & " and the startTagTrimmedString2 is: " & _
                            ' startTagTrimmedStringmcFallbackDoWhileLoop _
                            ' & ". The missing closed tag is: " & _
                            'endTagTrimmedStringmcFallbackDoWhileLoop)

                            lineListOfEntireUnspecifiedErrorFixedXMLDocument.Insert(endTagTrimmedmcFallbackStringLocationInteger - 1, "</mc:Fallback>")
                            lineListOfEntireUnspecifiedErrorFixedXMLDocument.Insert(endTagTrimmedmcFallbackStringLocationInteger, "</mc:AlternateContent>")
                            lineListOfEntireUnspecifiedErrorFixedXMLDocument.Insert(endTagTrimmedmcFallbackStringLocationInteger + 1, "</w:r>")
                            File.WriteAllLines(modDocumentXmlFullPath, lineListOfEntireUnspecifiedErrorFixedXMLDocument.ToArray)

                        End Try

                    Loop

                    'So here we finally deal with <mc:AlternateContent> or <m:oMath>
                    ' being in the wrong order usually with one other tag in between.

                    Dim altContPosLineNumberList As New List(Of Integer)
                    Dim altContentPos As Integer = 0

                    Do While altContentPos <> -1

                        altContentPos = lineListOfEntireUnspecifiedErrorFixedXMLDocument.IndexOf("<mc:AlternateContent>", _
                                                                                                 altContentPos + 1)
                        'MsgBox(altContentPos.ToString)

                        If lineListOfEntireUnspecifiedErrorFixedXMLDocument.Item(altContentPos + 2) = "<m:oMath>" Then

                            'We make a list of line position where <m:oMath> appears on the 
                            'second line below <mc:AlternateContent>. These are the source
                            'of only a small percentage of "end tag" errors in my experience. 

                            altContPosLineNumberList.Add(altContentPos)

                            lineListOfEntireUnspecifiedErrorFixedXMLDocument.Insert(altContentPos, "<m:oMath>")

                            Dim mathTagPos As Integer = _
                                lineListOfEntireUnspecifiedErrorFixedXMLDocument.IndexOf("<m:oMath>", altContentPos + 2)

                            ' MsgBox(mathTagPos.ToString)

                            lineListOfEntireUnspecifiedErrorFixedXMLDocument.RemoveRange(mathTagPos, 1)
                            File.WriteAllLines(modDocumentXmlFullPath, lineListOfEntireUnspecifiedErrorFixedXMLDocument.ToArray)

                        End If

                    Loop

                ElseIf startTagTrimmedString = "m:oMath" Then

                    'MsgBox("Going through the m:oMath first section.")
                    'So here we finally deal with <mc:AlternateContent> or <m:oMath> being in the wrong order
                    'usually with one other tag in between.

                    Dim altContPosLineNumberList As New List(Of Integer)
                    Dim altContentPos As Integer = 0

                    Do While altContentPos <> -1

                        altContentPos = lineListOfEntireUnspecifiedErrorFixedXMLDocument.IndexOf("<mc:AlternateContent>", _
                                                                                                 altContentPos + 1)
                        'MsgBox("<mc:AlternateContent> location is: " & altContentPos.ToString)

                        If lineListOfEntireUnspecifiedErrorFixedXMLDocument.Item(altContentPos + 2) = "<m:oMath>" Then

                            altContPosLineNumberList.Add(altContentPos)

                            lineListOfEntireUnspecifiedErrorFixedXMLDocument.Insert(altContentPos, "<m:oMath>")

                            Dim mathTagPos As Integer = _
                                lineListOfEntireUnspecifiedErrorFixedXMLDocument.IndexOf("<m:oMath>", altContentPos + 2)

                            'MsgBox("<m:oMath> position is: " & mathTagPos.ToString)

                            lineListOfEntireUnspecifiedErrorFixedXMLDocument.RemoveRange(mathTagPos, 1)
                            File.WriteAllLines(modDocumentXmlFullPath, lineListOfEntireUnspecifiedErrorFixedXMLDocument.ToArray)

                        End If

                    Loop

                    'ElseIf String.IsNullOrEmpty(startTagTrimmedString) = False Then

                    '    Dim regularTagError As Exception = Nothing
                    '    Dim regularTagsShouldBeFixed As Boolean = False

                    '    Do Until regularTagsShouldBeFixed = True

                    '        Dim modDocumentXmlFullPath22 As String = modDocumentXmlFullPath
                    '        Dim regularTagErrorReader As XmlReader = XmlReader.Create(modDocumentXmlFullPath22)
                    '        Dim endTagTrimmedRegularStringLocationInteger As Integer = 0

                    '        Try

                    '            While (regularTagErrorReader.Read())
                    '            End While

                    '            MsgBox("regularTagErrorReader has completed without " _
                    '                   & "error, the Word file should now open without issue")
                    '            regularTagsShouldBeFixed = True

                    '        Catch regularTagError

                    '            regularTagErrorReader.Close()

                    '            Dim byteMatchStartTagRegularDoWhileLoop As Match = _
                    '                Regex.Match(regularTagError.Message, ("The '[a-zA-Z0-9]+:[a-zA-Z0-9]+"))
                    '            Dim byteMatchStringStartTagRegularDoWhileLoop As String = _
                    '                byteMatchStartTagRegularDoWhileLoop.ToString
                    '            Dim startTagTrimmedStringRegularDoWhileLoop As String = _
                    '                DelFromLeft("The '", byteMatchStringStartTagRegularDoWhileLoop)
                    '            Dim byteMatchEndTagRegularDoWhileLoop As Match = _
                    '                Regex.Match(regularTagError.Message, ("\. Line [0-9]+"))
                    '            Dim byteMatchStringEndTagRegularDoWhileLoop As String = _
                    '                byteMatchEndTagRegularDoWhileLoop.ToString
                    '            Dim endTagTrimmedStringRegularDoWhileLoop = _
                    '                DelFromLeft(". Line ", byteMatchStringEndTagRegularDoWhileLoop)

                    '            Integer.TryParse(endTagTrimmedStringRegularDoWhileLoop, _
                    '                             endTagTrimmedRegularStringLocationInteger)

                    '            'MsgBox("endTagTrimmedStringLocationInteger.ToString is: " _
                    '            '       & endTagTrimmedRegularStringLocationInteger.ToString)

                    '            MsgBox("The regularTagError is: " _
                    '                    & regularTagError.ToString _
                    '                    & ". The missing closed tag is: " _
                    '                    & startTagTrimmedStringRegularDoWhileLoop _
                    '                    & " and the location it needs to go is at: " _
                    '                    & byteMatchStringEndTagRegularDoWhileLoop)

                    '            lineListOfEntireUnspecifiedErrorFixedXMLDocument.Insert(endTagTrimmedRegularStringLocationInteger - 1, _
                    '                                        "</" & startTagTrimmedStringRegularDoWhileLoop & ">")
                    '            File.WriteAllLines(modDocumentXmlFullPath, lineListOfEntireUnspecifiedErrorFixedXMLDocument.ToArray)

                    '        Finally

                    '            regularTagErrorReader.Close()

                    '        End Try

                    '    Loop

                    '    MsgBox("Inserted tag is: " & startTagTrimmedString)
                    '    lineListOfEntireUnspecifiedErrorFixedXMLDocument.Insert(endTagTrimmedStringLocationInteger - 1, _
                    '                                                            "</" & startTagTrimmedString & ">")
                    '    File.WriteAllLines(modDocumentXmlFullPath, lineListOfEntireUnspecifiedErrorFixedXMLDocument.ToArray)

                Else

                    'File.Copy(modDocumentXmlFullPath, corruptFileXMLSubFileFullName, True)
                    Process_XML_Files_With_Xmllint_First(corruptFileXMLSubFileFullName, modDocumentXmlFullPath)
                    'MsgBox("Treated " & modDocumentXmlFullPath & " with Process_XML_Files_With_Xmllint_First.")

                End If

            Finally

                reader.Close()

            End Try

        Loop

    End Sub

    Private Function Deal_With_Everything_Else(sFileFullName As String)

        Dim sFile As String = sFileFullName
        Dim sFileZip As String = sFile & ".zip"
        Dim sFileZipPrelimRepair As String = sFileZip & ".prelim_repair.zip"
        Dim xmlValidateReader As StreamReader = Nothing
        Dim xmlValidateErrorReader As StreamReader = Nothing
        Dim xmlValidateReader2 As StreamReader = Nothing
        Dim xmlValidateErrorReader2 As StreamReader = Nothing
        Dim byteMatch As Match = Nothing
        Dim extractedCorruptFileDirInfo As DirectoryInfo
        Dim corruptFileDirRetrievedFilesInfoArray As FileInfo()
        Dim sFileInfo As FileInfo = New FileInfo(sFile)
        Dim byteErrorLocationInteger As Integer = 0
        Dim truncatedLength As Integer = 0
        Dim intTruncationAmount As Integer = 0
        Dim xmlValidateCompOut As String = Nothing
        Dim xmlValidateErrorOut As String = Nothing
        Dim xmlValidateArguments As String = Nothing
        Dim byteMatchString As String = Nothing
        Dim truncatedLengthAsString As String = Nothing
        Dim truncateArguments As String = Nothing
        Dim xmlValidateCompOut2 As String = Nothing
        Dim xmlValidateErrorOut2 As String = Nothing
        Dim xmlRecoverArguments As String = Nothing
        Dim oldXMLRepairedFileName As String = Nothing
        Dim oldXMLRepairedFullPath As String = Nothing
        Dim sevenZipUpArguments As String = Nothing
        Dim sFileExtension As String = Nothing
        Dim sevenZipFullPath As String = Nothing
        Dim sevenZipExtractArguments As String = Nothing
        Dim extractedCorruptFileDirPath As String = Nothing
        Dim modExtractedCorruptFileDirPath As String = Nothing
        Dim repairZipArguments As String = Nothing
        Dim zipFullPath As String = Nothing
        Dim byteErrorLocation As String = Nothing
        Dim sFileName As String = LCase(sFileInfo.Name)
        Dim sFileFullNameLC As String = LCase(sFileInfo.FullName)
        Dim zipRepairedsFileName As String = "zipRepaired_" & sFileName & ".zip"
        'Dim sFileBasePath As String = LCase(sFileInfo.Directory.ToString)
        Dim sFileBasePath As String = tempPath
        Dim xmlRepairedFileName As String = "xml_error_repaired_" & sFileName
        Dim xmlRepairedFullPath As String = sFileBasePath & "\" & xmlRepairedFileName
        Dim originalFileFullPath As String = TextBox1.Text
        Dim originalFileFullPathRepairedName As String = originalFileFullPath & ".savvy_fix.docx"
        Dim zipExtensionXMLRepairedFullPath As String = xmlRepairedFullPath & ".zip"
        Dim zipRepairedBasePathAndFileName As String = sFileBasePath & "\" & zipRepairedsFileName
        Dim zipRepairedBasePathAndFileNameIndexLastPeriod As Integer = zipRepairedBasePathAndFileName.LastIndexOf(".")
        Dim extractedRepairedZipOutputDirectory As String = _
            zipRepairedBasePathAndFileName.Remove(zipRepairedBasePathAndFileNameIndexLastPeriod - 5)
        Dim prettyPrintedDocumentXmlsFile As String = "pretty_printed_" & sFileName
        Dim prettyPrintedDocumentXmlFullPath As String = sFileBasePath & "\" & prettyPrintedDocumentXmlsFile
        Dim corruptDocumentXMLPath As String = extractedRepairedZipOutputDirectory & "\word\document.xml"
        Dim modifiedExtractedRepairedZipOutputDirectory As String = extractedRepairedZipOutputDirectory & "_mod"
        Dim salvagedsFileName As String = "salvaged_" & sFileName
        Dim salvagedsFileNameNoSpace As String = salvagedsFileName.Replace(" ", "_") & ".txt"
        Dim salvagedsFileNameAndBasePathNoSpace As String = sFileBasePath & "\" & salvagedsFileNameNoSpace
        Dim salvagedsFileNameNoSpaceInfo As FileInfo = New FileInfo(salvagedsFileNameAndBasePathNoSpace)
        Dim salvagedsFileName2 As String = "salvaged_2_" & sFileName
        Dim salvagedsFileNameNoSpace2 As String = salvagedsFileName2.Replace(" ", "_") & ".txt"
        Dim salvagedsFileNameAndBasePathNoSpace2 As String = sFileBasePath & "\" & salvagedsFileNameNoSpace2
        Dim salvagedsFileNameAndBasePathNoSpaceOld As String = sFileBasePath & "\old_" & salvagedsFileNameNoSpace
        Dim salvagedsFileNameAndBasePathNoSpaceOld2 As String = sFileBasePath & "\old_" & salvagedsFileNameNoSpace2
        Dim zipRepairedsFileNameNonZipExt As String = "zipRepaired_2_" & sFileName
        Dim zipRepairedsFileNameNonZipExtNoSpace As String = zipRepairedsFileNameNonZipExt.Replace(" ", "_")
        Dim zipRepairedsFileNameNonZipExtNoSpaceFullName As String = sFileBasePath & _
            "\" & zipRepairedsFileNameNonZipExtNoSpace
        Dim objWordFirstPassEverythingElse As Word.Application
        Dim objDocFirstPassEverythingElse As Word.Document
        Dim objWordAfterXMLLintFirstTreatment As Word.Application
        Dim objDocAfterXMLLintFirstTreatment As Word.Document
        Dim objWordAfterTruncatingFirst As Word.Application
        Dim objDocAfterTruncatingFirst As Word.Document

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        If Not Deal_With_Zip_Repair_AndPretty_Print_Routine(sFile) = "OK" Then

            Deal_With_Everything_Else = "False"

            Exit Function

        End If

        objWordFirstPassEverythingElse = New Word.Application
        objDocFirstPassEverythingElse = New Word.Document

        Try

            objDocFirstPassEverythingElse = objWordFirstPassEverythingElse.Documents.OpenNoRepairDialog(prettyPrintedDocumentXmlFullPath, OpenAndRepair:=True)
            objWordFirstPassEverythingElse.WindowState = Word.WdWindowState.wdWindowStateMaximize
            objWordFirstPassEverythingElse.Visible = True
            objWordFirstPassEverythingElse.Activate()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            'ProgressBar1.Hide()

            'MsgBox("That's weird, the file shouldn't open at this point. " _
            ' & "We have only altered it to show what line an error is at." _
            ' & "It should have opened prior to that.")

            Deal_With_Everything_Else = "OK"

            Exit Function

        Catch ex_objWordFirstPassEverythingElse As Exception

            CloseWordApplicationWithDocument(objWordFirstPassEverythingElse, objDocFirstPassEverythingElse)

            sFileExtension = sFileInfo.Extension

            'First we repair the zip as with the previous button.We start by getting the path from the textbox up top where  
            'we loaded the file. If the extension is .doc, .xls or ppt, it probably is not a docx, xlsx or ppts format file.

            If File.Exists(zipRepairedBasePathAndFileName) Then

                File.Delete(zipRepairedBasePathAndFileName)

            End If

            File.Copy(prettyPrintedDocumentXmlFullPath, zipRepairedBasePathAndFileName, True)

            'ProgressBar1.Value= 30

            'Now we extract the repaired file.

            If Directory.Exists(extractedRepairedZipOutputDirectory) Then

                Directory.Delete(extractedRepairedZipOutputDirectory, True)

            End If

            If Not Extract_Zip(zipRepairedBasePathAndFileName, extractedRepairedZipOutputDirectory) = "OK" Then

                MsgBox("Savvy DOCX Recovery could not extract the zip structure " _
                       & "of your target file. There is apparently nothing " _
                       & "recoverable but you can try Method 4, ""Salvage Text " _
                       & "or Data If Possible"", on the Recover Menu to be sure.")

                Deal_With_Everything_Else = "Failed"

                Exit Function

            End If

            If Directory.Exists(modifiedExtractedRepairedZipOutputDirectory) Then

                Directory.Delete(modifiedExtractedRepairedZipOutputDirectory, True)

            End If

            If Not Extract_Zip(zipRepairedBasePathAndFileName, modifiedExtractedRepairedZipOutputDirectory) = "OK" Then

                MsgBox("Savvy DOCX Recovery could not extract the zip structure " _
                       & "of your target file. There is apparently nothing " _
                       & "recoverable but you can try Method 4, ""Salvage Text " _
                       & "or Data If Possible"", on the Recover Menu to be sure.")

                Deal_With_Everything_Else = "Failed"

                Exit Function

            End If

            'ProgressBar1.Value= 35

            extractedCorruptFileDirPath = extractedRepairedZipOutputDirectory & "\"

            extractedCorruptFileDirInfo = New DirectoryInfo(extractedCorruptFileDirPath)
            corruptFileDirRetrievedFilesInfoArray = _
                extractedCorruptFileDirInfo.GetFiles("*.*", SearchOption.AllDirectories)

            'ProgressBar1.Value= 40

            Dim corruptFileDirRetrievedFileInfo As FileInfo
            Dim extractedCorruptFileDirInfoArrayCount As Integer = _
                corruptFileDirRetrievedFilesInfoArray.GetLength(0)
            Dim progressBarIncrement As Integer = 50 \ (extractedCorruptFileDirInfoArrayCount + 1)
            Dim corruptFileDirRetrievedFileName As String = Nothing
            Dim corruptFileXMLSubFileFullName As String = Nothing
            Dim modDocumentXmlFullPath As String = Nothing

            If progressBarIncrement = 0 Then

                progressBarIncrement = 1

            End If

            For Each corruptFileDirRetrievedFileInfo In corruptFileDirRetrievedFilesInfoArray

                corruptFileDirRetrievedFileName = corruptFileDirRetrievedFileInfo.Name
                corruptFileXMLSubFileFullName = corruptFileDirRetrievedFileInfo.FullName
                modDocumentXmlFullPath = modifiedExtractedRepairedZipOutputDirectory & _
                    corruptFileXMLSubFileFullName.Substring(extractedRepairedZipOutputDirectory.Length)

                If InStr(corruptFileDirRetrievedFileInfo.Extension, ".xml") Or _
                    InStr(corruptFileDirRetrievedFileInfo.Extension, ".rels") Then

                    If ProgressBar1.Value > 90 Then

                        'ProgressBar1.Value= 90

                    Else

                        'ProgressBar1.Value=ProgressBar1.Value+ progressBarIncrement

                    End If

                    Process_XML_Files_With_Xmllint_First(corruptFileXMLSubFileFullName, modDocumentXmlFullPath)

                    Continue For

                End If

            Next

            'Now we will rezip our directory and open as a docx.

            If File.Exists(zipExtensionXMLRepairedFullPath) Then

                File.Delete(zipExtensionXMLRepairedFullPath)

            End If

            If Not Re_Zip(modifiedExtractedRepairedZipOutputDirectory, zipExtensionXMLRepairedFullPath) = "OK" Then

                MsgBox("You docx file's zip structure was presumably repaired and the file's " _
                       & "word/document.xml subfile was prepared for providing the location " _
                       & "of the error in case it was not specified in ther opening error " _
                       & "message. However the file could not be rezipped. This is unusual. " _
                       & "Try again or try individual recovery methods from the Recover menu.")

                Deal_With_Everything_Else = "Failed"

                Exit Function

            End If

            If File.Exists(zipExtensionXMLRepairedFullPath) Then

                File.Copy(zipExtensionXMLRepairedFullPath, xmlRepairedFullPath, True)

            End If

            objWordAfterXMLLintFirstTreatment = New Word.Application
            objDocAfterXMLLintFirstTreatment = New Word.Document

            Try

                'The next few lines bypass the section where the corrupt 
                'file is copied to one whose name is a random number.

                If File.Exists(originalFileFullPathRepairedName) Then

                    File.Delete(originalFileFullPathRepairedName)

                End If

                File.Copy(xmlRepairedFullPath, originalFileFullPathRepairedName, True)

                If File.Exists(xmlRepairedFullPath) Then

                    File.Delete(xmlRepairedFullPath)

                End If

                'Hopefully xmllint fixed any corruption, if not, next stop is bad xml file truncation.

                objDocAfterXMLLintFirstTreatment = objWordAfterXMLLintFirstTreatment.Documents.OpenNoRepairDialog(originalFileFullPathRepairedName, OpenAndRepair:=True)
                objWordAfterXMLLintFirstTreatment.Visible = True
                objWordAfterXMLLintFirstTreatment.WindowState = Word.WdWindowState.wdWindowStateMaximize
                objWordAfterXMLLintFirstTreatment.Application.DisplayFullScreen = False
                objWordAfterXMLLintFirstTreatment.Activate()

                Me.Cursor = System.Windows.Forms.Cursors.Default
                'ProgressBar1.Hide()

                Deal_With_Everything_Else = "OK"

                Exit Function

            Catch ex_AfterXMLLintTreatmentFirst As Exception

                CloseWordApplicationWithDocument(objWordAfterXMLLintFirstTreatment, objDocAfterXMLLintFirstTreatment)

                If File.Exists(zipRepairedBasePathAndFileName) Then

                    File.Delete(zipRepairedBasePathAndFileName)

                End If

                File.Copy(prettyPrintedDocumentXmlFullPath, zipRepairedBasePathAndFileName, True)

                'ProgressBar1.Value= 30

                'Now we extract the repaired file.

                If Directory.Exists(extractedRepairedZipOutputDirectory) Then

                    Directory.Delete(extractedRepairedZipOutputDirectory, True)

                End If

                If Not Extract_Zip(zipRepairedBasePathAndFileName, extractedRepairedZipOutputDirectory) = "OK" Then

                    MsgBox("Savvy DOCX Recovery could not extract the zip structure " _
                           & "of your target file. There is apparently nothing " _
                           & "recoverable but you can try Method 4, ""Salvage Text " _
                           & "or Data If Possible"", on the Recover Menu to be sure.")

                    Deal_With_Everything_Else = "Failed"

                    Exit Function

                End If

                If Directory.Exists(modifiedExtractedRepairedZipOutputDirectory) Then

                    Directory.Delete(modifiedExtractedRepairedZipOutputDirectory, True)

                End If

                If Not Extract_Zip(zipRepairedBasePathAndFileName, modifiedExtractedRepairedZipOutputDirectory) = "OK" Then

                    MsgBox("Savvy DOCX Recovery could not extract the zip structure " _
                           & "of your target file. There is apparently nothing " _
                           & "recoverable but you can try Method 4, ""Salvage Text " _
                           & "or Data If Possible"", on the Recover Menu to be sure.")

                    Deal_With_Everything_Else = "Failed"

                    Exit Function

                End If

                'ProgressBar1.Value= 35

                If File.Exists(zipRepairedBasePathAndFileName) Then

                    File.Delete(zipRepairedBasePathAndFileName)

                End If

                extractedCorruptFileDirPath = extractedRepairedZipOutputDirectory & "\"

                extractedCorruptFileDirInfo = New DirectoryInfo(extractedCorruptFileDirPath)
                corruptFileDirRetrievedFilesInfoArray = _
                    extractedCorruptFileDirInfo.GetFiles("*.*", SearchOption.AllDirectories)

                'ProgressBar1.Value= 40

                If progressBarIncrement = 0 Then

                    progressBarIncrement = 1

                End If

                For Each corruptFileDirRetrievedFileInfo In corruptFileDirRetrievedFilesInfoArray

                    corruptFileDirRetrievedFileName = corruptFileDirRetrievedFileInfo.Name
                    corruptFileXMLSubFileFullName = corruptFileDirRetrievedFileInfo.FullName
                    modDocumentXmlFullPath = modifiedExtractedRepairedZipOutputDirectory & _
                        corruptFileXMLSubFileFullName.Substring(extractedRepairedZipOutputDirectory.Length)

                    If InStr(corruptFileDirRetrievedFileInfo.Extension, ".xml") Or _
                        InStr(corruptFileDirRetrievedFileInfo.Extension, ".rels") Then

                        If ProgressBar1.Value > 90 Then

                            'ProgressBar1.Value= 90

                        Else

                            'ProgressBar1.Value=ProgressBar1.Value+ progressBarIncrement

                        End If

                        Process_XML_Files_With_Truncator_First(modDocumentXmlFullPath)

                        Continue For

                    End If

                Next

                'Now we will rezip our directory and open as a docx.

                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                    File.Delete(zipExtensionXMLRepairedFullPath)

                End If

                If Not Re_Zip(modifiedExtractedRepairedZipOutputDirectory, zipExtensionXMLRepairedFullPath) = "OK" Then

                    MsgBox("You docx file's zip structure was presumably repaired and the file's " _
                           & "word/document.xml subfile was prepared for providing the location " _
                           & "of the error in case it was not specified in ther opening error " _
                           & "message. However the file could not be rezipped. This is unusual. " _
                           & "Try again or try individual recovery methods from the Recover menu.")

                    Deal_With_Everything_Else = "Failed"

                    Exit Function

                End If

                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                    File.Copy(zipExtensionXMLRepairedFullPath, xmlRepairedFullPath, True)

                End If

                objWordAfterTruncatingFirst = New Word.Application
                objDocAfterTruncatingFirst = New Word.Document

                Try

                    'The next few lines bypass the section where the corrupt 
                    'file is copied to one whose name is a random number.

                    If File.Exists(originalFileFullPathRepairedName) Then

                        File.Delete(originalFileFullPathRepairedName)

                    End If

                    File.Copy(xmlRepairedFullPath, originalFileFullPathRepairedName, True)

                    objDocAfterTruncatingFirst = objWordAfterTruncatingFirst.Documents.OpenNoRepairDialog(originalFileFullPathRepairedName, OpenAndRepair:=True)
                    objWordAfterTruncatingFirst.WindowState = Word.WdWindowState.wdWindowStateMaximize
                    objWordAfterTruncatingFirst.Visible = True
                    objWordAfterTruncatingFirst.Activate()
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    'ProgressBar1.Hide()

                    Deal_With_Everything_Else = "OK"

                    Exit Function

                Catch ex_AfterTruncatingFirst As Exception

                    CloseWordApplicationWithDocument(objWordAfterTruncatingFirst, objDocAfterTruncatingFirst)

                    'If this truncated XML files didn't please Word, we will now try to extract
                    'the text with two different setting of SilverCoder's DoToText. The first setting
                    'claims to fix the XML in the document.xml. The second, simply removes the XML tags 
                    'from the document.xml file. What is useful here is most recovery programs
                    'seem to truncate rather than give you all the document.xml warts and all.
                    'DoctotText will give you the good xml at the end or middle of the file next to the
                    'bad stuff and allow you to do the surgery.

                    MsgBox("The Savvy DOCX Recovery failed to recover your file as a Word file (with possibly both text " _
                        & "and formatting). You can try separately to have the program salvage just the text as a text " _
                        & "file using the command line program made by SilverCoders called DocToText. Select the 4th choice " _
                        & "on the Recover menu, ""Method 4 - Salvage Text or Data If Possible"".", MsgBoxStyle.Exclamation)

                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    'ProgressBar1.Hide()

                    Deal_With_Everything_Else = "Failed"

                    Exit Function

                Finally

                    GC.Collect()
                    GC.WaitForPendingFinalizers()

                    ' GC needs to be called twice in order to get the Finalizers called  
                    ' - the first time in, it simply makes a list of what is to be  
                    ' finalized, the second time in, it actually is finalizing. Only  
                    ' then will the object do its automatic ReleaseComObject. 

                    GC.Collect()
                    GC.WaitForPendingFinalizers()

                End Try

            Finally

                GC.Collect()
                GC.WaitForPendingFinalizers()

                ' GC needs to be called twice in order to get the Finalizers called  
                ' - the first time in, it simply makes a list of what is to be  
                ' finalized, the second time in, it actually is finalizing. Only  
                ' then will the object do its automatic ReleaseComObject. 

                GC.Collect()
                GC.WaitForPendingFinalizers()

            End Try

        Finally

            Try

                'MsgBox("Going through EverythingElse Finally.")
                Me.Cursor = System.Windows.Forms.Cursors.Default

                If Directory.Exists(extractedRepairedZipOutputDirectory) Then

                    Directory.Delete(extractedRepairedZipOutputDirectory, True)

                End If

                If Directory.Exists(modifiedExtractedRepairedZipOutputDirectory) Then

                    Directory.Delete(modifiedExtractedRepairedZipOutputDirectory, True)

                End If

                If File.Exists(sFile) Then

                    File.Delete(sFile)

                End If

                If File.Exists(sFileZip) Then

                    File.Delete(sFileZip)

                End If

                If File.Exists(sFileZipPrelimRepair) Then

                    File.Delete(sFileZipPrelimRepair)

                End If

                If File.Exists(zipRepairedBasePathAndFileName) Then

                    File.Delete(zipRepairedBasePathAndFileName)

                End If

                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                    File.Delete(zipExtensionXMLRepairedFullPath)

                End If

                If File.Exists(xmlRepairedFullPath) Then

                    File.Delete(xmlRepairedFullPath)

                End If

                If File.Exists(zipRepairedBasePathAndFileName) Then

                    File.Delete(zipRepairedBasePathAndFileName)

                End If


                If File.Exists(prettyPrintedDocumentXmlFullPath) Then

                    File.Delete(prettyPrintedDocumentXmlFullPath)

                End If
            Catch

            End Try

            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' GC needs to be called twice in order to get the Finalizers called  
            ' - the first time in, it simply makes a list of what is to be  
            ' finalized, the second time in, it actually is finalizing. Only  
            ' then will the object do its automatic ReleaseComObject. 

            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try

    End Function

    Private Function Deal_With_Everything_Else_After_Unspecified_Error_Fix(preliminaryXmlRepairedFullPath As String, _
            extractedRepairedZipOutputDirectory As String, modifiedextractedRepairedZipOutputDirectory As String)

        Dim sFile As String = preliminaryXmlRepairedFullPath
        Dim sFileZip As String = sFile & ".zip"
        Dim sFileZipPrelimRepair As String = sFileZip & ".prelim_repair.zip"
        Dim xmlValidateReader As StreamReader = Nothing
        Dim xmlValidateErrorReader As StreamReader = Nothing
        Dim xmlValidateReader2 As StreamReader = Nothing
        Dim xmlValidateErrorReader2 As StreamReader = Nothing
        Dim byteMatch As Match = Nothing
        Dim extractedCorruptFileDirInfo As DirectoryInfo
        Dim corruptFileDirRetrievedFilesInfoArray As FileInfo()
        Dim sFileInfo As FileInfo = New FileInfo(sFile)
        Dim byteErrorLocationInteger As Integer = 0
        Dim truncatedLength As Integer = 0
        Dim intTruncationAmount As Integer = 0
        Dim xmlValidateCompOut As String = Nothing
        Dim xmlValidateErrorOut As String = Nothing
        Dim xmlValidateArguments As String = Nothing
        Dim byteMatchString As String = Nothing
        Dim truncatedLengthAsString As String = Nothing
        Dim truncateArguments As String = Nothing
        Dim xmlValidateCompOut2 As String = Nothing
        Dim xmlValidateErrorOut2 As String = Nothing
        Dim xmlRecoverArguments As String = Nothing
        Dim oldXMLRepairedFileName As String = Nothing
        Dim oldXMLRepairedFullPath As String = Nothing
        Dim sevenZipUpArguments As String = Nothing
        Dim sFileExtension As String = Nothing
        Dim sevenZipFullPath As String = Nothing
        Dim sevenZipExtractArguments As String = Nothing
        Dim extractedCorruptFileDirPath As String = Nothing
        Dim modExtractedCorruptFileDirPath As String = Nothing
        Dim repairZipArguments As String = Nothing
        Dim zipFullPath As String = Nothing
        Dim byteErrorLocation As String = Nothing
        Dim sFileName As String = LCase(sFileInfo.Name)
        Dim sFileFullNameLC As String = LCase(sFileInfo.FullName)
        Dim zipRepairedsFileName As String = "zipRepaired_" & sFileName & ".zip"
        'Dim sFileBasePath As String = LCase(sFileInfo.Directory.ToString)
        Dim sFileBasePath As String = tempPath
        Dim xmlRepairedFileName As String = "xml_error_repaired_" & sFileName
        Dim xmlRepairedFullPath As String = sFileBasePath & "\" & xmlRepairedFileName
        Dim originalFileFullPath As String = TextBox1.Text
        Dim originalFileFullPathRepairedName As String = originalFileFullPath & ".savvy_fix.docx"
        Dim zipExtensionXMLRepairedFullPath As String = xmlRepairedFullPath & ".zip"
        Dim zipRepairedBasePathAndFileName As String = sFileBasePath & "\" & zipRepairedsFileName
        Dim zipRepairedBasePathAndFileNameIndexLastPeriod As Integer = zipRepairedBasePathAndFileName.LastIndexOf(".")
        Dim prettyPrintedDocumentXmlsFile As String = "pretty_printed_" & sFileName
        Dim prettyPrintedDocumentXmlFullPath As String = sFileBasePath & "\" & prettyPrintedDocumentXmlsFile
        Dim corruptDocumentXMLPath As String = extractedRepairedZipOutputDirectory & "\word\document.xml"
        Dim salvagedsFileName As String = "salvaged_" & sFileName
        Dim salvagedsFileNameNoSpace As String = salvagedsFileName.Replace(" ", "_") & ".txt"
        Dim salvagedsFileNameAndBasePathNoSpace As String = sFileBasePath & "\" & salvagedsFileNameNoSpace
        Dim salvagedsFileNameNoSpaceInfo As FileInfo = New FileInfo(salvagedsFileNameAndBasePathNoSpace)
        Dim salvagedsFileName2 As String = "salvaged_2_" & sFileName
        Dim salvagedsFileNameNoSpace2 As String = salvagedsFileName2.Replace(" ", "_") & ".txt"
        Dim salvagedsFileNameAndBasePathNoSpace2 As String = sFileBasePath & "\" & salvagedsFileNameNoSpace2
        Dim salvagedsFileNameAndBasePathNoSpaceOld As String = sFileBasePath & "\old_" & salvagedsFileNameNoSpace
        Dim salvagedsFileNameAndBasePathNoSpaceOld2 As String = sFileBasePath & "\old_" & salvagedsFileNameNoSpace2
        Dim zipRepairedsFileNameNonZipExt As String = "zipRepaired_2_" & sFileName
        Dim zipRepairedsFileNameNonZipExtNoSpace As String = zipRepairedsFileNameNonZipExt.Replace(" ", "_")
        Dim zipRepairedsFileNameNonZipExtNoSpaceFullName As String = sFileBasePath & _
            "\" & zipRepairedsFileNameNonZipExtNoSpace
        Dim objWordAfterXMLLintFirstTreatment As Word.Application
        Dim objDocAfterXMLLintFirstTreatment As Word.Document
        Dim objWordAfterTruncatingFirst As Word.Application
        Dim objDocAfterTruncatingFirst As Word.Document

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        sFileExtension = sFileInfo.Extension

        'First we repair the zip as with the previous button.We start by getting the path from the textbox up top where  
        'we loaded the file. If the extension is .doc, .xls or ppt, it probably is not a docx, xlsx or ppts format file.

        If File.Exists(zipRepairedBasePathAndFileName) Then

            File.Delete(zipRepairedBasePathAndFileName)

        End If

        File.Copy(preliminaryXmlRepairedFullPath, zipRepairedBasePathAndFileName, True)

        ProgressBar1.Value = 30

        'Now we extract the repaired file.

        'If Directory.Exists(extractedRepairedZipOutputDirectory) Then

        '    Directory.Delete(extractedRepairedZipOutputDirectory, True)

        'End If

        If Not Extract_Zip(zipRepairedBasePathAndFileName, extractedRepairedZipOutputDirectory) = "OK" Then

            MsgBox("Savvy DOCX Recovery could not extract the zip structure " _
                   & "of your target file. There is apparently nothing " _
                   & "recoverable but you can try the ""Try Salvaging " _
                   & "Just Text "" choice on the recover Menu to be sure.")

            Deal_With_Everything_Else_After_Unspecified_Error_Fix = "Failed"

            Exit Function

        End If

        If Directory.Exists(modifiedextractedRepairedZipOutputDirectory) Then

            Directory.Delete(modifiedextractedRepairedZipOutputDirectory, True)

        End If

        If Not Extract_Zip(zipRepairedBasePathAndFileName, modifiedextractedRepairedZipOutputDirectory) = "OK" Then

            MsgBox("Savvy DOCX Recovery could not extract the zip structure " _
                   & "of your target file. There is apparently nothing " _
                   & "recoverable but you can try the ""Try Salvaging " _
                   & "Just Text"" choice on the Recover menu to be sure.")

            Deal_With_Everything_Else_After_Unspecified_Error_Fix = "Failed"

            Exit Function

        End If

        ProgressBar1.Value = 35

        extractedCorruptFileDirPath = extractedRepairedZipOutputDirectory & "\"

        extractedCorruptFileDirInfo = New DirectoryInfo(extractedCorruptFileDirPath)
        corruptFileDirRetrievedFilesInfoArray = _
            extractedCorruptFileDirInfo.GetFiles("*.*", SearchOption.AllDirectories)

        ProgressBar1.Value = 40

        Dim corruptFileDirRetrievedFileInfo As FileInfo
        Dim extractedCorruptFileDirInfoArrayCount As Integer = _
            corruptFileDirRetrievedFilesInfoArray.GetLength(0)
        Dim progressBarIncrement As Integer = 50 \ (extractedCorruptFileDirInfoArrayCount + 1)
        Dim corruptFileDirRetrievedFileName As String = Nothing
        Dim corruptFileXMLSubFileFullName As String = Nothing
        Dim modDocumentXmlFullPath As String = Nothing

        If progressBarIncrement = 0 Then

            progressBarIncrement = 1

        End If

        For Each corruptFileDirRetrievedFileInfo In corruptFileDirRetrievedFilesInfoArray

            corruptFileDirRetrievedFileName = corruptFileDirRetrievedFileInfo.Name
            corruptFileXMLSubFileFullName = corruptFileDirRetrievedFileInfo.FullName
            modDocumentXmlFullPath = modifiedextractedRepairedZipOutputDirectory & _
                corruptFileXMLSubFileFullName.Substring(extractedRepairedZipOutputDirectory.Length)

            If InStr(corruptFileDirRetrievedFileInfo.Extension, ".xml") Or _
                InStr(corruptFileDirRetrievedFileInfo.Extension, ".rels") Then

                Process_XML_Files_With_Xmllint_First(corruptFileXMLSubFileFullName, modDocumentXmlFullPath)

                Continue For

            End If

        Next

        'Now we will rezip our directory and open as a docx.

        If File.Exists(zipExtensionXMLRepairedFullPath) Then

            File.Delete(zipExtensionXMLRepairedFullPath)

        End If

        If Not Re_Zip(modifiedextractedRepairedZipOutputDirectory, zipExtensionXMLRepairedFullPath) = "OK" Then

            MsgBox("You docx file's zip structure was presumably repaired and the file's " _
                   & "word/document.xml subfile was prepared for providing the location " _
                   & "of the error in case it was not specified in ther opening error " _
                   & "message. However the file could not be rezipped. This is unusual. " _
                   & "Try again or try individual recovery methods from the Recover menu.")

            Deal_With_Everything_Else_After_Unspecified_Error_Fix = "Failed"

            Exit Function

        End If

        If File.Exists(zipExtensionXMLRepairedFullPath) Then

            File.Copy(zipExtensionXMLRepairedFullPath, xmlRepairedFullPath, True)

        End If

        objWordAfterXMLLintFirstTreatment = New Word.Application
        objDocAfterXMLLintFirstTreatment = New Word.Document

        Try

            'The next few lines bypass the section where the corrupt 
            'file is copied to one whose name is a random number.

            If File.Exists(originalFileFullPathRepairedName) Then

                File.Delete(originalFileFullPathRepairedName)

            End If

            File.Copy(xmlRepairedFullPath, originalFileFullPathRepairedName, True)

            If File.Exists(xmlRepairedFullPath) Then

                File.Delete(xmlRepairedFullPath)

            End If

            'Hopefully xmllint fixed any corruption, if not, next stop is bad xml file truncation.

            objDocAfterXMLLintFirstTreatment = objWordAfterXMLLintFirstTreatment.Documents.OpenNoRepairDialog(originalFileFullPathRepairedName, OpenAndRepair:=True)
            objWordAfterXMLLintFirstTreatment.Visible = True
            objWordAfterXMLLintFirstTreatment.WindowState = Word.WdWindowState.wdWindowStateMaximize
            'objWordAfterXMLLintFirstTreatment.Application.DisplayFullScreen = False
            objWordAfterXMLLintFirstTreatment.Activate()

            Me.Cursor = System.Windows.Forms.Cursors.Default
            'ProgressBar1.Hide()

            Deal_With_Everything_Else_After_Unspecified_Error_Fix = "OK"

            Exit Function

        Catch ex_AfterXMLLintTreatmentFirst As Exception

            CloseWordApplicationWithDocument(objWordAfterXMLLintFirstTreatment, objDocAfterXMLLintFirstTreatment)

            If File.Exists(zipRepairedBasePathAndFileName) Then

                File.Delete(zipRepairedBasePathAndFileName)

            End If

            File.Copy(preliminaryXmlRepairedFullPath, zipRepairedBasePathAndFileName, True)

            'ProgressBar1.Value= 30

            'Now we extract the repaired file.

            'If Directory.Exists(extractedRepairedZipOutputDirectory) Then

            '    Directory.Delete(extractedRepairedZipOutputDirectory, True)

            'End If

            If Not Extract_Zip(zipRepairedBasePathAndFileName, extractedRepairedZipOutputDirectory) = "OK" Then

                MsgBox("Savvy DOCX Recovery could not extract the zip structure " _
                       & "of your target file. There is apparently nothing " _
                       & "recoverable but you can try the ""Try Salvaging " _
                       & "Just Text"" choice on the Recover Menu to be sure.")

                Deal_With_Everything_Else_After_Unspecified_Error_Fix = "Failed"

                Exit Function

            End If

            If Directory.Exists(modifiedextractedRepairedZipOutputDirectory) Then

                Directory.Delete(modifiedextractedRepairedZipOutputDirectory, True)

            End If

            If Not Extract_Zip(zipRepairedBasePathAndFileName, modifiedextractedRepairedZipOutputDirectory) = "OK" Then

                MsgBox("Savvy DOCX Recovery could not extract the zip structure " _
                       & "of your target file. There is apparently nothing " _
                       & "recoverable but you can try the, ""Try Salvaging Just " _
                       & "Text"" choice on the Recover Menu to be sure.")

                Deal_With_Everything_Else_After_Unspecified_Error_Fix = "Failed"

                Exit Function

            End If

            'ProgressBar1.Value= 35

            If File.Exists(zipRepairedBasePathAndFileName) Then

                File.Delete(zipRepairedBasePathAndFileName)

            End If

            extractedCorruptFileDirPath = extractedRepairedZipOutputDirectory & "\"

            extractedCorruptFileDirInfo = New DirectoryInfo(extractedCorruptFileDirPath)
            corruptFileDirRetrievedFilesInfoArray = _
                extractedCorruptFileDirInfo.GetFiles("*.*", SearchOption.AllDirectories)

            For Each corruptFileDirRetrievedFileInfo In corruptFileDirRetrievedFilesInfoArray

                corruptFileDirRetrievedFileName = corruptFileDirRetrievedFileInfo.Name
                corruptFileXMLSubFileFullName = corruptFileDirRetrievedFileInfo.FullName
                modDocumentXmlFullPath = modifiedextractedRepairedZipOutputDirectory & _
                    corruptFileXMLSubFileFullName.Substring(extractedRepairedZipOutputDirectory.Length)

                If InStr(corruptFileDirRetrievedFileInfo.Extension, ".xml") Or _
                    InStr(corruptFileDirRetrievedFileInfo.Extension, ".rels") Then

                    Process_XML_Files_With_Truncator_First(modDocumentXmlFullPath)

                    Continue For

                End If

            Next

            'Now we will rezip our directory and open as a docx.

            If File.Exists(zipExtensionXMLRepairedFullPath) Then

                File.Delete(zipExtensionXMLRepairedFullPath)

            End If

            If Not Re_Zip(modifiedextractedRepairedZipOutputDirectory, zipExtensionXMLRepairedFullPath) = "OK" Then

                MsgBox("You docx file's zip structure was presumably repaired and the file's " _
                       & "word/document.xml subfile was prepared for providing the location " _
                       & "of the error in case it was not specified in ther opening error " _
                       & "message. However the file could not be rezipped. This is unusual. " _
                       & "Try again or try individual recovery methods from the Recover menu.")

                Deal_With_Everything_Else_After_Unspecified_Error_Fix = "Failed"

                Exit Function

            End If

            If File.Exists(zipExtensionXMLRepairedFullPath) Then

                File.Copy(zipExtensionXMLRepairedFullPath, xmlRepairedFullPath, True)

            End If

            objWordAfterTruncatingFirst = New Word.Application
            objDocAfterTruncatingFirst = New Word.Document

            Try

                'The next few lines bypass the section where the corrupt 
                'file is copied to one whose name is a random number.

                If File.Exists(originalFileFullPathRepairedName) Then

                    File.Delete(originalFileFullPathRepairedName)

                End If

                File.Copy(xmlRepairedFullPath, originalFileFullPathRepairedName, True)

                objDocAfterTruncatingFirst = objWordAfterTruncatingFirst.Documents.OpenNoRepairDialog(originalFileFullPathRepairedName, OpenAndRepair:=True)
                objWordAfterTruncatingFirst.WindowState = Word.WdWindowState.wdWindowStateMaximize
                objWordAfterTruncatingFirst.Visible = True
                objWordAfterTruncatingFirst.Activate()
                Me.Cursor = System.Windows.Forms.Cursors.Default
                'ProgressBar1.Hide()

                Deal_With_Everything_Else_After_Unspecified_Error_Fix = "OK"

                Exit Function

            Catch ex_AfterTruncatingFirst As Exception

                CloseWordApplicationWithDocument(objWordAfterTruncatingFirst, objDocAfterTruncatingFirst)

                'If this truncated XML files didn't please Word, we will now try to extract
                'the text with two different setting of SilverCoder's DoToText. The first setting
                'claims to fix the XML in the document.xml. The second, simply removes the XML tags 
                'from the document.xml file. What is useful here is most recovery programs
                'seem to truncate rather than give you all the document.xml warts and all.
                'DoctotText will give you the good xml at the end or middle of the file next to the
                'bad stuff and allow you to do the surgery.

                MsgBox("The Savvy DOCX Recovery failed to recover your file as a Word file (with possibly both text " _
                    & "and formatting). You can try separately to have the program salvage just the text as a text " _
                    & "file using the command line program made by SilverCoders called DocToText. Select the 2nd choice " _
                    & "on the Recover menu, ""Try Salvaging Just Text"".", MsgBoxStyle.Exclamation)

                Me.Cursor = System.Windows.Forms.Cursors.Default
                'ProgressBar1.Hide()

                Deal_With_Everything_Else_After_Unspecified_Error_Fix = "Failed"

                Exit Function

            Finally

                GC.Collect()
                GC.WaitForPendingFinalizers()

                ' GC needs to be called twice in order to get the Finalizers called  
                ' - the first time in, it simply makes a list of what is to be  
                ' finalized, the second time in, it actually is finalizing. Only  
                ' then will the object do its automatic ReleaseComObject. 

                GC.Collect()
                GC.WaitForPendingFinalizers()

            End Try

        Finally

            Try

                'MsgBox("Going through EverythingElse Finally.")
                Me.Cursor = System.Windows.Forms.Cursors.Default

                'If Directory.Exists(extractedRepairedZipOutputDirectory) Then

                '    Directory.Delete(extractedRepairedZipOutputDirectory, True)

                'End If

                If Directory.Exists(modifiedextractedRepairedZipOutputDirectory) Then

                    Directory.Delete(modifiedextractedRepairedZipOutputDirectory, True)

                End If

                If File.Exists(sFile) Then

                    File.Delete(sFile)

                End If

                If File.Exists(sFileZip) Then

                    File.Delete(sFileZip)

                End If

                If File.Exists(sFileZipPrelimRepair) Then

                    File.Delete(sFileZipPrelimRepair)

                End If

                If File.Exists(zipRepairedBasePathAndFileName) Then

                    File.Delete(zipRepairedBasePathAndFileName)

                End If

                If File.Exists(zipExtensionXMLRepairedFullPath) Then

                    File.Delete(zipExtensionXMLRepairedFullPath)

                End If

                If File.Exists(xmlRepairedFullPath) Then

                    File.Delete(xmlRepairedFullPath)

                End If

                If File.Exists(zipRepairedBasePathAndFileName) Then

                    File.Delete(zipRepairedBasePathAndFileName)

                End If


                If File.Exists(prettyPrintedDocumentXmlFullPath) Then

                    File.Delete(prettyPrintedDocumentXmlFullPath)

                End If
            Catch ex As Exception


            End Try

            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' GC needs to be called twice in order to get the Finalizers called  
            ' - the first time in, it simply makes a list of what is to be  
            ' finalized, the second time in, it actually is finalizing. Only  
            ' then will the object do its automatic ReleaseComObject. 

            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try

    End Function

    Private Sub Deal_With_Salvage_Text_Treatment_Menu_Choice()

        selectedFile = TextBox1.Text
        Dim rand1 As New Random()
        Dim number9Places As Integer = rand1.Next(1, 1000000000)
        Dim selectedFileInfo As FileInfo = New FileInfo(selectedFile)
        Dim selectedFileName As String = LCase(selectedFileInfo.Name)
        Dim selectedFileExtension As String = LCase(selectedFileInfo.Extension)
        Dim selectedFileBasePath As String = LCase(selectedFileInfo.Directory.ToString)
        Dim sFile As String = selectedFileBasePath & "\" & number9Places.ToString & selectedFileExtension
        File.Copy(selectedFile, sFile, True)
        'MsgBox("My random file full path is: " & sFile)
        Dim sFileInfo As FileInfo = New FileInfo(sFile)
        Dim sFileExtension As String = LCase(sFileInfo.Extension)
        Dim x As Integer = 0
        Dim sFileName As String = LCase(sFileInfo.Name)
        Dim sFileNameNoSpace As String = sFileName.Replace(" ", "_")
        'Dim sFileBasePath As String = LCase(sFileInfo.Directory.ToString)
        Dim sFileBasePath As String = tempPath
        Dim sFileFullNameLC As String = LCase(sFileInfo.FullName.ToString)
        Dim originalFileFullPath As String = TextBox1.Text
        Dim originalFileFullPathRepairedName As String = originalFileFullPath & ".savvy_fix.docx"
        Dim sFileNameWithoutExtension As String = sFileName.Substring(0, sFileInfo.Name.Length - 5)
        Dim sFileNameChangedtoDocExtension = sFileNameWithoutExtension & ".doc"
        Dim sFileBasePathAndNameChangedtoDocExtension = sFileBasePath & "\" & sFileNameChangedtoDocExtension
        Dim prettyPrintedDocumentXmlsFile As String = "pretty_printed_" & sFileName
        Dim prettyPrintedDocumentXmlFullPath As String = sFileBasePath & "\" & prettyPrintedDocumentXmlsFile
        Dim sFileZip As String = sFile & ".zip"
        Dim zipRepairedsFileName As String = "zipRepaired_" & sFileName & ".zip"
        Dim xmlRepairedFileName As String = "xml_error_repaired_" & sFileName
        Dim xmlRepairedFullPath As String = sFileBasePath & "\" & xmlRepairedFileName
        Dim zipExtensionXMLRepairedFullPath As String = xmlRepairedFullPath & ".zip"
        Dim zipRepairedBasePathAndFileName As String = sFileBasePath & "\" & zipRepairedsFileName
        Dim zipRepairedBasePathAndFileNameIndexLastPeriod As Integer = zipRepairedBasePathAndFileName.LastIndexOf(".")
        Dim extractedRepairedZipOutputDirectory As String = _
            zipRepairedBasePathAndFileName.Remove(zipRepairedBasePathAndFileNameIndexLastPeriod - 5)
        Dim modifiedextractedRepairedZipOutputDirectory As String = extractedRepairedZipOutputDirectory & "_mod"
        Dim salvagedsFileName As String = "salvaged_" & sFileName
        Dim salvagedsFileName2 As String = "salvaged_2_" & sFileName
        Dim salvagedsFileNameNoSpace As String = salvagedsFileName.Replace(" ", "_") & ".txt"
        Dim salvagedsFileNameNoSpace2 As String = salvagedsFileName2.Replace(" ", "_") & ".txt"
        Dim salvagedsFileNameAndBasePathNoSpace As String = sFileBasePath & "\" & salvagedsFileNameNoSpace
        Dim salvagedsFileNameAndBasePathNoSpace2 As String = sFileBasePath & "\" & salvagedsFileNameNoSpace2
        Dim salvagedsFileNameAndBasePathNoSpaceOld As String = sFileBasePath & "\old_" & salvagedsFileNameNoSpace
        Dim salvagedsFileNameAndBasePathNoSpaceOld2 As String = sFileBasePath & "\old_" & salvagedsFileNameNoSpace2
        Dim zipRepairedsFileNameNonZipExt As String = "zipRepaired_2_" & sFileName
        Dim zipRepairedsFileNameNonZipExtNoSpace As String = zipRepairedsFileNameNonZipExt.Replace(" ", "_")
        Dim zipRepairedsFileNameNonZipExtNoSpaceFullName As String = sFileBasePath & _
            "\" & zipRepairedsFileNameNonZipExtNoSpace

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        If String.IsNullOrEmpty(TextBox1.Text) = True Then

            SelectFile()

        Else

            ' initializes and sets up parameters of progress bar.
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

            'ProgressBar1.Style = ProgressBarStyle.Marquee
            'ProgressBar1.MarqueeAnimationSpeed = 100
            'ProgressBar1.Show()

            'Dim byteMatch As Match = Regex.Match(Trid_Identify(sFileFullNameLC), "\(.DOC\)")

            'If String.IsNullOrEmpty(byteMatch.ToString) = False Then

            '    MsgBox("Your file may not have the right extension. We are changing .docx " _
            '           & "to .doc on a copy of the file and then will attempt to open it.")

            '    File.Copy(sFileFullNameLC, sFileBasePathAndNameChangedtoDocExtension, True)

            '    Dim objWord As New Word.Application
            '    Dim objDoc As New Word.Document

            '    Try

            '        'Will try to open Word file with the Open and Repair feature.
            '        'If successful there's no need to do more and it makes Word visible
            '        'even though it was started invisibly. If unsuccessful the program 
            '        'moves to the first catch section.

            '        objWord.Visible = False
            '        objDoc = objWord.Documents.OpenNoRepairDialog(sFileBasePathAndNameChangedtoDocExtension, OpenAndRepair:=True)
            '        objWord.Visible = True
            '        Me.Cursor = System.Windows.Forms.Cursors.Default
            '        'ProgressBar1.Hide()

            '    Catch

            '        Impl_S2RecoverText_From_DOC(sFileBasePathAndNameChangedtoDocExtension)

            '        Me.Cursor = System.Windows.Forms.Cursors.Default
            '        'ProgressBar1.Hide()

            '    Finally

            '        If File.Exists(sFile) Then

            '            File.Delete(sFile)

            '        End If

            '        If Not objDoc Is Nothing Then

            '            objDoc.Close(False)
            '            Marshal.FinalReleaseComObject(objDoc)
            '            objDoc = Nothing

            '        End If

            '        If Not objWord Is Nothing Then

            '            objWord.Quit(False)
            '            Marshal.FinalReleaseComObject(objWord)
            '            objWord = Nothing

            '        End If

            '        'Garbage collecting code along with the Marshal.FinalReleaseComObject 
            '        'necessary to remove the instances of Word started by the program.

            '        GC.Collect()
            '        GC.WaitForPendingFinalizers()

            '        ' GC needs to be called twice in order to get the Finalizers called  
            '        ' - the first time in, it simply makes a list of what is to be  
            '        ' finalized, the second time in, it actually is finalizing. Only  
            '        ' then will the object do its automatic ReleaseComObject. 

            '        GC.Collect()
            '        GC.WaitForPendingFinalizers()

            '    End Try

            'Else

            '    'Creates Word objects needed to open our target corrupt file

            '    Dim objWord As New Word.Application
            '    Dim objDoc As New Word.Document
            '    ProgressBar1.Show()
            '   ProgressBar1.Value= 50

            '    Try

            '        'The next few lines bypass the section where the corrupt 
            '        'file is copied to one whose name is a random number.

            '        If File.Exists(originalFileFullPathRepairedName) Then

            '            File.Delete(originalFileFullPathRepairedName)

            '        End If

            '        File.Copy(sFileFullNameLC, originalFileFullPathRepairedName, True)

            '        'Will try to open Word file with the Open and Repair feature.
            '        'If successful there's no need to do more and it makes Word visible
            '        'even though it was started invisibly. If unsuccessful the program 
            '        'moves to the first catch section.

            '        objWord.Visible = False
            '        objDoc = objWord.Documents.OpenNoRepairDialog(originalFileFullPathRepairedName, OpenAndRepair:=True)
            '        objWord.Visible = True

            '        MsgBox("When opened, at least with open and repair function agreed to, this " _
            '               & "file does not appear to cause an error", MsgBoxStyle.Information)

            '        Me.Cursor = System.Windows.Forms.Cursors.Default
            '        ProgressBar1.Show()

            '    Catch ex As Exception

            '        'MessageBox.Show(ex.Message)
            '        'This section helps clean up the Word Objects which are fully released
            '        'in the finally sections with the garbage collectors. Without these
            '        'statements and the Garbage collectors, the hidden Word reference remain
            '        'activated and with repeated recovery attmpts, can build up, slowing down
            '        'your computer significantly.
            Try

                MsgBox("Two different configurations of DocToText for text salvaging of your " _
                    & "target file will be tried. The first will preserve basic things such " _
                    & "as paragraph breaks and tabs. The second will have no formatting " _
                    & "whatsoever. The processs should be fairly quick.", MsgBoxStyle.Information)

                File.Copy(sFileFullNameLC, zipRepairedsFileNameNonZipExtNoSpaceFullName, True)

                If File.Exists(salvagedsFileNameAndBasePathNoSpace2) = True Then

                    File.Delete(salvagedsFileNameAndBasePathNoSpace2)

                End If

                If File.Exists(salvagedsFileNameAndBasePathNoSpace) = True Then

                    File.Delete(salvagedsFileNameAndBasePathNoSpace)

                End If
                'MsgBox("Check that zipRepairedsFileNameNonZipExtNoSpaceFullName exists called: " & zipRepairedsFileNameNonZipExtNoSpaceFullName)
                DocToText_Process(zipRepairedsFileNameNonZipExtNoSpaceFullName, salvagedsFileNameNoSpace, sFileBasePath, True)

                Process.Start(salvagedsFileNameAndBasePathNoSpace)
                'ProgressBar1.Value= 50

                DocToText_Process(zipRepairedsFileNameNonZipExtNoSpaceFullName, salvagedsFileNameNoSpace2, sFileBasePath, False)
                Process.Start(salvagedsFileNameAndBasePathNoSpace2)
                Me.Cursor = System.Windows.Forms.Cursors.Default
                'ProgressBar1.Hide()

                If File.Exists(zipRepairedsFileNameNonZipExtNoSpaceFullName) Then

                    File.Delete(zipRepairedsFileNameNonZipExtNoSpaceFullName)

                End If

            Catch ex As Exception

                'MessageBox.Show(ex.Message)

            Finally

                Me.Cursor = System.Windows.Forms.Cursors.Default

                If File.Exists(sFile) Then

                    File.Delete(sFile)

                End If

                'If Not objDoc Is Nothing Then

                '    objDoc.Close(False)
                '    Marshal.FinalReleaseComObject(objDoc)
                '    objDoc = Nothing

                'End If

                'If Not objWord Is Nothing Then

                '    objWord.Quit(False)
                '    Marshal.FinalReleaseComObject(objWord)
                '    objWord = Nothing

                'End If

                ''Garbage collecting code along with the Marshal.FinalReleaseComObject 
                ''necessary to remove the instances of Word started by the program.

                'GC.Collect()
                'GC.WaitForPendingFinalizers()

                '' GC needs to be called twice in order to get the Finalizers called  
                '' - the first time in, it simply makes a list of what is to be  
                '' finalized, the second time in, it actually is finalizing. Only  
                '' then will the object do its automatic ReleaseComObject. 

                'GC.Collect()
                'GC.WaitForPendingFinalizers()

            End Try

        End If

    End Sub

    Private Sub Impl_S2RecoverText_From_DOC(sFileBasePathAndNameChangedtoDocExtension As String)

        Dim appWord As Word.Application = Nothing

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        Try

            MsgBox("This process might take a considerable amount of time, for instance if " _
                & "you have 100 pages it could take 2 minutes even on a new machine. Please " _
                & "be patient. Your text editor such as NotePad will probably open eventually " _
                & "however there might not be much indication of progress other than possibly " _
                & "Microsoft's ""spinning wheel"" or ""hourglass.""", MsgBoxStyle.Exclamation)

            Dim list As New List(Of String)
            Dim sFileInfo As New FileInfo(sFileBasePathAndNameChangedtoDocExtension)
            Dim sFileName As String = sFileInfo.Name
            Dim sFilePath As String = sFileInfo.Directory.ToString
            Dim recoveredTextFileName As String = sFilePath & "\recovered_text_" & sFileName & ".txt"
            Dim myStreamReaderL1 As System.IO.StreamReader
            Dim myStream As System.IO.StreamWriter
            Dim myStr As String = Nothing

            myStreamReaderL1 = System.IO.File.OpenText(sFileBasePathAndNameChangedtoDocExtension)
            myStr = myStreamReaderL1.ReadToEnd()
            myStreamReaderL1.Close()

            Dim chrIntegerCounter As Integer = 0

            While chrIntegerCounter < 9

                myStr = myStr.Replace(Chr(chrIntegerCounter), "")
                chrIntegerCounter = chrIntegerCounter + 1

            End While

            chrIntegerCounter = 11

            While chrIntegerCounter < 13

                myStr = myStr.Replace(Chr(chrIntegerCounter), "")
                chrIntegerCounter = chrIntegerCounter + 1

            End While

            chrIntegerCounter = 14

            While chrIntegerCounter < 32

                myStr = myStr.Replace(Chr(chrIntegerCounter), "")
                chrIntegerCounter = chrIntegerCounter + 1

            End While
            chrIntegerCounter = 127

            While chrIntegerCounter < 128

                myStr = myStr.Replace(Chr(chrIntegerCounter), "")
                chrIntegerCounter = chrIntegerCounter + 1

            End While

            chrIntegerCounter = 129

            While chrIntegerCounter < 153

                myStr = myStr.Replace(Chr(chrIntegerCounter), "")
                chrIntegerCounter = chrIntegerCounter + 1

            End While

            chrIntegerCounter = 154

            While chrIntegerCounter < 162

                myStr = myStr.Replace(Chr(chrIntegerCounter), "")
                chrIntegerCounter = chrIntegerCounter + 1

            End While

            chrIntegerCounter = 164

            While chrIntegerCounter < 165

                myStr = myStr.Replace(Chr(chrIntegerCounter), "")
                chrIntegerCounter = chrIntegerCounter + 1

            End While

            chrIntegerCounter = 166

            While chrIntegerCounter < 169

                myStr = myStr.Replace(Chr(chrIntegerCounter), "")
                chrIntegerCounter = chrIntegerCounter + 1

            End While

            chrIntegerCounter = 170

            While chrIntegerCounter < 174

                myStr = myStr.Replace(Chr(chrIntegerCounter), "")
                chrIntegerCounter = chrIntegerCounter + 1

            End While

            chrIntegerCounter = 175

            While chrIntegerCounter < 188

                myStr = myStr.Replace(Chr(chrIntegerCounter), "")
                chrIntegerCounter = chrIntegerCounter + 1

            End While

            myStr = myStr.Replace("�", "")
            myStr = myStr.Replace("ࡱ", "")
            myStr = myStr.Replace("  ", "")
            myStr = myStr.Replace("!""#$%&'()*+", "")
            myStr = myStr.Replace("Root EntryZ", "")
            myStr = myStr.Replace("CONTENTSRCompObjVSPELLING", "")
            'Save myStr
            recoveredTextFileName = sFilePath & "\recovered_text_" & sFileName & ".txt"
            myStream = System.IO.File.CreateText(recoveredTextFileName)
            myStream.WriteLine(myStr)
            myStream.Close()

            Using officeFileHandler As Process = New Process

                Process.Start(recoveredTextFileName)

            End Using

        Catch ex As Exception

            'MessageBox.Show(ex.Message)

        Finally

        End Try

    End Sub

    Function Zip_Repair(sFileZip As String, zipRepairedBasePathAndFileName As String)

        Dim testZipReader As StreamReader = Nothing
        Dim testZipErrorReader As StreamReader = Nothing
        Dim testZipCompOut As String = Nothing
        Dim testZipErrorOut As String = Nothing
        Dim repairZipReader As StreamReader = Nothing
        Dim repairZipErrorReader As StreamReader = Nothing
        Dim repairZipCompOut As String = Nothing
        Dim repairZipErrorOut As String = Nothing
        Dim testZipReader2 As StreamReader = Nothing
        Dim testZipErrorReader2 As StreamReader = Nothing
        Dim testZipCompOut2 As String = Nothing
        Dim testZipErrorOut2 As String = Nothing
        Dim repairZipReader2 As StreamReader = Nothing
        Dim repairZipErrorReader2 As StreamReader = Nothing
        Dim repairZipCompOut2 As String = Nothing
        Dim repairZipErrorOut2 As String = Nothing
        Dim testZipReader3 As StreamReader = Nothing
        Dim testZipErrorReader3 As StreamReader = Nothing
        Dim testZipCompOut3 As String = Nothing
        Dim testZipErrorOut3 As String = Nothing
        Dim sFileZipPrelimRepair As String = sFileZip & ".prelim_repair.zip"

        Dim testZipArguments As String = "-T -TT ""unzip -tqq"" """ & sFileZip & """"

        'MsgBox("testZipArguments  is: " & testZipArguments)

        Using testZip As New Process

            testZip.StartInfo.FileName = "zip.exe"
            testZip.StartInfo.Arguments = testZipArguments
            testZip.StartInfo.UseShellExecute = False
            testZip.StartInfo.CreateNoWindow = True
            testZip.StartInfo.RedirectStandardInput = True
            testZip.StartInfo.RedirectStandardOutput = True
            testZip.StartInfo.RedirectStandardError = True
            testZip.Start()
            testZip.StandardInput.WriteLine("y")
            testZipReader = testZip.StandardOutput
            testZipCompOut = testZipReader.ReadToEnd
            testZipErrorReader = testZip.StandardError
            testZipErrorOut = testZipErrorReader.ReadToEnd
            testZip.WaitForExit()
            testZip.Close()

        End Using

        'MsgBox("testZipCompOut.Contains is:" & testZipCompOut)

        If Not testZipCompOut.Contains("OK") Then

            Dim repairZipArguments As String = "-FF """ & sFileZip & """ --out """ _
            & sFileZipPrelimRepair & """"

            Using repairZip As New Process

                repairZip.StartInfo.FileName = "zip.exe"
                repairZip.StartInfo.Arguments = repairZipArguments
                repairZip.StartInfo.UseShellExecute = False
                repairZip.StartInfo.CreateNoWindow = True
                repairZip.StartInfo.RedirectStandardInput = True
                repairZip.StartInfo.RedirectStandardOutput = True
                repairZip.StartInfo.RedirectStandardError = True
                repairZip.Start()
                repairZip.StandardInput.WriteLine("y")
                repairZipReader = repairZip.StandardOutput
                repairZipCompOut = repairZipReader.ReadToEnd
                repairZipErrorReader = repairZip.StandardError
                repairZipErrorOut = repairZipErrorReader.ReadToEnd
                repairZip.WaitForExit()
                repairZip.Close()

            End Using

            'MsgBox("repairZipCompOut is: " & repairZipCompOut & _
            '" repairZipErrorOut is: " & repairZipErrorOut)

            If Not File.Exists(sFileZipPrelimRepair) Then

                File.Copy(sFileZip, sFileZipPrelimRepair, True)

            End If

            Dim testZipArguments2 As String = "-T -TT ""unzip -tqq"" """ & sFileZipPrelimRepair & """"

            Using testZip2 As New Process

                testZip2.StartInfo.FileName = "zip.exe"
                testZip2.StartInfo.Arguments = testZipArguments2
                testZip2.StartInfo.UseShellExecute = False
                testZip2.StartInfo.CreateNoWindow = True
                testZip2.StartInfo.RedirectStandardInput = True
                testZip2.StartInfo.RedirectStandardOutput = True
                testZip2.StartInfo.RedirectStandardError = True
                testZip2.Start()
                testZip2.StandardInput.WriteLine("y")
                testZipReader2 = testZip2.StandardOutput
                testZipCompOut2 = testZipReader2.ReadToEnd
                testZipErrorReader2 = testZip2.StandardError
                testZipErrorOut2 = testZipErrorReader2.ReadToEnd
                testZip2.WaitForExit()
                testZip2.Close()

            End Using

            'MsgBox("testZipCompOut2 is: " & testZipCompOut2 & _
            '" testZipErrorOut2 is: " & testZipErrorOut2)

            If Not testZipCompOut2.Contains("OK") Then

                Dim repairZipArguments2 As String = "-FF """ & sFileZipPrelimRepair & """ --out """ _
                    & zipRepairedBasePathAndFileName & """"

                Using repairZip2 As New Process

                    repairZip2.StartInfo.FileName = "zip.exe"
                    repairZip2.StartInfo.Arguments = repairZipArguments2
                    repairZip2.StartInfo.UseShellExecute = False
                    repairZip2.StartInfo.CreateNoWindow = True
                    repairZip2.StartInfo.RedirectStandardInput = True
                    repairZip2.StartInfo.RedirectStandardOutput = True
                    repairZip2.StartInfo.RedirectStandardError = True
                    repairZip2.Start()
                    repairZip2.StandardInput.WriteLine("y")
                    repairZipReader2 = repairZip2.StandardOutput
                    repairZipCompOut2 = repairZipReader2.ReadToEnd
                    repairZipErrorReader2 = repairZip2.StandardError
                    repairZipErrorOut2 = repairZipErrorReader2.ReadToEnd
                    repairZip2.WaitForExit()
                    repairZip2.Close()

                End Using

                'MsgBox("repairZipCompOut2 is: " & repairZipCompOut2 & _
                '" repairZipErrorOut2 is: " & repairZipErrorOut2)

                If Not File.Exists(zipRepairedBasePathAndFileName) Then

                    File.Copy(sFileZip, zipRepairedBasePathAndFileName, True)

                End If

                If Not testZipCompOut.Contains("no local entry: word/document.xml") Then

                    Zip_Repair = "OK"
                    Exit Function

                Else

                    Zip_Repair = "Failed"
                    Exit Function

                End If
                Zip_Repair = "OK"
            Else

                'MsgBox("testZipCompOut2.Contains OK")
                File.Copy(sFileZipPrelimRepair, zipRepairedBasePathAndFileName, True)
                Zip_Repair = "OK"
                Exit Function

            End If
            Zip_Repair = "OK"
        Else

            'MsgBox("testZipCompOut.Contains OK")
            File.Copy(sFileZip, zipRepairedBasePathAndFileName, True)
            Zip_Repair = "OK"
            Exit Function

        End If

        'MsgBox("Look for zip repaired version of the file: " & zipRepairedBasePathAndFileName)

    End Function

    Private Function Extract_Zip(zipRepairedBasePathAndFileName As String, extractedRepairedZipOutputDirectory As String)

        'I use 7zip to extract files as I found it better 
        'than other unzippers at extracting repaired zip files.

        Dim sevenZipExtractArguments As String = "x """ & zipRepairedBasePathAndFileName & """ -o""" & _
                                                extractedRepairedZipOutputDirectory & """"
        Dim extractZipReader As StreamReader = Nothing
        Dim extractZipErrorReader As StreamReader = Nothing
        Dim extractZipCompOut As String = Nothing
        Dim extractZipErrorOut As String = Nothing

        Using extractZip As Process = New Process

            extractZip.StartInfo.FileName = "7z.exe"
            extractZip.StartInfo.Arguments = sevenZipExtractArguments
            extractZip.StartInfo.UseShellExecute = False
            extractZip.StartInfo.CreateNoWindow = True
            extractZip.StartInfo.RedirectStandardInput = True
            extractZip.StartInfo.RedirectStandardOutput = True
            extractZip.StartInfo.RedirectStandardError = True
            extractZip.Start()
            extractZip.StandardInput.WriteLine("A")
            extractZipReader = extractZip.StandardOutput
            extractZipCompOut = extractZipReader.ReadToEnd
            extractZipErrorReader = extractZip.StandardError
            extractZipErrorOut = extractZipErrorReader.ReadToEnd
            extractZip.WaitForExit()
            extractZip.Close()

        End Using

        If Not extractZipCompOut.Contains("Error: Can not open file as archive") Then

            Return "OK"

            Exit Function

        Else

            Return "Failed"

            Exit Function

        End If

    End Function

    Private Function Make_Document_XML_Pretty_Print(extractedRepairedZipOutputDirectory As String)

        'All this code does is take every instance of ><, e.g. the end of one tag and the beginning the next
        'and puts a carriage return \r, character which coded for by  Chr(13), in between the two angle
        'brackets. This is enough to make each tag in the whole document appear on a separate line.

        Dim corruptDocumentXMLPath As String = extractedRepairedZipOutputDirectory & "\word\document.xml"

        If File.Exists(corruptDocumentXMLPath) Then

            My.Computer.FileSystem.WriteAllText(corruptDocumentXMLPath, _
            My.Computer.FileSystem.ReadAllText(corruptDocumentXMLPath).Replace(Chr(62).ToString() & Chr(60).ToString(), _
                    Chr(62).ToString() & Chr(13).ToString() & Chr(10).ToString() & Chr(60).ToString()), False)

            Make_Document_XML_Pretty_Print = "OK"

            'MsgBox("Check if document has been pretty printed.")

        Else

            MsgBox("Your DOCX file does not contain the word/document.xml subfile where " _
                   & "all the text is stored. It still may be possible to recover content  " _
                   & "such footnotes using the SilverCoder's command line program DocToText. " _
                   & "Savvy DOCX will now engage that program for you.", MsgBoxStyle.Exclamation)

            Make_Document_XML_Pretty_Print = "Failed"

        End If

    End Function

    Function Re_Zip(extractedRepairedZipOutputDirectory As String, zipExtensionXMLRepairedFullPath As String)

        'Seven zip command line works well in rezipping too.

        Dim sevenZipReZipReader As StreamReader = Nothing
        Dim sevenZipReZipErrorReader As StreamReader = Nothing
        Dim sevenZipReZipCompOut As String = Nothing
        Dim sevenZipReZipErrorOut As String = Nothing

        Dim sevenZipUpArguments As String = "a -r """ & zipExtensionXMLRepairedFullPath & """ """ & _
               extractedRepairedZipOutputDirectory & """\*"

        Using sevenZipReZip As Process = New Process

            sevenZipReZip.StartInfo.FileName = "7z.exe"
            sevenZipReZip.StartInfo.Arguments = sevenZipUpArguments
            sevenZipReZip.StartInfo.UseShellExecute = False
            sevenZipReZip.StartInfo.CreateNoWindow = True
            sevenZipReZip.StartInfo.RedirectStandardOutput = True
            sevenZipReZip.StartInfo.RedirectStandardError = True
            sevenZipReZip.Start()
            sevenZipReZipReader = sevenZipReZip.StandardOutput
            sevenZipReZipCompOut = sevenZipReZipReader.ReadToEnd
            sevenZipReZipErrorReader = sevenZipReZip.StandardError
            sevenZipReZipErrorOut = sevenZipReZipErrorReader.ReadToEnd
            sevenZipReZip.WaitForExit()
            sevenZipReZip.Close()

        End Using

        'MsgBox("sevenZipReZipCompOut is: " & sevenZipReZipCompOut)

        If Not sevenZipReZipCompOut.Contains("Everything is Ok") Then

            Re_Zip = "Failed"
            Exit Function

        Else

            Re_Zip = "OK"
            Exit Function

        End If

    End Function

    Private Function XML_Validate(corruptFileDirRetrievedXMLFileFullPath As String)

        'This function validates an XML based file and reports the byte location of 
        'any xml well-formedness errors. It uses a executable compiled Perl script

        Dim xmlValidateArguments As String = """" & corruptFileDirRetrievedXMLFileFullPath & """"
        Dim xmlValidateReader As StreamReader = Nothing
        Dim xmlValidateErrorReader As StreamReader = Nothing
        Dim xmlValidateCompOut As String = Nothing
        Dim xmlValidateErrorOut As String = Nothing

        Using xmlValidate As Process = New Process

            xmlValidate.StartInfo.FileName = "xmlval.exe"
            xmlValidate.StartInfo.Arguments = xmlValidateArguments
            xmlValidate.StartInfo.UseShellExecute = False
            xmlValidate.StartInfo.RedirectStandardOutput = True
            xmlValidate.StartInfo.RedirectStandardError = True
            xmlValidate.StartInfo.CreateNoWindow = True
            xmlValidate.Start()
            xmlValidateReader = xmlValidate.StandardOutput
            xmlValidateCompOut = xmlValidateReader.ReadToEnd
            xmlValidateErrorReader = xmlValidate.StandardError
            xmlValidateErrorOut = xmlValidateErrorReader.ReadToEnd
            xmlValidate.WaitForExit()
            xmlValidate.Close()

        End Using

        Return xmlValidateErrorOut

    End Function

    Private Function Trid_Identify(sFileFullName As String)

        Dim tridDocTypeIdentifierReader As StreamReader = Nothing
        Dim tridDocTypeIdentifierErrorReader As StreamReader = Nothing
        Dim tridDocTypeIdentifierCompOut As String = Nothing
        Dim tridDocTypeIdentifierErrorOut As String = Nothing

        Using tridDocTypeIdentifier As Process = New Process

            tridDocTypeIdentifier.StartInfo.FileName = "trid.exe"
            tridDocTypeIdentifier.StartInfo.Arguments = """" & sFileFullName & """"
            tridDocTypeIdentifier.StartInfo.RedirectStandardOutput = True
            tridDocTypeIdentifier.StartInfo.RedirectStandardError = True
            tridDocTypeIdentifier.StartInfo.UseShellExecute = False
            tridDocTypeIdentifier.StartInfo.CreateNoWindow = True
            tridDocTypeIdentifier.Start()
            tridDocTypeIdentifierReader = tridDocTypeIdentifier.StandardOutput
            tridDocTypeIdentifierCompOut = tridDocTypeIdentifierReader.ReadToEnd
            tridDocTypeIdentifierErrorReader = tridDocTypeIdentifier.StandardError
            tridDocTypeIdentifierErrorOut = tridDocTypeIdentifierErrorReader.ReadToEnd
            tridDocTypeIdentifier.WaitForExit()
            tridDocTypeIdentifier.Close()

        End Using

        Return tridDocTypeIdentifierCompOut

    End Function

    Private Sub Truncate_XML(corruptFileDirRetrievedXMLFileFullPath As String, truncatedLengthAsString As String)

        'This code just uses a command line file truncator I found on Google code called "trunc".

        Dim truncateArguments As String = """" & corruptFileDirRetrievedXMLFileFullPath & """ " & truncatedLengthAsString

        Using truncate As Process = New Process

            truncate.StartInfo.FileName = "trunc.exe"
            truncate.StartInfo.Arguments = truncateArguments
            truncate.StartInfo.UseShellExecute = False
            truncate.StartInfo.CreateNoWindow = True
            truncate.Start()
            truncate.WaitForExit()
            truncate.Close()

        End Using

    End Sub

    Private Sub DocToText_Process(zipRepairedsFileNameNonZipExtNoSpace As String, salvagedsFileNameNoSpace As String, sFileBasePath As String, UseFix As Boolean)

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        'I'm not using this doctotext.exe invoking function at the moment.

        Dim fixOrStripValue As String = Nothing

        If UseFix = True Then

            fixOrStripValue = "--ooxml --fix-xml"

        Else

            fixOrStripValue = "--ooxml --strip-xml"

        End If

        Dim sb As New StringBuilder(256)
        GetShortPathName(zipRepairedsFileNameNonZipExtNoSpace, sb, sb.Capacity)

        Dim shortName As String = sb.ToString()
        'Dim gwdOf7z As String = Directory.GetCurrentDirectory & "\7z.exe"
        'MsgBox("zipRepairedsFileNameNonZipExtNoSpace is: " & zipRepairedsFileNameNonZipExtNoSpace _
        ' & "shortName is: " & shortName)
        Dim salvagedsFileNameNoSpaceFullPath As String = sFileBasePath & "\" & salvagedsFileNameNoSpace
        Dim doctotextReader As StreamReader = Nothing
        Dim doctotextErrorReader As StreamReader = Nothing
        Dim doctotextCompOut As String = Nothing
        Dim doctotextErrorOut As String = Nothing
        Dim doctotextArguments As String = fixOrStripValue & _
            " --unzip-cmd=""7z.exe x %a %f -o%d"" " & shortName
        'MsgBox("doctotextArguments is: " & "doctotext.exe " & doctotextArguments)
        Using doctotextProcess As Process = New Process

            doctotextProcess.StartInfo.FileName = "doctotext.exe"
            doctotextProcess.StartInfo.Arguments = doctotextArguments
            doctotextProcess.StartInfo.UseShellExecute = False
            doctotextProcess.StartInfo.CreateNoWindow = True
            doctotextProcess.StartInfo.RedirectStandardOutput = True
            doctotextProcess.StartInfo.RedirectStandardError = True
            doctotextProcess.Start()
            doctotextReader = doctotextProcess.StandardOutput
            doctotextCompOut = doctotextReader.ReadToEnd
            doctotextErrorReader = doctotextProcess.StandardError
            doctotextErrorOut = doctotextErrorReader.ReadToEnd
            doctotextProcess.WaitForExit()
            doctotextProcess.Close()

        End Using

        My.Computer.FileSystem.WriteAllText(salvagedsFileNameNoSpaceFullPath, doctotextCompOut, False, System.Text.Encoding.Default)
        My.Computer.FileSystem.WriteAllText(sFileBasePath & "\savvy_docx_recovery_text_extraction_error.txt", doctotextErrorOut, False)

    End Sub

    Private Sub Unspecified_Error_Document_XmlTreatment(ex2 As Exception, corruptFileXMLSubFileFullName As String, modDocumentXmlFullPath As String)

        'Here is the code where the document.xml file from an unspecified error
        'word file is treated and the bad section between m:oMath, w:tbl
        'or w:fldChar-w:fldCharType tags is removed. I'm not actually using
        'it at the moment. I put it back in thread so I could do a "Continue For"
        'for any file that didn't have normal kinds of unspecified error corruptions.

        Dim lineList As New List(Of String)
        Dim secondLineColumnErrorLocationInteger As Integer = 0
        Dim indexBadTagOpenLine As Integer = 0
        Dim indexBadTagCloseLine As Integer = 0
        Dim numberOfCharactersInBadTags As Integer = 0
        Dim secondLineColumnErrorLocationString As String = Nothing
        Dim excisedString As String = Nothing
        Dim byteMatch As Match = Regex.Match(ex2.Message, ("Line: [0-9]+"))
        Dim byteMatchString As String = byteMatch.ToString

        secondLineColumnErrorLocationString = DelFromLeft("Line: ", byteMatchString)
        Integer.TryParse(secondLineColumnErrorLocationString, secondLineColumnErrorLocationInteger)

        Using sr As StreamReader = New StreamReader(corruptFileXMLSubFileFullName)

            Do While sr.Peek() >= 0
                lineList.Add(sr.ReadLine())
            Loop

            sr.Close()

        End Using

        indexBadTagOpenLine = lineList.LastIndexOf("<m:oMath>", secondLineColumnErrorLocationInteger) + 0
        indexBadTagCloseLine = lineList.IndexOf("</m:oMath>", secondLineColumnErrorLocationInteger) + 1

        If indexBadTagOpenLine = -1 Then

            indexBadTagOpenLine = lineList.LastIndexOf("<w:tbl>", secondLineColumnErrorLocationInteger) + 0

        End If

        If indexBadTagCloseLine = 0 Then

            indexBadTagCloseLine = lineList.IndexOf("</w:tbl>", secondLineColumnErrorLocationInteger) + 1

        End If

        If indexBadTagOpenLine = -1 Then

            indexBadTagOpenLine = lineList.LastIndexOf("<w:fldChar w:fldCharType=""begin""/>", secondLineColumnErrorLocationInteger) + 0

        End If

        If indexBadTagCloseLine = 0 Then

            indexBadTagCloseLine = lineList.IndexOf("<w:fldChar w:fldCharType=""end""/>", secondLineColumnErrorLocationInteger) + 1

        End If

        If indexBadTagOpenLine = -1 Then

            MsgBox("Your unspecified error appears to be of an unusual type. This " _
                   & "program is reverting to another form of recovery that uses " _
                   & "the command line program xmllint", MsgBoxStyle.Exclamation)

            Xmllint_Treatment(modDocumentXmlFullPath)

        End If

        'MsgBox(indexBadTagOpenLine.ToString)

        RichTextBox1.Clear()
        numberOfCharactersInBadTags = indexBadTagCloseLine - indexBadTagOpenLine

        For Each excisedString In lineList.GetRange(indexBadTagOpenLine, numberOfCharactersInBadTags)

            RichTextBox1.AppendText(excisedString)

        Next

        lineList.RemoveRange(indexBadTagOpenLine, numberOfCharactersInBadTags)
        File.WriteAllLines(modDocumentXmlFullPath, lineList.ToArray)

    End Sub

    Private Sub Process_XML_Files_With_Xmllint_First(corruptFileXMLSubFileFullName As String, modDocumentXmlFullPath As String)

        Dim sFile As String = TextBox1.Text
        Dim xmlValidateReader As StreamReader = Nothing
        Dim xmlValidateErrorReader As StreamReader = Nothing
        Dim xmlValidateReader2 As StreamReader = Nothing
        Dim xmlValidateErrorReader2 As StreamReader = Nothing
        Dim xmlValidateReader3 As StreamReader = Nothing
        Dim xmlValidateErrorReader3 As StreamReader = Nothing
        Dim byteMatch As Match = Nothing
        Dim sFileInfo As FileInfo = New FileInfo(sFile)
        Dim byteErrorLocationInteger As Integer = 0
        Dim truncatedLength As Integer = 0
        Dim intTruncationAmount As Integer = 0
        Dim xmlValidateCompOut As String = Nothing
        Dim xmlValidateErrorOut As String = Nothing
        Dim xmlValidateArguments As String = Nothing
        Dim byteMatchString As String = Nothing
        Dim truncatedLengthAsString As String = Nothing
        Dim truncateArguments As String = Nothing
        Dim xmlValidateCompOut2 As String = Nothing
        Dim xmlValidateErrorOut2 As String = Nothing
        Dim xmlValidateCompOut3 As String = Nothing
        Dim xmlValidateErrorOut3 As String = Nothing
        Dim xmlRecoverArguments As String = Nothing
        Dim byteErrorLocation As String = Nothing
        Dim sFileName As String = LCase(sFileInfo.Name)
        'Dim sFileBasePath As String = LCase(sFileInfo.Directory.ToString)
        Dim sFileBasePath As String = tempPath

        xmlValidateArguments = """" & corruptFileXMLSubFileFullName & """"

        Using xmlValidate As Process = New Process

            xmlValidate.StartInfo.FileName = "xmlval.exe"
            xmlValidate.StartInfo.Arguments = xmlValidateArguments
            xmlValidate.StartInfo.UseShellExecute = False
            xmlValidate.StartInfo.RedirectStandardOutput = True
            xmlValidate.StartInfo.RedirectStandardError = True
            xmlValidate.StartInfo.CreateNoWindow = True
            xmlValidate.Start()
            xmlValidateReader = xmlValidate.StandardOutput
            xmlValidateCompOut = xmlValidateReader.ReadToEnd
            xmlValidateErrorReader = xmlValidate.StandardError
            xmlValidateErrorOut = xmlValidateErrorReader.ReadToEnd
            xmlValidate.WaitForExit()
            xmlValidate.Close()

        End Using

        'If the validate process returns an error with the 
        'word byte in it then the xml file needs to be processed.

        If xmlValidateErrorOut.Contains("byte") Then

            'We first try to process it with xmllint  
            'alone, which is what this subroutine does.

            Xmllint_Treatment(modDocumentXmlFullPath)
            'File.Copy(corruptFileXMLSubFileFullName, modDocumentXmlFullPath, True)

            Dim xmlValidateArguments2 = """" & modDocumentXmlFullPath & """"

            Using xmlValidate2 As Process = New Process

                xmlValidate2.StartInfo.FileName = "xmlval.exe"
                xmlValidate2.StartInfo.Arguments = xmlValidateArguments2
                xmlValidate2.StartInfo.UseShellExecute = False
                xmlValidate2.StartInfo.RedirectStandardOutput = True
                xmlValidate2.StartInfo.RedirectStandardError = True
                xmlValidate2.StartInfo.CreateNoWindow = True
                xmlValidate2.Start()
                xmlValidateReader2 = xmlValidate2.StandardOutput
                xmlValidateCompOut2 = xmlValidateReader2.ReadToEnd
                xmlValidateErrorReader2 = xmlValidate2.StandardError
                xmlValidateErrorOut2 = xmlValidateErrorReader2.ReadToEnd
                xmlValidate2.WaitForExit()
                xmlValidate2.Close()

            End Using

            'If xmllint alone didn't fix it, we try 
            'the xml truncating at first error treatment.

            If xmlValidateErrorOut2.Contains("byte") Then

                Dim loopCounter As Integer = 0

                Do

                    'The validator will register an error and indicate the byte location of 
                    'the error if the document.xml file has an error. We isolate this byte 
                    'location with a Regex and the DelFromleft function, changed the byte to 
                    'an integer and subtract first 0 then 50 then 100 bytes to try to steer 
                    'clear of any additional bad xml if there is some just before the error.

                    loopCounter = loopCounter + 1

                    If loopCounter = 1 Then

                        intTruncationAmount = 50

                    ElseIf loopCounter = 2 Then

                        intTruncationAmount = 100

                    Else

                        intTruncationAmount = 150

                    End If

                    byteMatch = Regex.Match(xmlValidateErrorOut2, _
                     "byte [0-9]+")
                    byteMatchString = byteMatch.ToString
                    byteErrorLocation = DelFromLeft("byte ", byteMatchString)
                    Integer.TryParse(byteErrorLocation, byteErrorLocationInteger)

                    truncatedLength = byteErrorLocationInteger - intTruncationAmount
                    truncatedLengthAsString = String.Empty
                    truncatedLengthAsString = System.Convert.ToString(truncatedLength)
                    truncateArguments = """" & modDocumentXmlFullPath & """ " & truncatedLengthAsString

                    'Now we will truncate the file at bad byte minus 0, 50 or 100 bytes.

                    Using truncate As Process = New Process

                        truncate.StartInfo.FileName = "trunc.exe"
                        truncate.StartInfo.Arguments = truncateArguments
                        truncate.StartInfo.UseShellExecute = False
                        truncate.StartInfo.CreateNoWindow = True
                        truncate.Start()
                        truncate.WaitForExit()
                        truncate.Close()

                    End Using

                    xmlRecoverArguments = "--recover " & _
                        """" & modDocumentXmlFullPath & """" & " -o " & _
                        """" & modDocumentXmlFullPath & """"

                    'Now we use xmllint to reconstruct the nice xml ending tags
                    'to try to slip document.xml past Word's XML validator.

                    Using xmllintFixXMLEndTags As Process = New Process

                        xmllintFixXMLEndTags.StartInfo.FileName = "xmllint.exe"
                        xmllintFixXMLEndTags.StartInfo.Arguments = xmlRecoverArguments
                        xmllintFixXMLEndTags.StartInfo.UseShellExecute = False
                        xmllintFixXMLEndTags.StartInfo.CreateNoWindow = True
                        xmllintFixXMLEndTags.Start()
                        xmllintFixXMLEndTags.WaitForExit()
                        xmllintFixXMLEndTags.Close()

                    End Using

                    'We validate again, hoping our xmlRecovery is OK by our validator.

                    Using xmlValidate3 As Process = New Process

                        xmlValidate3.StartInfo.FileName = "xmlval.exe"
                        xmlValidate3.StartInfo.Arguments = xmlValidateArguments2
                        xmlValidate3.StartInfo.UseShellExecute = False
                        xmlValidate3.StartInfo.RedirectStandardOutput = True
                        xmlValidate3.StartInfo.RedirectStandardError = True
                        xmlValidate3.StartInfo.CreateNoWindow = True
                        xmlValidate3.Start()
                        xmlValidateReader3 = xmlValidate3.StandardOutput
                        xmlValidateCompOut3 = xmlValidateReader3.ReadToEnd
                        xmlValidateErrorReader3 = xmlValidate3.StandardError
                        xmlValidateErrorOut3 = xmlValidateErrorReader3.ReadToEnd
                        xmlValidate3.WaitForExit()
                        xmlValidate3.Close()

                    End Using

                    'If our validator says the XML is still bad, we 
                    'go through a 2nd of truncation and XML recovery

                    If loopCounter = 4 Then

                        Exit Do

                    End If

                    'We end the loops if the validator 
                    'error does not return anything.

                Loop Until Not xmlValidateErrorOut3.Contains("byte")

            End If

        End If

    End Sub

    Private Sub Process_XML_Files_With_Truncator_First(modDocumentXmlFullPath As String)

        Dim sFile As String = TextBox1.Text
        Dim xmlValidateReader2 As StreamReader = Nothing
        Dim xmlValidateErrorReader2 As StreamReader = Nothing
        Dim xmlValidateReader3 As StreamReader = Nothing
        Dim xmlValidateErrorReader3 As StreamReader = Nothing
        Dim byteMatch As Match = Nothing
        Dim sFileInfo As FileInfo = New FileInfo(sFile)
        Dim byteErrorLocationInteger As Integer = 0
        Dim truncatedLength As Integer = 0
        Dim intTruncationAmount As Integer = 0
        Dim byteMatchString As String = Nothing
        Dim truncatedLengthAsString As String = Nothing
        Dim truncateArguments As String = Nothing
        Dim xmlValidateCompOut2 As String = Nothing
        Dim xmlValidateErrorOut2 As String = Nothing
        Dim xmlRecoverArguments As String = Nothing
        Dim byteErrorLocation As String = Nothing
        Dim xmlValidateCompOut3 As String = Nothing
        Dim xmlValidateErrorOut3 As String = Nothing
        Dim sFileName As String = LCase(sFileInfo.Name)
        'Dim sFileBasePath As String = LCase(sFileInfo.Directory.ToString)
        Dim sFileBasePath As String = tempPath
        Dim xmlValidateArguments2 = """" & modDocumentXmlFullPath & """"

        Using xmlValidate2 As Process = New Process

            xmlValidate2.StartInfo.FileName = "xmlval.exe"
            xmlValidate2.StartInfo.Arguments = xmlValidateArguments2
            xmlValidate2.StartInfo.UseShellExecute = False
            xmlValidate2.StartInfo.RedirectStandardOutput = True
            xmlValidate2.StartInfo.RedirectStandardError = True
            xmlValidate2.StartInfo.CreateNoWindow = True
            xmlValidate2.Start()
            xmlValidateReader2 = xmlValidate2.StandardOutput
            xmlValidateCompOut2 = xmlValidateReader2.ReadToEnd
            xmlValidateErrorReader2 = xmlValidate2.StandardError
            xmlValidateErrorOut2 = xmlValidateErrorReader2.ReadToEnd
            xmlValidate2.WaitForExit()
            xmlValidate2.Close()

        End Using

        'With this subroutine, we go straight into code
        'that truncates at the first XML well-formedness error.

        If xmlValidateErrorOut2.Contains("byte") Then

            Dim loopCounter As Integer = 0

            Do

                'The validator will register an error and indicate the byte location of 
                'the error if the document.xml file has an error. We isolate this byte 
                'location with a Regex and the DelFromleft function, changed the byte to 
                'an integer and subtract first 0 then 50 then 100 bytes to try to steer 
                'clear of any additional bad xml if there is some just before the error.

                loopCounter = loopCounter + 1

                If loopCounter = 1 Then

                    intTruncationAmount = 150

                ElseIf loopCounter = 2 Then

                    intTruncationAmount = 100

                Else

                    intTruncationAmount = 150

                End If

                byteMatch = Regex.Match(xmlValidateErrorOut2, _
                 "byte [0-9]+")
                byteMatchString = byteMatch.ToString
                byteErrorLocation = DelFromLeft("byte ", byteMatchString)
                Integer.TryParse(byteErrorLocation, byteErrorLocationInteger)

                truncatedLength = byteErrorLocationInteger - intTruncationAmount
                truncatedLengthAsString = String.Empty
                truncatedLengthAsString = System.Convert.ToString(truncatedLength)
                truncateArguments = """" & modDocumentXmlFullPath & """ " & truncatedLengthAsString

                'Now we will truncate the file at bad byte minus 0, 50 or 100 bytes.

                Using truncate As Process = New Process

                    truncate.StartInfo.FileName = "trunc.exe"
                    truncate.StartInfo.Arguments = truncateArguments
                    truncate.StartInfo.UseShellExecute = False
                    truncate.StartInfo.CreateNoWindow = True
                    truncate.Start()
                    truncate.WaitForExit()
                    truncate.Close()

                End Using

                xmlRecoverArguments = "--recover " & _
                    """" & modDocumentXmlFullPath & """" & " -o " & _
                    """" & modDocumentXmlFullPath & """"

                'Now we use xmllint to reconstruct the nice xml ending tags
                'to try to slip document.xml past Word's XML validator.

                Using xmllintFixXMLEndTags As Process = New Process

                    xmllintFixXMLEndTags.StartInfo.FileName = "xmllint.exe"
                    xmllintFixXMLEndTags.StartInfo.Arguments = xmlRecoverArguments
                    xmllintFixXMLEndTags.StartInfo.UseShellExecute = False
                    xmllintFixXMLEndTags.StartInfo.CreateNoWindow = True
                    xmllintFixXMLEndTags.Start()
                    xmllintFixXMLEndTags.WaitForExit()
                    xmllintFixXMLEndTags.Close()

                End Using

                'We validate again, hoping our xmlRecovery is OK by our validator.


                Using xmlValidate3 As Process = New Process

                    xmlValidate3.StartInfo.FileName = "xmlval.exe"
                    xmlValidate3.StartInfo.Arguments = xmlValidateArguments2
                    xmlValidate3.StartInfo.UseShellExecute = False
                    xmlValidate3.StartInfo.RedirectStandardOutput = True
                    xmlValidate3.StartInfo.RedirectStandardError = True
                    xmlValidate3.StartInfo.CreateNoWindow = True
                    xmlValidate3.Start()
                    xmlValidateReader3 = xmlValidate3.StandardOutput
                    xmlValidateCompOut3 = xmlValidateReader3.ReadToEnd
                    xmlValidateErrorReader3 = xmlValidate3.StandardError
                    xmlValidateErrorOut3 = xmlValidateErrorReader3.ReadToEnd
                    xmlValidate3.WaitForExit()
                    xmlValidate3.Close()

                End Using

                'If our validator says the XML is still bad, we 
                'go through a 2nd of truncation and XML recovery

                If loopCounter = 4 Then

                    Exit Do

                End If

            Loop Until Not xmlValidateErrorOut3.Contains("byte")

        End If

    End Sub

    Private Sub Xmllint_Treatment(xmllintRecoverXmlTreatedFile As String)

        'This sub is for xmllint only.

        Dim xmllintArguments As String = "--recover """ & xmllintRecoverXmlTreatedFile _
                                         & """ -o """ & xmllintRecoverXmlTreatedFile & """"

        Using xmlProcess As Process = New Process

            xmlProcess.StartInfo.FileName = "xmllint.exe"
            xmlProcess.StartInfo.Arguments = xmllintArguments
            xmlProcess.StartInfo.UseShellExecute = False
            xmlProcess.StartInfo.CreateNoWindow = True
            xmlProcess.Start()
            xmlProcess.WaitForExit()
            xmlProcess.Close()

        End Using

    End Sub

    Sub Launch_URL(URL As String)

        Try

            Dim proc As New Process()
            proc.StartInfo.FileName = URL
            proc.StartInfo.Arguments = ""
            proc.StartInfo.UseShellExecute = True
            proc.Start()

        Catch ex As Exception

            'MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

        MsgBox("If you upload files on my form and pay the $22, I will analyze and repair your file " _
               & "if it is possible to repair. If it proves impossible, or unsatisfactory, I refund " _
               & "$17 of the fee. Turnaround time is generally 2-5 hours.")

        Launch_URL("http://saveofficedata.com/contact.htm")

    End Sub

    Private Sub picWordFixLink_Click(sender As Object, e As EventArgs) Handles picWordFixLink.Click

        MsgBox("If you upload files on my form and pay the $22, I will analyze and repair your file " _
               & "if it is possible to repair. If it proves impossible, or unsatisfactory, I refund " _
               & "$17 of the fee. Turnaround time is generally 2-5 hours.")

        Launch_URL("http://saveofficedata.com/contact.htm")


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If firstTimeThrough = False Then

            MsgBox("To keep Word document corruption to a minimum, it is important to stress the need to " _
       & "install your Microsoft Office version's most up to date service packs and hotfixes. " _
       & "If you haven't already done so, please use this program's buttons to do it. The " _
       & "Mr. Fixit link, is really a little program that fixes Math related end tag out " _
       & "of order issues that this present program should also fix.", MsgBoxStyle.Information)

            firstTimeThrough = True

        Else


        End If


        TextBox1.Text = Nothing
        SelectFile()

        If String.IsNullOrEmpty(TextBox1.Text) = True Then

            SelectFile()

        Else



        End If

    End Sub

    Private Sub PictureBox1_Click_1(sender As Object, e As EventArgs) Handles PictureBox1.Click

        MsgBox("Installing the Service Pack 3 for Office 2007 may fix the unspecified error " _
               & "connected to the table of contents markup and the w:fldChar tags. Of course " _
               & "it is always good practice to install the latest service packs. When writing " _
               & "this program on Oct. 19, 2013, Microsoft Office 2007 is up to Service Pack 3, " _
               & "Office 2010 is up to Service Pack 2 and Office 2013 has not had a service pack " _
               & "released yet. Service packs are cumulative, meaning if you have Office 2007, you " _
               & "can skip from SP1 to SP3 as SP2 (and SP1 for that matter) are included in SP3.", _
               MsgBoxStyle.Information)

        Launch_URL("http://www.microsoft.com/en-us/download/office-service-packs.aspx")

    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

        MsgBox("Most Word DOCX documents of the ""Unspecified Error line " _
        & "2, Column 0"" variety, will open in Zoho writer. From " _
        & "there they can be exported to Word DOC or Open Office ODT format.", _
        MsgBoxStyle.Information)

        Launch_URL("https://www.zoho.com/")

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click

        MsgBox("Most Word DOCX documents of the ""Unspecified Error line " _
        & "2, Column 0"" variety, will open in Zoho writer. From " _
        & "there they can be exported to Word DOC or Open Office ODT format.", _
        MsgBoxStyle.Information)

        Launch_URL("https://www.zoho.com/")

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

        MsgBox("This button opens links to two Microsoft community threads where MVPs and " _
               & "other Word experts will manually fix your ""Unspecified"" error Word files " _
               & "for free. There are others.", MsgBoxStyle.Information)

        Launch_URL("http://answers.microsoft.com/en-us/office/forum/office_2010-word/unspecified-error-worddocumentxml-line2-column-0/21971fa0-df44-4ba6-ac42-7d4b5cd4174f?page=23&tab=question&status=AllReplies&tm=1380920402797")

        Launch_URL("http://social.technet.microsoft.com/Forums/office/en-US/11c6f40d-e229-428a-9d76-cf0e5be3b8c8/corrupted-word-2007-file-unspecified-error?forum=officesetupdeploylegacy")

    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

        MsgBox("Installing the Service Pack 3 for Office 2007 may fix the unspecified error " _
               & "connected to the table of contents markup and the w:fldChar tags. Of course " _
               & "it is always good practice to install the latest service packs. When writing " _
               & "this program on Oct. 19, 2013, Microsoft Office 2007 is up to Service Pack 3, " _
               & "Office 2010 is up to Service Pack 2 and Office 2013 has not had a service pack " _
               & "released yet. Service packs are cumulative, meaning if you have Office 2007, you " _
               & "can skip from SP1 to SP3 as SP2 (and SP1 for that matter) are included in SP3.", _
               MsgBoxStyle.Information)

        Launch_URL("http://www.microsoft.com/en-us/download/office-service-packs.aspx")

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

        MsgBox("This button opens links to two Microsoft community threads where MVPs and " _
               & "other Word experts will manually fix your ""Unspecified"" error Word files " _
               & "for free. There are others.", MsgBoxStyle.Information)

        Launch_URL("http://answers.microsoft.com/en-us/office/forum/office_2010-word/unspecified-error-worddocumentxml-line2-column-0/21971fa0-df44-4ba6-ac42-7d4b5cd4174f?page=23&tab=question&status=AllReplies&tm=1380920402797")

        Launch_URL("http://social.technet.microsoft.com/Forums/office/en-US/11c6f40d-e229-428a-9d76-cf0e5be3b8c8/corrupted-word-2007-file-unspecified-error?forum=officesetupdeploylegacy")

    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

        MsgBox("This hotfix fix for Office and Word 2010 fixes the ""Unspecified error"" " _
               & "issue for future documents. The hotfix is not contained in Service Pack " _
               & "1 or 2 of Office 2010 and it requires either service pack to be installed " _
               & "in order to work.", MsgBoxStyle.Information)

        Launch_URL("http://support.microsoft.com/kb/2817583")

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click

        MsgBox("This hotfix fix for Office and Word 2010 fixes the ""Unspecified error"" " _
        & "issue for future documents. The hotfix is not contained in Service Pack " _
        & "1 or 2 of Office 2010 and it requires either service pack to be installed " _
        & "in order to work.", MsgBoxStyle.Information)

        Launch_URL("http://support.microsoft.com/kb/2817583")

    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

        MsgBox("This hotfix for Office and Word 2007, repairs the ""Unspecified error"" " _
        & "issue for future documents. With Word 2007 Unspecified Errors, the issue " _
        & "occurs after the installation of security patch MS08-072 and appears to be  " _
        & "different than Word 2010 unspecified errors in that a column number " _
        & "is given whereas with Word 2010 the column number is 0. The hotfix is " _
        & "not contained in Service Pack 1 or 2 of Office 2007 and it requires " _
        & "either service pack to be installed in order to work.", MsgBoxStyle.Information)

        Launch_URL("http://support.microsoft.com/kb/970942")

    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click

        MsgBox("This hotfix for Office and Word 2007, repairs the ""Unspecified error"" " _
        & "issue for future documents. With Word 2007 Unspecified Errors, the issue " _
        & "occurs after the installation of security patch MS08-072 and appears to be  " _
        & "different than Word 2010 unspecified errors in that a column number " _
        & "is given whereas with Word 2010 the column number is 0. The hotfix is " _
        & "not contained in Service Pack 1 or 2 of Office 2007 and it requires " _
        & "either service pack to be installed in order to work.", MsgBoxStyle.Information)

        Launch_URL("http://support.microsoft.com/kb/970942")

    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click

        MsgBox("This Mr Fixit applet will directly fix Word 2007-2013 files having the error: " _
        & """The name in the end tag of the element must match the element type in the " _
        & "start tag."" After applying the applet the problem is not permanently fixed and " _
        & "can reoccur. The issue is fixed in Office 2013 and Office 2010 Service Pack 1. " _
        & """This issue is related strictly to oMath tags and occurs when a graphical " _
        & "object or text box is anchored to the same paragraph that contains the equation.""", _
        MsgBoxStyle.Information)

        Launch_URL("http://support.microsoft.com/kb/2528942")

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

        MsgBox("This Mr Fixit applet will directly fix Word 2007-2013 files having the error: " _
        & """The name in the end tag of the element must match the element type in the " _
        & "start tag."" After applying the applet the problem is not permanently fixed and " _
        & "can reoccur. The issue is fixed in Office 2013 and Office 2010 Service Pack 1. " _
        & """This issue is related strictly to oMath tags and occurs when a graphical " _
        & "object or text box is anchored to the same paragraph that contains the equation.""", _
        MsgBoxStyle.Information)

        Launch_URL("http://support.microsoft.com/kb/2528942")

    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click

        Launch_URL("http://legacy.s2services.com/word_repair.htm")

    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

        Launch_URL("http://legacy.s2services.com/word_repair.htm")

    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click

    End Sub

    Private Sub ToolStripMenuItem10_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem10.Click

        MsgBox("It is important to stress the need to install your Microsoft Office version's " _
            & "most up to date service packs and hotfixes to keep Word document corruption " _
            & "to a minimum. If you haven't already done so, please use the buttons of " _
            & "this programs to do it. The Mr. Fixit, fixes Math related end tag out of " _
            & "order issues that this present program also fixes.", MsgBoxStyle.Information)

        TextBox1.Text = Nothing
        SelectFile()

        If String.IsNullOrEmpty(TextBox1.Text) = True Then

            SelectFile()

        Else

            Impl_Automatic_Recovery()

        End If

    End Sub

    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click

        Try

            Me.Close()

        Catch ex As Exception

            'MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub FreeServicesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FreeServicesToolStripMenuItem.Click

    End Sub

    Private Sub ShadowExplorerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShadowExplorerToolStripMenuItem.Click

        Launch_URL("http://www.shadowexplorer.com/")

    End Sub

    Private Sub PreviousVersionFileRecovererToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PreviousVersionFileRecovererToolStripMenuItem.Click

        Launch_URL("http://sourceforge.net/projects/vistaprevrsrcvr/")

    End Sub

    Private Sub UnsavedOfficeFileRecoveryStepsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UnsavedOfficeFileRecoveryStepsToolStripMenuItem.Click

        Launch_URL("http://www.makeuseof.com/tag/recover-unsaved-ms-word-2010-document-seconds/")

    End Sub

    Private Sub StepsToRecoverAWordFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StepsToRecoverAWordFileToolStripMenuItem.Click

        Launch_URL("http://legacy.s2services.com/word_repair.htm")

    End Sub

    Private Sub FreewareToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FreewareToolStripMenuItem.Click

    End Sub

    Private Sub S2ServicesCorruptFileRecoveryFreewareToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles S2ServicesCorruptFileRecoveryFreewareToolStripMenuItem.Click

        Launch_URL("http://sourceforge.net/users/socrtwo22")

    End Sub

    Private Sub StepsToRecoveringAnOpenOfficeFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StepsToRecoveringAnOpenOfficeFileToolStripMenuItem.Click

        Launch_URL("http://legacy.s2services.com/open_office.htm")

    End Sub

    Private Sub GetDatasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetDatasToolStripMenuItem.Click

        Launch_URL("http://www.repairmyword.com/")

    End Sub

    Private Sub MicrosoftsFreeOffVisMightHelpWithDOCCorruptionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MicrosoftsFreeOffVisMightHelpWithDOCCorruptionToolStripMenuItem.Click

        MsgBox("OffVis is an Advanced tool for expert analysis of doc format files. Luckily " _
               & "there is an automatic repair featue on the Tools Menu. Be patient as " _
               & "there is no progress indicator. Savvy Recovery for Word will now " _
               & "open the download link to the white paper that describes the tool and " _
               & "will engage the link to directly download the program from: " _
               & "http://go.microsoft.com/fwlink/?LinkId=158791.", MsgBoxStyle.Exclamation)

        Launch_URL("http://www.microsoft.com/en-us/download/details.aspx?id=2096")

        Launch_URL("http://go.microsoft.com/fwlink/?LinkId=158791")

    End Sub

    Private Sub SilverCodersDocToTextToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SilverCodersDocToTextToolStripMenuItem.Click

        Launch_URL("http://silvercoders.com/en/products/doctotext/")

    End Sub

    Private Sub TryFreeOnlineServiceUseCouponS2SERVICESToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TryFreeOnlineServiceUseCouponS2SERVICESToolStripMenuItem.Click

        MsgBox("Until Nov. 1, 2014, for a free $39 value file repair attempt, first choose your file and then click " _
               & "on the green button ""Secure Upload and Repair"". When repair is complete choose the first link " _
               & """Download demo-repaired file."" In the next window, scroll down past ""Demo Results"" and enter in " _
               & "the coupon code ""S2SERVICES"" in the field above the ""Submit Code"" button at the end of the ""Full " _
               & "Results"" section. Use all caps for the code but don't include the quotes.", MsgBoxStyle.Information)

        Launch_URL("https://online.officerecovery.com/")

    End Sub

    Private Sub StepsToRecoverOfficeFilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StepsToRecoverOfficeFilesToolStripMenuItem.Click

    End Sub

    Private Sub ToolStripMenuItem17_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem17.Click

        Launch_URL("http://www.cimaware.com/info/info.php?lang=en&id=622&path=wordfix.html")

    End Sub

    Private Sub OfficeFixToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OfficeFixToolStripMenuItem.Click

        Launch_URL("http://www.cimaware.com/info/info.php?lang=en&id=622&path=officefix.html")

    End Sub

    Private Sub TryOurManual22RepairToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TryOurManual22RepairToolStripMenuItem.Click

        Launch_URL("http://saveofficedata.com/contact.htm")

    End Sub

    Private Sub UploadToZohoOfficeManyUnspecifiedAndAlsoMaybeEndTagErrrorsWillLoadInZohoWriterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UploadToZohoOfficeManyUnspecifiedAndAlsoMaybeEndTagErrrorsWillLoadInZohoWriterToolStripMenuItem.Click

        MsgBox("Most Word DOCX documents of the ""Unspecified Error line 2, " _
               & "Column 0"" variety, will open in Zoho writer. From there " _
               & "they can be exported to a Word DOC or Open Office ODT format.", MsgBoxStyle.Information)

        Launch_URL("https://www.zoho.com/")

    End Sub

    Private Sub MrFixitThisAppletWillFixEndTagNotMatchingErrorsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MrFixitThisAppletWillFixEndTagNotMatchingErrorsToolStripMenuItem.Click

        MsgBox("This Mr Fixit applet will directly fix Word 2007-2013 files having the error: " _
               & """The name in the end tag of the element must match the element type in the " _
               & "start tag."" After applying the applet the problem is not permanently fixed and " _
               & "can reoccur. The issue is fixed in Office 2013 and Office 2010 Service Pack 1. " _
               & """This issue is related strictly to oMath tags and occurs when a graphical " _
               & "object or text box is anchored to the same paragraph that contains the equation.""", _
               MsgBoxStyle.Information)

        Launch_URL("http://support.microsoft.com/kb/2528942")

    End Sub

    Private Sub SubmitToAMicrosoftUserGroupForAWOrdMVPOrExperToFixToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SubmitToAMicrosoftUserGroupForAWOrdMVPOrExperToFixToolStripMenuItem.Click

        MsgBox("This button opens links to two Microsoft community threads where MVPs and " _
               & "other Word experts will manually fix your ""Unspecified"" error Word files " _
               & "for free. There are other similar Microsoft community threads.", MsgBoxStyle.Information)

        Launch_URL("http://answers.microsoft.com/en-us/office/forum/office_2010-word/unspecified-error-worddocumentxml-line2-column-0/21971fa0-df44-4ba6-ac42-7d4b5cd4174f?page=23&tab=question&status=AllReplies&tm=1380920402797")
        Launch_URL("http://social.technet.microsoft.com/Forums/office/en-US/11c6f40d-e229-428a-9d76-cf0e5be3b8c8/corrupted-word-2007-file-unspecified-error?forum=officesetupdeploylegacy")

    End Sub

    Private Sub S2ServicesFreeServiceCurrentlyUnavailableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles S2ServicesFreeServiceCurrentlyUnavailableToolStripMenuItem.Click

        MsgBox("To get a coupon code for 10 free recoveries, leave a Tweet, Facebook  " _
               & "or blog entry. See the pricing tab for the link and details.", MsgBoxStyle.Information)

        Launch_URL("http://onlinerecovery.munsoft.com/")

    End Sub

    Private Sub HotfixForWord2007ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HotfixForWord2007ToolStripMenuItem.Click

        MsgBox("This hotfix for Office and Word 2007, repairs the ""Unspecified error"" issue " _
               & "for future documents. With Word 2007 Unspecified Errors, the issue occurs " _
               & "after the installation of security patch MS08-072 and appears to be different " _
               & "than Word 2010 unspecified errors in that a column/character number is given " _
               & "whereas with the Word 2010 errors of this type, the column number is 0. The " _
               & "hotfix is not contained in Service Pack 1 or 2 of Office 2007 and it requires " _
               & "either service pack to be installed in order to work.", MsgBoxStyle.Information)

        Launch_URL("http://support.microsoft.com/kb/970942")

    End Sub

    Private Sub HotfixForWord2010ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HotfixForWord2010ToolStripMenuItem.Click

        MsgBox("This hotfix fix for Office and Word 2010 fixes the ""Unspecified error"" " _
               & "issue for future documents. The hotfix is not contained in Service Pack " _
               & "1 or 2 of Office 2010 and it requires either service pack to be installed " _
               & "in order to work.", MsgBoxStyle.Information)

        Launch_URL("http://support.microsoft.com/kb/2817583")

    End Sub

    Private Sub DownloadAndInstallTheLatestOfficeServicePackToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DownloadAndInstallTheLatestOfficeServicePackToolStripMenuItem.Click

        MsgBox("Installing the Service Pack 3 for Office 2007 may fix the unspecified error " _
               & "connected to the table of contents markup and the w:fldChar tags. Of course " _
               & "it is always good practice to install the latest service packs. When writing " _
               & "this program on Oct. 19, 2013, Microsoft Office 2007 is up to Service Pack 3, " _
               & "Office 2010 is up to Service Pack 2 and Office 2013 has not had a service pack " _
               & "released yet. Service packs are cumulative, meaning if you have Office 2007, you " _
               & "can skip from SP1 to SP3 as SP2 (and SP1 for that matter) are included in SP3.", _
               MsgBoxStyle.Information)

        Launch_URL("http://www.microsoft.com/en-us/download/office-service-packs.aspx")

    End Sub

    Private Sub ToolStripMenuItem16_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem16.Click

        Launch_URL("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=UBLX24NP4S44L")

    End Sub

    Private Sub ToolStripMenuItem7_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click

        Launch_URL("http://s2services.com/quick-word-help.htm")

    End Sub

    Private Sub ToolStripMenuItem12_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem12.Click

        Impl_Automatic_Recovery()

    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click

        AboutBox1.Show()

    End Sub

    Private Sub ToolStripMenuItem15_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem15.Click

        Deal_With_Salvage_Text_Treatment_Menu_Choice()

    End Sub

    Private Sub RecoverButton_Click(sender As Object, e As EventArgs) Handles RecoverButton.Click

        Impl_Automatic_Recovery()
        ProgressBar1.Value = ProgressBar1.Maximum
        ProgressBar1.Hide()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click

        MsgBox("The Mr. Fixit solution works to fix the Math Tag out of order problem which gives " _
            & "rise to the error: ""The name in the end tag of the element must match the " _
            & "element type in the start tag"". However the Mr. Fixit does not appear to work " _
            & "if Word 2013 and/or Windows 8.1 is installed. Tony Jolan's Word template will " _
            & "work in Windows 8.1 and if you have Word 2013 installed. Unfortunately it has " _
            & "become apparent that there are other end causes of end tag not matching errors. " _
            & "Savvy DOCX recovery tries to repair all of these.", MsgBoxStyle.Information)

        Launch_URL("http://www.wordarticles.com/temp/Rebuilder.dotm")

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            Directory.Delete(tempPath, True)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CloseWordApplicationWithDocument(objWord As Microsoft.Office.Interop.Word.Application, objDoc As Microsoft.Office.Interop.Word.Document)
        Try
            If (objDoc IsNot Nothing) Then
                objDoc.Close(False)
                Marshal.FinalReleaseComObject(objDoc)

                objDoc = Nothing
            End If


            If (objWord IsNot Nothing) Then
                objWord.Quit(False)
                Marshal.FinalReleaseComObject(objWord)

                objWord = Nothing
            End If

            'Garbage collecting code along with the Marshal.FinalReleaseComObject 
            'necessary to remove the instances of Word started by the program.

            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' GC needs to be called twice in order to get the Finalizers called  
            ' - the first time in, it simply makes a list of what is to be  
            ' finalized, the second time in, it actually is finalizing. Only  
            ' then will the object do its automatic ReleaseComObject. 

            GC.Collect()
            GC.WaitForPendingFinalizers()
        Catch ex As Exception

        End Try

    End Sub

End Class
