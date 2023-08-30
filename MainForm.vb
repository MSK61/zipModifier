'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' program:      zip modifier

' file:         MainForm.vb

' function:     methods of the MainForm class

' description:  modifies the contents of zip files

' author:       Mohammed Safwat (MS)

' environment:  visual studio.net enterprise architect 2003,
'               windows xp professional()

' notes:        This is a private program.

' revisions:    1.00  31/12/2005 (MS) starting construction
'               1.01  6/1/2005   (MS) first release

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Imports java.io
Imports java.util
Imports java.util.zip
Imports System.Collections.Specialized
Imports System.IO
Public Class MainForm
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Private WithEvents lblZipPrompt As System.Windows.Forms.Label
    Private WithEvents txtZip As System.Windows.Forms.TextBox
    Private WithEvents cmdBrowse As System.Windows.Forms.Button
    Private WithEvents lblPattern As System.Windows.Forms.Label
    Private WithEvents txtPattern As System.Windows.Forms.TextBox
    Private WithEvents cmdFind As System.Windows.Forms.Button
    Private WithEvents lstResult As System.Windows.Forms.ListBox
    Private WithEvents cmdDelete As System.Windows.Forms.Button
    Private WithEvents lblResult As System.Windows.Forms.Label
    Private WithEvents lblMatch As System.Windows.Forms.Label
    Private WithEvents cmdStop As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblZipPrompt = New System.Windows.Forms.Label
        Me.txtZip = New System.Windows.Forms.TextBox
        Me.cmdBrowse = New System.Windows.Forms.Button
        Me.lblPattern = New System.Windows.Forms.Label
        Me.txtPattern = New System.Windows.Forms.TextBox
        Me.cmdFind = New System.Windows.Forms.Button
        Me.lstResult = New System.Windows.Forms.ListBox
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.lblResult = New System.Windows.Forms.Label
        Me.lblMatch = New System.Windows.Forms.Label
        Me.cmdStop = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lblZipPrompt
        '
        Me.lblZipPrompt.AutoSize = True
        Me.lblZipPrompt.Location = New System.Drawing.Point(8, 16)
        Me.lblZipPrompt.Name = "lblZipPrompt"
        Me.lblZipPrompt.Size = New System.Drawing.Size(41, 16)
        Me.lblZipPrompt.TabIndex = 0
        Me.lblZipPrompt.Text = "&Zip file:"
        '
        'txtZip
        '
        Me.txtZip.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtZip.Location = New System.Drawing.Point(48, 16)
        Me.txtZip.Name = "txtZip"
        Me.txtZip.Size = New System.Drawing.Size(168, 20)
        Me.txtZip.TabIndex = 1
        Me.txtZip.Text = ""
        '
        'cmdBrowse
        '
        Me.cmdBrowse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowse.Location = New System.Drawing.Point(224, 16)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.Size = New System.Drawing.Size(56, 23)
        Me.cmdBrowse.TabIndex = 2
        Me.cmdBrowse.Text = "&Browse"
        '
        'lblPattern
        '
        Me.lblPattern.AutoSize = True
        Me.lblPattern.Location = New System.Drawing.Point(8, 48)
        Me.lblPattern.Name = "lblPattern"
        Me.lblPattern.Size = New System.Drawing.Size(95, 16)
        Me.lblPattern.TabIndex = 3
        Me.lblPattern.Text = "&Pattern(s) to Find:"
        '
        'txtPattern
        '
        Me.txtPattern.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPattern.Location = New System.Drawing.Point(96, 48)
        Me.txtPattern.Name = "txtPattern"
        Me.txtPattern.Size = New System.Drawing.Size(120, 20)
        Me.txtPattern.TabIndex = 4
        Me.txtPattern.Text = ""
        '
        'cmdFind
        '
        Me.cmdFind.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdFind.Location = New System.Drawing.Point(224, 48)
        Me.cmdFind.Name = "cmdFind"
        Me.cmdFind.Size = New System.Drawing.Size(56, 23)
        Me.cmdFind.TabIndex = 5
        Me.cmdFind.Text = "&Find"
        '
        'lstResult
        '
        Me.lstResult.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstResult.HorizontalScrollbar = True
        Me.lstResult.Location = New System.Drawing.Point(8, 96)
        Me.lstResult.Name = "lstResult"
        Me.lstResult.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lstResult.Size = New System.Drawing.Size(208, 160)
        Me.lstResult.TabIndex = 6
        '
        'cmdDelete
        '
        Me.cmdDelete.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.cmdDelete.Location = New System.Drawing.Point(224, 152)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(56, 32)
        Me.cmdDelete.TabIndex = 7
        Me.cmdDelete.Text = "&Delete selected"
        '
        'lblResult
        '
        Me.lblResult.AutoSize = True
        Me.lblResult.Location = New System.Drawing.Point(8, 80)
        Me.lblResult.Name = "lblResult"
        Me.lblResult.Size = New System.Drawing.Size(35, 16)
        Me.lblResult.TabIndex = 8
        Me.lblResult.Text = "result:"
        '
        'lblMatch
        '
        Me.lblMatch.AutoSize = True
        Me.lblMatch.Location = New System.Drawing.Point(48, 80)
        Me.lblMatch.Name = "lblMatch"
        Me.lblMatch.Size = New System.Drawing.Size(0, 16)
        Me.lblMatch.TabIndex = 9
        '
        'cmdStop
        '
        Me.cmdStop.Enabled = False
        Me.cmdStop.Location = New System.Drawing.Point(224, 80)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(56, 23)
        Me.cmdStop.TabIndex = 10
        Me.cmdStop.Text = "&Stop"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.lblMatch)
        Me.Controls.Add(Me.lblResult)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.lstResult)
        Me.Controls.Add(Me.cmdFind)
        Me.Controls.Add(Me.txtPattern)
        Me.Controls.Add(Me.lblPattern)
        Me.Controls.Add(Me.cmdBrowse)
        Me.Controls.Add(Me.txtZip)
        Me.Controls.Add(Me.lblZipPrompt)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Zip Modifier"
        Me.ResumeLayout(False)

    End Sub

#End Region

    ' This method clears the results of any previous serach.
    Private Sub ClearResults()

        lstResult.Items.Clear() ' Empty the result list from any
        ' prior results.
        lblMatch.Text = String.Empty

    End Sub ' end of method ClearResults

    Private Sub cmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowse.Click

        Dim zipFileSelector As New OpenFileDialog

        ' Initialize an open file dialog box to select a zip
        ' file.
        zipFileSelector.Title = "Select Zip File"
        zipFileSelector.Filter = "Zip Files (*.zip)|*.zip"

        ' If the user has selected a zip file, show the selected
        ' zip file name in the text box.
        If zipFileSelector.ShowDialog() = DialogResult.OK Then

            txtZip.Text = zipFileSelector.FileName

        End If

    End Sub ' end of method cmdBrowse_Click

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Dim a_CurrentEntryContents() As SByte ' contents of the
        ' current entry in the source zip file
        Dim currentEntry As ZipEntry ' current entry in the
        ' source zip file
        Dim currentEntryStream As BufferedInputStream, _
            iReadBytes As Integer ' number of bytes read from a
        ' Stream at a time
        Dim lRemainingBytes As Long ' number of remaining bytes
        ' in a stream
        Dim sourceEntries As Enumeration ' entries in the source
        ' zip file
        Dim sourceZipFile As ZipFile = Nothing, _
            strTempFileName As String, _
            tempZipFile As ZipOutputStream = Nothing

        If lstResult.SelectedItems.Count = 0 Then ' if no entries
            ' within the selected zip file are selected
            MessageBox.Show("Either no zip file is selected " & _
            "or no entries in the selected one are selected" & _
            ". Please, make sure that a valid zip file is " & _
            "selected and entries within the selected one " & _
            "are selected.", "Missing Data", _
            MessageBoxButtons.OK, MessageBoxIcon.Error)

        ElseIf (CInt( _
            System.IO.File.GetAttributes(strItsZipFile)) And _
            CInt(FileAttributes.ReadOnly)) = 0 Then ' if the
            ' selected zip file can be modified.

            ToggleControls(False) ' Disable the user interface.
            Try

                sourceZipFile = New ZipFile(strItsZipFile) ' Open
                ' the source zip file.
                strTempFileName = Path.GetTempFileName() ' Get a
                ' temporary file name.
                ' Open a temporary zip file with the name just
                ' obtained.
                tempZipFile = New ZipOutputStream( _
                    New BufferedOutputStream( _
                    New FileOutputStream(strTempFileName)))
                sourceEntries = sourceZipFile.entries() ' Get the
                ' entries of the source zip file.

                Do While sourceEntries.hasMoreElements() ' Copy
                    ' all the entries of the source zip file to
                    ' the temporary zip file except for the
                    ' entries selected to be deleted.

                    ' Get the current entry in the source zip
                    ' file.
                    currentEntry = sourceEntries.nextElement()

                    If Not lstResult.SelectedItems.Contains( _
                        currentEntry.getName()) And _
                        Not currentEntry.isDirectory() Then ' if
                        ' the current entry in the source zip
                        ' file isn't to be deleted and isn't a
                        ' Directory()

                        ' Create an identical entry in the
                        ' temporary zip file.
                        tempZipFile.putNextEntry(currentEntry)
                        currentEntryStream = _
                            New BufferedInputStream( _
                            sourceZipFile.getInputStream( _
                            currentEntry)) ' Obtain a stream for
                        ' the current source entry.
                        lRemainingBytes = currentEntry.getSize()
                        ReDim a_CurrentEntryContents( _
                            lRemainingBytes - 1)

                        Do While lRemainingBytes > 0 ' Read the
                            ' current entry in the source zip
                            ' file until all its contents are
                            ' read.

                            iReadBytes = _
                                currentEntryStream.read( _
                                a_CurrentEntryContents, 0, _
                                lRemainingBytes) ' Get the
                            ' contents of the current entry in
                            ' the source zip file.
                            tempZipFile.write( _
                                a_CurrentEntryContents, 0, _
                                iReadBytes) ' Write the contents
                            ' of the current entry in the source
                            ' zip file to the corresponding entry
                            ' in the temporary zip file.
                            lRemainingBytes = lRemainingBytes - _
                                iReadBytes ' Get the contents of
                            ' the current entry in the source zip
                            ' file.

                        Loop

                        Erase a_CurrentEntryContents
                        currentEntryStream.close() ' Close the
                        ' current entrty stream in the source zip
                        ' file.
                        tempZipFile.closeEntry() ' Close the
                        ' entry in the temporary zip file.

                    End If

                Loop

                sourceZipFile.close() ' Close the source zip
                ' file.
                sourceZipFile = Nothing ' Prevent the source zip
                ' file form being closed again in the future.
                System.IO.File.Delete(strItsZipFile) ' Delete the
                ' source zip file so that the modified zip file
                ' may be moved.
                tempZipFile.close() ' close the temporary zip
                ' file so that it may be moved.
                tempZipFile = Nothing ' Prevent the temporary zip
                ' file form being closed again in the future.
                System.IO.File.Move( _
                    strTempFileName, strItsZipFile) ' Move the
                ' modified zip file so that it may replace the
                ' original zip file.
                lstResult.Hide() ' Hide the result list box to
                ' overcome the flicker that may happen during
                ' deleting the old entries.

                Do While lstResult.SelectedItems.Count > 0

                    lstResult.Items.Remove( _
                        lstResult.SelectedItems(0)) ' Delete the
                    ' old entries.

                Loop

                lstResult.Show() ' Show the list box after
                ' deleting the entries.
                UpdateMatchLabel()

            Catch ex As java.io.FileNotFoundException

                MessageBox.Show("The file '" & ex.getMessage( _
                    ) & "' doesn't exist. Please, select " & _
                    "another file.", "File Not Found", _
                    MessageBoxButtons.OK, MessageBoxIcon.Error)

            Catch ex As ArgumentException

                MessageBox.Show( _
                    ex.Message, "Invalid File Name", _
                    MessageBoxButtons.OK, MessageBoxIcon.Error)

            Catch ex As DirectoryNotFoundException

                MessageBox.Show( _
                    ex.Message, "Invalid directory", _
                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                ' Highlight the zip file text box.
                txtZip.Focus()
                txtZip.SelectAll()

            Finally ' Close the used file streams.

                If Not (sourceZipFile Is Nothing) Then

                    sourceZipFile.close()

                End If

                If Not (tempZipFile Is Nothing) Then

                    tempZipFile.close()

                End If

            End Try

            ToggleControls(True) ' Enable the user interface.

        Else ' if the selected zip file is read-only

            MessageBox.Show("The selected zip file is " & _
                "read-only. It can't be modified.", _
                "Invalid Operation", MessageBoxButtons.OK, _
                MessageBoxIcon.Error) ' Notify the user.

        End If

    End Sub ' end of method cmdDelete_Click

    Private Sub cmdFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFind.Click
        Dim bytLastExtension As Byte, _
            count As Short, _
            currentEntry As ZipEntry, _
            extensions As New StringCollection ' array to hold
        ' the extensions to match
        Dim iEmptyStringIndex As Integer ' index of an empty
        ' string in the array of extensions
        Dim strCurrentEntryName, _
            strLowerCaseEntryName As String ' lower case version
        ' of the current entry name
        Dim zipEntries As Enumeration ' entries of the zip file
        Dim zippedFile As ZipFile

        ToggleControls(False) ' Disable the user interface.
        Try

            ' Get the file extensions the user wants to search
            ' for inside the selected zip file.
            extensions.AddRange(txtPattern.Text.Replace( _
                " ", String.Empty).Split(New Char() {";"c}))
            ' Filter the array of extensions so that it may
            ' contain only valid extensions.
            iEmptyStringIndex = extensions.IndexOf(String.Empty)

            Do Until iEmptyStringIndex = -1

                extensions.RemoveAt(iEmptyStringIndex)
                iEmptyStringIndex = _
                    extensions.IndexOf(String.Empty)

            Loop

            If extensions.Count = 0 Then ' if no valid extensions
                ' exist

                ' Prompt the user.
                MessageBox.Show("Please, enter valid " & _
                    "extensions.", "Invalid extensions", _
                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                ' Highlight the text box of the extensions.
                txtPattern.Focus()
                txtPattern.SelectAll()

            Else

                cmdStop.Enabled = True
                ' Add the period extension notation to each
                ' extension.
                bytLastExtension = extensions.Count - 1

                For count = 0 To bytLastExtension

                    extensions(count) = _
                        "." & extensions(count).ToLower()

                Next count

                zippedFile = New ZipFile(strItsZipFile)
                zipEntries = zippedFile.entries() ' Get the
                ' entries of the zip file.
                ClearResults() ' Clear any previous results.
                lstResult.Hide() ' Hide the result list box to
                ' overcome the flicker that may happen during
                ' publishing the zip entries.

                Do While Not bItsCancelSearch And _
                    zipEntries.hasMoreElements() ' Publish all
                    ' the entries in the zip file that match one
                    ' of the pattern extensions.

                    Application.DoEvents() ' Process any current
                    ' events.
                    currentEntry = zipEntries.nextElement() ' Get
                    ' the current entry.

                    If Not currentEntry.isDirectory() Then

                        strCurrentEntryName = _
                            currentEntry.getName()
                        strLowerCaseEntryName = _
                            strCurrentEntryName.ToLower()
                        count = 0

                        Do While count < _
                            extensions.Count AndAlso _
                            Not strLowerCaseEntryName.EndsWith( _
                            extensions(count)) ' See if the
                            ' current entry matches any of the
                            ' pattern extensions.

                            count += 1

                        Loop

                        If count < extensions.Count Then ' if the
                            ' current entry matches any of the
                            ' pattern extensions

                            lstResult.Items.Add( _
                                strCurrentEntryName) ' Add this
                            ' entry to the list.

                        End If

                    End If

                Loop

                zippedFile.close() ' Close the source zip file.
                cmdStop.Enabled = False
                UpdateMatchLabel()
                lstResult.Show() ' Show the list box after
                ' publishing.

                ' If the search was cancelled by the user,
                ' disable the cancel signal to allow future
                ' searches.
                If bItsCancelSearch Then bItsCancelSearch = False

            End If

        Catch ex As java.io.FileNotFoundException

            MessageBox.Show("The file '" & ex.getMessage() & _
                "' doesn't exist. Please, select another " & _
                "file.", "File Not Found", _
                MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtZip.Focus()

        Catch ex As ArgumentException

            MessageBox.Show(ex.Message, "Invalid File Name", _
                MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' Highlight the zip file text box.
            txtZip.Focus()
            txtZip.SelectAll()

        End Try

        ToggleControls(True) ' Enable the user interface.

    End Sub ' end of method cmdFind_Click

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click

        bItsCancelSearch = True

    End Sub ' end of method cmdStop_Click

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        cmdDelete.Top = lstResult.Top + _
            (lstResult.Height - cmdDelete.Height) / 2

    End Sub ' end of method frmMain_Load

    ' This method either enables or disables user interface controls.
    Private Sub ToggleControls(ByVal bState As Boolean)

        For Each count As Control In Controls

            If Not (count Is cmdStop) Then count.Enabled = bState

        Next count

        If bState Then

            Cursor = Cursors.Default

        Else

            Cursor = Cursors.WaitCursor
            cmdStop.Cursor = Cursors.Default

        End If

    End Sub ' end of method ToggleControls

    Private Sub txtPattern_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPattern.KeyDown

        ' If the user has pressed the enter key, find the entered
        ' extensions within the zip file.
        If e.KeyCode = Keys.Enter Then cmdFind_Click(cmdFind, e)

    End Sub ' end of method txtPattern_KeyDown

    Private Sub txtZip_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtZip.TextChanged

        If lstResult.Items.Count = 0 Then ' if no results for a
            ' previous serach exist
            strItsZipFile = txtZip.Text
            ' If the user has selected a different zip file while
            ' there are results for a previous search, make sure
            ' the user wants to overwrite the old results.
        ElseIf String.Compare(txtZip.Text, strItsZipFile, _
            True) <> 0 Then

            If MessageBox.Show("There're relevant data to a " & _
                "previous search. Selecting a new zip file " & _
                "will wipe out these data. Are you sure you " & _
                "want to select a new zip file?", _
                "Overwrite Data?", MessageBoxButtons.YesNo, _
                MessageBoxIcon.Warning) = DialogResult.Yes Then

                ClearResults() ' Wipe out the old results.
                strItsZipFile = txtZip.Text ' Change the selected
                ' zip file.

            Else ' if the user refused to overwrite the old
                ' results

                txtZip.Text = strItsZipFile ' Reset the text box
                ' of the selected zip file to the currently
                ' selected file.

            End If

        End If

    End Sub ' end of method txtZip_TextChanged

    Private Sub UpdateMatchLabel()

        ' Report the number of matches.
        lblMatch.Text = _
            "found " & lstResult.Items.Count & " matches"

    End Sub ' end of method UpdateMatchLabel
    Private bItsCancelSearch As Boolean = False ' cancels the
    ' running search if it's true
    Private strItsZipFile As String ' the selected zip file
End Class ' end of class MainForm
