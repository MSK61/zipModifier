/************************************************************

program:      zip modifier

file:         MainForm.cs

function:     methods of the MainForm class

description:  modifies the contents of zip files

author:       Mohammed Safwat (MS)

environment:  visual studio.net enterprise architect 2003, windows xp
              professional

notes:        This is a private program.

revisions:    1.00  2/11/2005 (MS) starting construction
              1.01  6/11/2005 (MS) first release
              1.02  14/11/2005 (MS) fixing the bug concerning
              directory entries in a zip file
              1.03  14/11/2005 (MS) fixing the bug concerning writing
              uninitialized parts of buffers to the output zip file
              1.5  6/1/2005   (MS) fixing some variable names and
              code formatting

************************************************************/
using java.io;
using java.util;
using java.util.zip;
using System;
using System.Drawing;
using System.Collections;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using System.Data;

namespace ZipModifier
{
    /// <summary>
    /// Summary description for MainForm.
    /// </summary>
    public class MainForm : System.Windows.Forms.Form
    {
        private System.Windows.Forms.Label zipPromptLabel;
        private System.Windows.Forms.Button browseButton;
        private System.Windows.Forms.TextBox zipTextBox;
        private System.Windows.Forms.Label patternLabel;
        private System.Windows.Forms.TextBox patternTextBox;
        private System.Windows.Forms.Button findButton;
        private System.Windows.Forms.ListBox resultListBox;
        private System.Windows.Forms.Button deleteButton;
        private System.Windows.Forms.Label resultLabel;
        private System.Windows.Forms.Label matchLabel;
        private System.Windows.Forms.Button stopButton;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;

        public MainForm()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            //
            // TODO: Add any constructor code after InitializeComponent call
            //
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose( bool disposing )
        {
            if( disposing )
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose( disposing );
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.zipPromptLabel = new System.Windows.Forms.Label();
            this.browseButton = new System.Windows.Forms.Button();
            this.zipTextBox = new System.Windows.Forms.TextBox();
            this.patternLabel = new System.Windows.Forms.Label();
            this.patternTextBox = new System.Windows.Forms.TextBox();
            this.findButton = new System.Windows.Forms.Button();
            this.resultListBox = new System.Windows.Forms.ListBox();
            this.deleteButton = new System.Windows.Forms.Button();
            this.resultLabel = new System.Windows.Forms.Label();
            this.matchLabel = new System.Windows.Forms.Label();
            this.stopButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            //
            // zipPromptLabel
            //
            this.zipPromptLabel.AutoSize = true;
            this.zipPromptLabel.Location = new System.Drawing.Point(8, 16);
            this.zipPromptLabel.Name = "zipPromptLabel";
            this.zipPromptLabel.Size = new System.Drawing.Size(41, 16);
            this.zipPromptLabel.TabIndex = 0;
            this.zipPromptLabel.Text = "&Zip file:";
            //
            // browseButton
            //
            this.browseButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.browseButton.Location = new System.Drawing.Point(224, 16);
            this.browseButton.Name = "browseButton";
            this.browseButton.Size = new System.Drawing.Size(56, 23);
            this.browseButton.TabIndex = 2;
            this.browseButton.Text = "&Browse";
            this.browseButton.Click += new System.EventHandler(this.browseButton_Click);
            //
            // zipTextBox
            //
            this.zipTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                | System.Windows.Forms.AnchorStyles.Right)));
            this.zipTextBox.Location = new System.Drawing.Point(48, 16);
            this.zipTextBox.Name = "zipTextBox";
            this.zipTextBox.Size = new System.Drawing.Size(168, 20);
            this.zipTextBox.TabIndex = 1;
            this.zipTextBox.Text = "";
            this.zipTextBox.TextChanged += new System.EventHandler(this.zipTextBox_TextChanged);
            //
            // patternLabel
            //
            this.patternLabel.AutoSize = true;
            this.patternLabel.Location = new System.Drawing.Point(8, 48);
            this.patternLabel.Name = "patternLabel";
            this.patternLabel.Size = new System.Drawing.Size(95, 16);
            this.patternLabel.TabIndex = 3;
            this.patternLabel.Text = "&Pattern(s) to Find:";
            //
            // patternTextBox
            //
            this.patternTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                | System.Windows.Forms.AnchorStyles.Right)));
            this.patternTextBox.Location = new System.Drawing.Point(96, 48);
            this.patternTextBox.Name = "patternTextBox";
            this.patternTextBox.Size = new System.Drawing.Size(120, 20);
            this.patternTextBox.TabIndex = 4;
            this.patternTextBox.Text = "";
            this.patternTextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.patternTextBox_KeyDown);
            //
            // findButton
            //
            this.findButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.findButton.Location = new System.Drawing.Point(224, 48);
            this.findButton.Name = "findButton";
            this.findButton.Size = new System.Drawing.Size(56, 23);
            this.findButton.TabIndex = 5;
            this.findButton.Text = "&Find";
            this.findButton.Click += new System.EventHandler(this.findButton_Click);
            //
            // resultListBox
            //
            this.resultListBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                | System.Windows.Forms.AnchorStyles.Left)
                | System.Windows.Forms.AnchorStyles.Right)));
            this.resultListBox.HorizontalScrollbar = true;
            this.resultListBox.Location = new System.Drawing.Point(8, 96);
            this.resultListBox.Name = "resultListBox";
            this.resultListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.resultListBox.Size = new System.Drawing.Size(208, 160);
            this.resultListBox.TabIndex = 6;
            //
            // deleteButton
            //
            this.deleteButton.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.deleteButton.Location = new System.Drawing.Point(224, 152);
            this.deleteButton.Name = "deleteButton";
            this.deleteButton.Size = new System.Drawing.Size(56, 32);
            this.deleteButton.TabIndex = 7;
            this.deleteButton.Text = "&Delete selected";
            this.deleteButton.Click += new System.EventHandler(this.deleteButton_Click);
            //
            // resultLabel
            //
            this.resultLabel.AutoSize = true;
            this.resultLabel.Location = new System.Drawing.Point(8, 80);
            this.resultLabel.Name = "resultLabel";
            this.resultLabel.Size = new System.Drawing.Size(35, 16);
            this.resultLabel.TabIndex = 8;
            this.resultLabel.Text = "result:";
            //
            // matchLabel
            //
            this.matchLabel.AutoSize = true;
            this.matchLabel.Location = new System.Drawing.Point(48, 80);
            this.matchLabel.Name = "matchLabel";
            this.matchLabel.Size = new System.Drawing.Size(0, 16);
            this.matchLabel.TabIndex = 9;
            //
            // stopButton
            //
            this.stopButton.Enabled = false;
            this.stopButton.Location = new System.Drawing.Point(224, 80);
            this.stopButton.Name = "stopButton";
            this.stopButton.Size = new System.Drawing.Size(56, 23);
            this.stopButton.TabIndex = 10;
            this.stopButton.Text = "&Stop";
            this.stopButton.Click += new System.EventHandler(this.stopButton_Click);
            //
            // MainForm
            //
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(292, 266);
            this.Controls.Add(this.stopButton);
            this.Controls.Add(this.matchLabel);
            this.Controls.Add(this.resultLabel);
            this.Controls.Add(this.deleteButton);
            this.Controls.Add(this.resultListBox);
            this.Controls.Add(this.findButton);
            this.Controls.Add(this.patternTextBox);
            this.Controls.Add(this.patternLabel);
            this.Controls.Add(this.zipTextBox);
            this.Controls.Add(this.browseButton);
            this.Controls.Add(this.zipPromptLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Zip Modifier";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);

        }
        #endregion

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.Run(new MainForm());
        }

        private void browseButton_Click(object sender, System.EventArgs e)
        {
            OpenFileDialog zipFileSelector = new OpenFileDialog();

            // Initialize an open file dialog box to select a zip
            // file.
            zipFileSelector.Title = "Select Zip File";
            zipFileSelector.Filter = "Zip Files (*.zip)|*.zip";

            if (zipFileSelector.ShowDialog() == DialogResult.OK)
                zipTextBox.Text = zipFileSelector.FileName;// If the
            // user has selected a zip file, show the selected zip
            // file name in the text box.

        }// end of method browseButton_Click

        /// <summary>
        /// This method clears the results of any previous serach.
        /// </summary>
        private void ClearResults()
        {

            resultListBox.Items.Clear();// Empty the result list from
            // any prior results.
            matchLabel.Text = String.Empty;// Clear the label of
            // matches.

        }// end of method ClearResults

        private void deleteButton_Click(object sender, System.EventArgs e)
        {
            ZipEntry currentEntry;// current entry in the source zip
                // file
            sbyte[] currentEntryContents;// contents of the current
                // entry in the source zip file
            BufferedInputStream currentEntryStream;
            int readBytes;// number of bytes read from a stream at a
                // time
            ulong remainingBytes;// number of remaining bytes in a
                // stream
            Enumeration sourceEntries;// entries in the source zip
                // file
            ZipFile sourceZipFile = null;
            string tempFileName;
            ZipOutputStream tempZipFile = null;

            if (resultListBox.SelectedItems.Count == 0)// if no
                // entries within the selected zip file are selected
                MessageBox.Show("Either no zip file is selected or " +
                    "no entries in the selected one are selected. " +
                    "Please, make sure that a valid zip file is " +
                    "selected and entries within the selected one " +
                    "are selected.", "Missing Data",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            else if ((System.IO.File.GetAttributes(itsZipFile) &
                FileAttributes.ReadOnly) == 0)// if the selected zip
                // file can be modified
            {

                ToggleControls(false);// Disable the user interface.
                try
                {

                    sourceZipFile = new ZipFile(itsZipFile);// Open
                        // the source zip file.
                    tempFileName = Path.GetTempFileName();// Get a
                        // temporary file name.
                    tempZipFile =
                        new ZipOutputStream(new BufferedOutputStream(
                        new FileOutputStream(tempFileName)));// Open
                        // a temporary zip file with the name just
                        // obtained.
                    sourceEntries = sourceZipFile.entries();// Get the
                        // entries of the source zip file.

                    while (sourceEntries.hasMoreElements())// Copy all
                        // the entries of the source zip file to the
                        // temporary zip file except for the entries
                        // selected to be deleted.
                    {

                        // Get the current entry in the source zip
                        // file.
                        currentEntry =
                            (ZipEntry)(sourceEntries.nextElement());

                        if (! resultListBox.SelectedItems.Contains(
                            currentEntry.getName()) &&
                            ! currentEntry.isDirectory())// if the
                            // current entry in the source zip file
                            // isn't to be deleted and isn't a
                            // directory
                        {

                            // Create an identical entry in the
                            // temporary zip file.
                            tempZipFile.putNextEntry(currentEntry);
                            currentEntryStream =
                                new BufferedInputStream(
                                sourceZipFile.getInputStream(
                                currentEntry));// Obtain a stream for
                                // the current source entry.
                            remainingBytes =
                                (ulong)currentEntry.getSize();
                            currentEntryContents =
                                new sbyte[remainingBytes];

                            // Read the current entry in the source
                            // zip file until all its contents are
                            // read.
                            for (; remainingBytes > 0;
                                remainingBytes -= (ulong)readBytes)
                            {

                                readBytes = currentEntryStream.read(
                                    currentEntryContents, 0,
                                    (int)(remainingBytes));// Get the
                                    // contents of the current entry
                                    // in the source zip file.
                                tempZipFile.write(
                                    currentEntryContents, 0,
                                    readBytes);// Write the contents
                                    // of the current entry in the
                                    // source zip file to the
                                    // corresponding entry in the
                                    // temporary zip file.

                            }

                            currentEntryStream.close();// Close the
                                // current entrty stream in the source
                                // zip file.
                            tempZipFile.closeEntry();// Close the
                                // entry in the temporary zip file.

                        }

                    }

                    sourceZipFile.close();// Close the source zip
                        // file.
                    sourceZipFile = null;// Prevent the source zip
                        // file form being closed again in the future.
                    System.IO.File.Delete(itsZipFile);// Delete the
                        // source zip file so that the modified zip
                        // file may be moved.
                    tempZipFile.close();// close the temporary zip
                        // file so that it may be moved.
                    tempZipFile = null;// Prevent the temporary zip
                    // file form being closed again in the future.
                    // Move the modified zip file so that it may
                    // replace the original zip file.
                    System.IO.File.Move(tempFileName, itsZipFile);
                    resultListBox.Hide();// Hide the result list box
                    // to overcome the flicker that may happen
                    // during deleting the old entries.

                    while (resultListBox.SelectedItems.Count > 0)
                        resultListBox.Items.Remove(
                            resultListBox.SelectedItems[0]);// Delete
                        // the old entries.

                    resultListBox.Show();// Show the list box after
                    // deleting the entries.
                    UpdateMatchLabel();

                }
                catch (java.io.FileNotFoundException error)
                {

                    MessageBox.Show(
                        "The file '" + error.getMessage() +
                        "' doesn't exist. Please, select another " +
                        "file.", "File Not Found",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                catch (ArgumentException error)
                {

                    MessageBox.Show(
                        error.Message, "Invalid File Name",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                catch (DirectoryNotFoundException error)
                {

                    MessageBox.Show(
                        error.Message, "Invalid directory",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    // Highlight the zip file text box.
                    zipTextBox.Focus();
                    zipTextBox.SelectAll();

                }
                finally// Close the used file streams.
                {

                    if (sourceZipFile != null) sourceZipFile.close();

                    if (tempZipFile != null) tempZipFile.close();

                }
                ToggleControls(true);// Enable the user interface.

            }
            else MessageBox.Show(
                     "The selected zip file is read-only. It can't " +
                     "be modified.", "Invalid Operation",
                     MessageBoxButtons.OK, MessageBoxIcon.Error);// If
                // the selected zip file is read-only, notify the
                // user.

        }// end of method deleteButton_Click

        private void findButton_Click(object sender, System.EventArgs e)
        {
            ushort count;
            ZipEntry currentEntry;
            string currentEntryName,
                lowerCaseEntryName;// lower case version of the
                // current entry name
            int emptyStringIndex;// index of an empty string in the
                // array of extensions
            // array to hold the extensions to match
            StringCollection extensions = new StringCollection();
            Enumeration zipEntries;// entries of the zip file
            ZipFile zipFile;

            ToggleControls(false);// Disable the user interface.
            try
            {

                extensions.AddRange(patternTextBox.Text.Replace(
                    " ", String.Empty).Split(new char[]{';'}));// Get
                    // the file extensions the user wants to search
                    // for inside the selected zip file.

                for (emptyStringIndex = extensions.IndexOf(String.Empty);
                    emptyStringIndex != -1;
                    emptyStringIndex = extensions.IndexOf(String.Empty))
                    extensions.RemoveAt(emptyStringIndex);// Filter the
                // array of extensions so that it may contain only
                // valid extensions.

                if (extensions.Count == 0)// if no valid extensions
                    // exist
                {

                    MessageBox.Show("Please, enter valid extensions.",
                        "Invalid extensions", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);// Prompt the user.
                    // Highlight the text box of the extensions.
                    patternTextBox.Focus();
                    patternTextBox.SelectAll();

                }
                else
                {

                    stopButton.Enabled = true;

                    // Add the period extension notation to each
                    // extension.
                    for (count = 0; count < extensions.Count; count++)
                        extensions[count] =
                            '.' + extensions[count].ToLower();

                    zipFile = new ZipFile(itsZipFile);
                    zipEntries = zipFile.entries();// Get the entries
                        // of the zip file.
                    ClearResults();// Clear any previous results.
                    resultListBox.Hide();// Hide the result list box
                        // to overcome the flicker that may happen
                        // during publishing the zip entries.

                    while (! itsCancelSearch &&
                        zipEntries.hasMoreElements())// Publish all
                        // the entries in the zip file that match one
                        // of the pattern extensions.
                    {

                        Application.DoEvents();// Process any current
                            // events.
                        // Get the current entry.
                        currentEntry =
                            (ZipEntry)(zipEntries.nextElement());

                        if (! currentEntry.isDirectory())
                        {

                            currentEntryName = currentEntry.getName();
                            lowerCaseEntryName =
                                currentEntryName.ToLower();

                            for (count = 0;
                                count < extensions.Count &&
                                ! lowerCaseEntryName.EndsWith(
                                extensions[count]); count++);// See if
                                // the current entry matches any of
                                // the pattern extensions.

                            if (count < extensions.Count)
                                resultListBox.Items.Add(
                                    currentEntryName);// If the
                                // current entry matches any of the
                                // pattern extensions, add this entry
                                // to the list.

                        }

                    }

                    zipFile.close();// Close the source zip file.
                    stopButton.Enabled = false;
                    UpdateMatchLabel();
                    resultListBox.Show();// Show the list box after
                        // publishing.

                    if (itsCancelSearch) itsCancelSearch = false;// If
                        // the search was cancelled by the user,
                        // disable the cancel signal to allow future
                        // searches.

                }

            }
            catch (java.io.FileNotFoundException error)
            {

                MessageBox.Show("The file '" + error.getMessage() +
                    "' doesn't exist. Please, select another file.",
                    "File Not Found", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                zipTextBox.Focus();

            }
            catch (ArgumentException error)
            {

                MessageBox.Show(error.Message, "Invalid File Name",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Highlight the zip file text box.
                zipTextBox.Focus();
                zipTextBox.SelectAll();

            }
            ToggleControls(true);// Enable the user interface.

        }// end of method findButton_Click

        private void MainForm_Load(object sender, System.EventArgs e)
        {

            // Position the deletion button visually in a suitable
            // way.
            deleteButton.Top = resultListBox.Top +
                (resultListBox.Height - deleteButton.Height) / 2;

        }// end of method MainForm_Load

        private void patternTextBox_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Enter)// if the user has pressed the
                // enter key
                findButton_Click(findButton, e);// Find the entered
                // extensions within the zip file.

        }// end of method patternTextBox_KeyDown

        private void stopButton_Click(object sender, System.EventArgs e)
        {

            itsCancelSearch = true;

        }// end of method stopButton_Click

        /// <summary>
        /// This method either enables or disables user interface controls.
        /// </summary>
        private void ToggleControls(bool state)
        {

            foreach (Control count in Controls)
                if (count != stopButton) count.Enabled = state;

            if (state) Cursor = Cursors.Default;
            else
            {

                Cursor = Cursors.WaitCursor;
                stopButton.Cursor = Cursors.Default;

            }

        }// end of method ToggleControls

        /// <summary>
        /// This method updates the match label with the specified number of matches.
        /// </summary>
        private void UpdateMatchLabel()
        {

            // Report the number of matches.
            matchLabel.Text =
                "found " + resultListBox.Items.Count + " matches";

        }// end of method UpdateMatchLabel

        private void zipTextBox_TextChanged(object sender, System.EventArgs e)
        {

            if (resultListBox.Items.Count == 0)// if no results for a
                // previous serach exist
                itsZipFile = zipTextBox.Text;
            // If the user has selected a different zip file while
            // there are results for a previous search, make sure the
            // user wants to overwrite the old results.
            else if (
                String.Compare(zipTextBox.Text, itsZipFile, true) !=
                0) if (MessageBox.Show(
                       "There're relevant data to a previous " +
                       "search. Selecting a new zip file will wipe " +
                       "out these data. Are you sure you want to " +
                       "select a new zip file?", "Overwrite Data?",
                       MessageBoxButtons.YesNo,
                       MessageBoxIcon.Warning) == DialogResult.Yes)
                   {

                       ClearResults();// Wipe out the old results.
                       itsZipFile = zipTextBox.Text;// Change the
                        // selected zip file.

                   }
                   else zipTextBox.Text = itsZipFile;// If the user
                // refused to overwrite the old results, reset the
                // text box of the selected zip file to the currently
                // selected file.

        }// end of method zipTextBox_TextChanged

        /// <summary>
        /// cancels the running search if it's true
        /// </summary>
        private bool itsCancelSearch = false;
        /// <summary>
        /// the selected zip file
        /// </summary>
        private string itsZipFile = "";
    }// end of class MainForm
}
