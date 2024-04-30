Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.Office.Interop

Public Class Form1

    Private emailIdToMailItem As New Dictionary(Of String, Outlook.MailItem)()
    Private addedConversationsList1 As New HashSet(Of String)
    Private addedConversationsList2 As New HashSet(Of String)
    Private addedConversationsList3 As New HashSet(Of String)
    Private dataGridOriginalItems1 As New List(Of (String, String))
    Private dataGridOriginalItems2 As New List(Of (String, String))
    Private dataGridOriginalItems3 As New List(Of (String, String))
    Private emailIdInfo As New Dictionary(Of String, (String, String)) ' Dictionary to hold EID and corresponding (name, color)


    Private WithEvents itemsEvents As Outlook.Items


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeItemsEventHandler()
        LoadEmails()

    End Sub

    Private Sub InitializeItemsEventHandler()
        Dim outlookApp As Outlook.Application = New Outlook.Application
        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
        Dim rootFolder As Outlook.Folder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent
        Dim testFolder As Outlook.Folder = Nothing

        For Each folder In rootFolder.Folders
            If folder.Name = "Laptop Refresh" Then
                testFolder = folder
                Exit For
            End If
        Next

        If testFolder IsNot Nothing Then
            ' Attach the Items.ItemAdd event handler to the test folder's items collection
            itemsEvents = testFolder.Items
            AddHandler itemsEvents.ItemAdd, AddressOf Items_ItemAdd
        End If
    End Sub

    Private Function FindTestFolder() As Outlook.Folder
        Dim outlookApp As Outlook.Application = New Outlook.Application
        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
        Dim rootFolder As Outlook.Folder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent

        If rootFolder.Folders IsNot Nothing Then
            For Each folder In rootFolder.Folders
                If folder.Name = "Laptop Refresh" Then
                    Return folder
                End If
            Next
        End If

        Return Nothing
    End Function

    Private Function ExtractPID(subject As String) As String
        Dim regex As New Regex("PID:(\d+)")
        Dim match As Match = regex.Match(subject)
        If match.Success Then
            Return match.Groups(1).Value
        Else
            Return String.Empty
        End If
    End Function

    Private Sub Items_ItemAdd(ByVal Item As Object)
        ' Cast the item to a MailItem (assumes all items are MailItems; adjust as needed)
        Dim newMail As Outlook.MailItem = TryCast(Item, Outlook.MailItem)

        If newMail IsNot Nothing Then
            Dim conversationId As String = newMail.ConversationID
            Dim pid As String = ExtractPID(newMail.Subject) ' Extract PID from the subject line
            Dim currentUser As String = System.Environment.UserName
            Dim filePath As String = "\\R0327234\MRIT\UserIDs.txt" ' Your file path here
            Dim lineContainsPIDAndUser As Boolean = False

            ' Check if the PID and username are in the text file
            If File.Exists(filePath) Then
                Dim lines As String() = File.ReadAllLines(filePath)
                For Each line As String In lines
                    If line.Contains($"ID:{pid}") AndAlso line.Contains(currentUser) Then
                        lineContainsPIDAndUser = True
                        Exit For
                    End If
                Next
            End If
            ' Check if this conversation ID is already in any of your lists
            ' For example, checking in one of the added conversations list
            If addedConversationsList1.Contains(conversationId) Then
                ' If yes, find the DataGridView where this conversation is displayed
                ' and update it. This is a simplified example; you'll need to adapt it
                ' to your specific logic for determining which DataGridView to update.
                If lineContainsPIDAndUser Then
                    Me.Invoke(Sub() UpdateDataGridViewForConversation(conversationId, newMail, DataGridView1))
                Else
                    Me.Invoke(Sub() UpdateDataGridViewForConversation(conversationId, newMail, DataGridView2))
                    Me.Invoke(Sub() UpdateDataGridViewForConversation(conversationId, newMail, DataGridView3))
                End If

            ElseIf Not addedConversationsList1.Contains(conversationId) Then
                ' If the conversation is not present, decide where it needs to be added
                ' This example assumes it goes into DataGridView1, but your logic may vary
                If lineContainsPIDAndUser Then
                    Me.Invoke(Sub() AddNewEmailToDataGridView(newMail, DataGridView1))
                Else
                    Me.Invoke(Sub() AddNewEmailToDataGridView(newMail, DataGridView2))
                    Me.Invoke(Sub() AddNewEmailToDataGridView(newMail, DataGridView3))
                End If
            End If
            ' Repeat the checks for addedConversationsList2 and addedConversationsList3 as needed
        End If
    End Sub

    Private Sub AddNewEmailToDataGridView(ByVal newMail As Outlook.MailItem, ByVal dgv As DataGridView)
        ' Generate display text for the new email
        Dim displayText As String = $"{newMail.Subject} - {newMail.ReceivedTime.ToString("g")}"
        Dim conversationId As String = newMail.ConversationID
        Dim subjectEID As String = ExtractEID(newMail.Subject) ' Assuming ExtractEID works as intended

        ' Check if the conversation ID is not already in the added list to avoid duplicates
        If Not addedConversationsList1.Contains(conversationId) Then
            ' Insert a new row at the top of the DataGridView
            dgv.Rows.Insert(0) ' This creates a new row at the top
            Dim newRow As DataGridViewRow = dgv.Rows(0)
            newRow.Cells(0).Value = displayText ' Update this row with the email's display text
            newRow.Tag = conversationId ' Optionally use Tag to store the conversation ID

            ' Here, set the Row Header's value to EID
            newRow.HeaderCell.Value = subjectEID

            ' Update the dictionary to include this new mail item
            If Not emailIdToMailItem.ContainsKey(displayText) Then
                emailIdToMailItem.Add(displayText, newMail)
            End If

            ' Mark as unread if the new mail is unread
            If newMail.UnRead Then
                newRow.DefaultCellStyle.ForeColor = Color.Blue ' Indicating unread
            End If

            ' Update your tracking structures to include this new conversation
            addedConversationsList1.Add(conversationId)
            dataGridOriginalItems1.Insert(0, (displayText, newMail.Subject)) ' Keep track of original items
        End If
    End Sub


    Private Sub UpdateDataGridViewForConversation(ByVal conversationId As String, ByVal newMail As Outlook.MailItem, ByVal dgv As DataGridView)
        ' Commit any pending edits
        If dgv.IsCurrentCellInEditMode Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If

        Dim displayText As String = $"{newMail.Subject} - {newMail.ReceivedTime.ToString("g")}"
        Dim foundRowIndex As Integer = -1

        Dim subjectEID As String = ExtractEID(newMail.Subject)
        Dim colorCode As String = FindColorByEID(subjectEID, "\\R0327234\MRIT\EmailIDs.txt")
        Dim color As Color = If(String.IsNullOrEmpty(colorCode), Color.White, ColorTranslator.FromHtml(colorCode))

        For Each row As DataGridViewRow In dgv.Rows
            Dim mailItem As Outlook.MailItem = TryCast(emailIdToMailItem(row.Cells(0).Value), Outlook.MailItem)
            If mailItem IsNot Nothing AndAlso mailItem.ConversationID = conversationId Then
                foundRowIndex = row.Index
                Exit For
            End If
        Next

        If foundRowIndex >= 0 Then
            dgv.SuspendLayout()

            ' Update the found row directly instead of removing and re-adding
            Dim updateRow As DataGridViewRow = dgv.Rows(foundRowIndex)
            updateRow.Cells(0).Value = displayText
            updateRow.HeaderCell.Value = subjectEID
            updateRow.DefaultCellStyle.BackColor = color

            ' Preserve the mapping in emailIdToMailItem
            emailIdToMailItem(displayText) = newMail

            If newMail.UnRead Then
                updateRow.DefaultCellStyle.ForeColor = Color.Blue
            End If

            ' Move the updated row to the top
            dgv.Rows.RemoveAt(foundRowIndex)
            dgv.Rows.Insert(0, updateRow)

            dgv.ResumeLayout()
        End If
    End Sub







    Private Sub LoadEmails()
        ClearDataStructures()
        Dim testFolder As Outlook.Folder = FindTestFolder()

        If testFolder Is Nothing Then
            MessageBox.Show("The 'test' folder could not be found.")
            Return
        End If

        Try
            AddItemsToDataGridViews(testFolder)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub ClearDataStructures()
        emailIdToMailItem.Clear()
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        DataGridView3.Rows.Clear()
        addedConversationsList1.Clear()
        addedConversationsList2.Clear()
        addedConversationsList3.Clear()
        dataGridOriginalItems1.Clear()
        dataGridOriginalItems2.Clear()
        dataGridOriginalItems3.Clear()
    End Sub

    Private Sub AddItemsToDataGridViews(testFolder As Outlook.Folder)

        ClearDataStructures()
        Dim fileName As String = "UserIDs.txt"
        Dim filePath As String = Path.Combine("\\R0327234\MRIT", fileName)
        Dim fileName2 As String = "EmailIDs.txt"
        Dim filePath2 As String = Path.Combine("\\R0327234\MRIT", fileName2)
        Dim lineCount = File.ReadAllLines(filePath).Length


        Dim UserID As String
        Dim userIDList As String() = File.ReadAllLines(filePath)
        Dim foundId As String = String.Empty

        Dim currentUser As String = System.Environment.UserName

        Try
            Dim emailIdLines As String() = File.ReadAllLines(filePath2)
            For Each line As String In emailIdLines
                Dim parts As String() = line.Split(","c)
                If parts.Length >= 3 Then
                    ' Assuming format is "name, EID:#, color"
                    Dim eid As String = parts(1).Trim()
                    Dim name As String = parts(0).Trim()
                    Dim color As String = parts(2).Trim()
                    If Not emailIdInfo.ContainsKey(eid) Then
                        emailIdInfo.Add(eid, (name, color))
                    End If
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Error reading EmailIDs.txt: " & ex.Message)
        End Try

        Try
            UserID = File.ReadAllText(filePath)
            If Not UserID.Contains(currentUser) Then
                File.AppendAllText(filePath, Environment.NewLine + currentUser + " - ID:" + lineCount.ToString())
            End If
        Catch ex As Exception
            MessageBox.Show("Error reading the file: " & ex.Message)
        End Try

        For Each item As String In userIDList
            If item.StartsWith(currentUser) Then
                Dim parts As String() = item.Split(" "c)
                foundId = parts(parts.Length - 1)
            End If
        Next

        For Each mailItem In testFolder.Items
            Dim item As Outlook.MailItem = TryCast(mailItem, Outlook.MailItem)
            If item IsNot Nothing Then
                Dim conversationId As String = item.ConversationID
                Dim displayText As String = $"{item.Subject} - {item.ReceivedTime.ToString("g")}"
                Dim searchText As String = item.Subject
                Dim subjectId As String = ExtractSubjectId(item.Subject)

                Dim subjectEID As String = ExtractEID(item.Subject)

                AddItemToDataGridView(DataGridView3, displayText, searchText, subjectEID, item.UnRead, conversationId, item, addedConversationsList3, dataGridOriginalItems3)

                If item.Subject.Contains(foundId) Then
                    AddItemToDataGridView(DataGridView1, displayText, searchText, subjectEID, item.UnRead, conversationId, item, addedConversationsList1, dataGridOriginalItems1)
                Else
                    AddItemToDataGridView(DataGridView2, displayText, searchText, subjectEID, item.UnRead, conversationId, item, addedConversationsList2, dataGridOriginalItems2)
                End If

            End If

        Next

        ApplyColorsToRows()
    End Sub

    Private Sub AddItemToDataGridView(dgv As DataGridView, displayText As String, searchText As String, subjectEID As String, unread As Boolean, conversationId As String, mailItem As Outlook.MailItem, addedConversationsList As HashSet(Of String), dataGridOriginalItems As List(Of (String, String)))
        If Not addedConversationsList.Contains(conversationId) Then
            Dim rowIndex = dgv.Rows.Add(displayText)
            dgv.Rows(rowIndex).HeaderCell.Value = subjectEID
            dataGridOriginalItems.Add((displayText, searchText))
            emailIdToMailItem(displayText) = mailItem
            addedConversationsList.Add(conversationId)
            If unread Then
                dgv.Rows(rowIndex).DefaultCellStyle.ForeColor = Color.Blue ' Indicate unread
            End If
        End If
    End Sub


    ' Call this method immediately after all rows have been added
    Private Sub ApplyColorsToRows()
        ' Assuming filePath2 is correctly defined elsewhere
        Dim filePath2 As String = Path.Combine("\\R0327234\MRIT", "EmailIDs.txt")
        ' Apply colors for each DataGridView separately
        ApplyColorByEID(DataGridView1, filePath2)
        ApplyColorByEID(DataGridView2, filePath2)
        ApplyColorByEID(DataGridView3, filePath2)
    End Sub

    Private Sub ApplyColorByEID(dgv As DataGridView, filePath As String)
        For Each row As DataGridViewRow In dgv.Rows
            ' Assuming the EID can be derived from the row's data directly
            ' Adjust this logic based on how you can extract the EID or related data from the row
            Dim eid As String = ExtractEID(row.Cells(0).Value.ToString())
            Dim hexColorCode As String = FindColorByEID(eid, filePath)
            If Not String.IsNullOrEmpty(hexColorCode) Then
                Try
                    row.DefaultCellStyle.BackColor = ColorTranslator.FromHtml(hexColorCode)
                Catch ex As Exception
                    ' Log or handle the exception as needed
                End Try
            End If
        Next
    End Sub

    Private Function FindColorByEID(eid As String, filePath As String) As String
        Try
            Dim lines As String() = File.ReadAllLines(filePath)
            For Each line In lines
                Dim parts As String() = line.Split(","c)
                If parts.Length >= 3 Then
                    ' Remove "EID:" from the extracted EID part and trim spaces
                    Dim currentEID As String = parts(1).Replace("EID:", "").Trim()
                    If String.Equals(currentEID, eid, StringComparison.OrdinalIgnoreCase) Then
                        Return parts(2).Trim() ' Return the color code
                    End If
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Error reading EmailIDs.txt: " & ex.Message)
        End Try
        Return Nothing ' Return Nothing if no match is found or an error occurs
    End Function

    Private Sub DataGridView_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView3.SelectionChanged, DataGridView1.SelectionChanged, DataGridView2.SelectionChanged
        Dim dataGridView As DataGridView = TryCast(sender, DataGridView)
        If dataGridView IsNot Nothing Then
            DisplayEmailContent(dataGridView)
        End If
    End Sub

    Private Sub DisplayEmailContent(dataGridView As DataGridView)
        Dim richTextBox As RichTextBox = Nothing

        ' Determine which RichTextBox corresponds to the DataGridView
        Select Case dataGridView.Name
            Case "DataGridView1"
                richTextBox = RichTextBox1
            Case "DataGridView2"
                richTextBox = RichTextBox2
            Case "DataGridView3"
                richTextBox = RichTextBox3
        End Select

        If richTextBox IsNot Nothing AndAlso dataGridView.CurrentCell IsNot Nothing Then
            Dim selectedItemText As String = TryCast(dataGridView.CurrentCell.Value, String)
            If Not String.IsNullOrEmpty(selectedItemText) AndAlso emailIdToMailItem.ContainsKey(selectedItemText) Then
                Dim mailItem As Outlook.MailItem = emailIdToMailItem(selectedItemText)
                richTextBox.Text = mailItem.Body ' Or mailItem.HTMLBody as per your requirement
            Else
                richTextBox.Clear()
            End If
        End If
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        ClearDataStructures()
        LoadEmails()

        Dim searchQuery As String = TextBoxSearch.Text.Trim().ToLower() ' Use Trim() to handle spaces
        'FilterDataGridView(searchQuery, DataGridView1, dataGridOriginalItems1)

        Dim searchQuery2 As String = TextBoxSearch2.Text.Trim().ToLower() ' Use Trim() to handle spaces
        'FilterDataGridView(searchQuery2, DataGridView2, dataGridOriginalItems2)

        Dim searchQuery3 As String = TextBoxSearch3.Text.Trim().ToLower() ' Use Trim() to handle spaces
        'FilterDataGridView(searchQuery3, DataGridView3, dataGridOriginalItems3)
    End Sub

    Private Sub TextBoxSearch_KeyUp_All(sender As Object, e As KeyEventArgs) Handles TextBoxSearch3.KeyUp, TextBoxSearch2.KeyUp, TextBoxSearch.KeyUp
        Dim searchTextBox As TextBox = TryCast(sender, TextBox)
        If searchTextBox Is Nothing Then Return

        Dim searchQuery As String = searchTextBox.Text.Trim().ToLower() ' Use Trim() to handle spaces
        Dim dataGridView As DataGridView = Nothing
        Dim dataGridOriginalItems As List(Of (String, String)) = Nothing

        ' Determine which DataGridView and data source list correspond to the TextBox
        Select Case searchTextBox.Name
            Case "TextBoxSearch"
                dataGridView = DataGridView1
                dataGridOriginalItems = dataGridOriginalItems1
            Case "TextBoxSearch2"
                dataGridView = DataGridView2
                dataGridOriginalItems = dataGridOriginalItems2
            Case "TextBoxSearch3"
                dataGridView = DataGridView3
                dataGridOriginalItems = dataGridOriginalItems3
        End Select

        ' Filter the corresponding DataGridView
        If dataGridView IsNot Nothing AndAlso dataGridOriginalItems IsNot Nothing Then
            FilterDataGridView(searchQuery, dataGridView, dataGridOriginalItems)
        End If
    End Sub

    Private Sub FilterDataGridView(searchQuery As String, data As DataGridView, items As List(Of (String, String)))
        data.Rows.Clear() ' Clear existing rows

        ' Convert search query to lower case for a case-insensitive search
        searchQuery = searchQuery.ToLower()

        Dim filePath2 As String = Path.Combine("\\R0327234\MRIT", "EmailIDs.txt") ' Ensure this path is correct

        ' Add rows back that match the search query
        For Each itemTuple In items
            If itemTuple.Item2.ToLower().Contains(searchQuery) Then
                Dim rowIndex As Integer = data.Rows.Add(itemTuple.Item1)

                ' Extract EID from the subject (or wherever it's stored)
                Dim eid As String = ExtractEID(itemTuple.Item2) ' Ensure this extracts correctly for your data
                data.Rows(rowIndex).HeaderCell.Value = eid ' Set the row header value to the EID

                ' Find the color by EID
                Dim hexColorCode As String = FindColorByEID(eid, filePath2)
                Dim colorToSet As Color = SystemColors.Window ' Default color
                If Not String.IsNullOrEmpty(hexColorCode) Then
                    Try
                        ' Convert hex code to Color
                        colorToSet = ColorTranslator.FromHtml(hexColorCode)
                    Catch ex As Exception
                        ' Handle invalid color code if necessary
                    End Try
                End If
                ' Apply the color
                data.Rows(rowIndex).DefaultCellStyle.BackColor = colorToSet
            End If
        Next
    End Sub

    Private Function ExtractSubjectId(subject As String) As String
        ' Use regular expressions to find a pattern that matches "ID:#"
        Dim regex As New Regex("PID:(\d+)")
        Dim match As Match = regex.Match(subject)
        If match.Success Then
            ' If a match is found, return the ID value
            Return match.Groups(1).Value ' This returns the whole "ID:#" string, or use match.Groups(1).Value to return just the number
        Else
            Return "Unknown" ' Fallback if no ID is found
        End If
    End Function

    Private Function ExtractEID(subject As String) As String
        ' Use regular expressions to find a pattern that matches "ID:#"
        Dim regex As New Regex("EID:(\d+)")
        Dim match As Match = regex.Match(subject)
        If match.Success Then
            ' If a match is found, return the ID value
            Return match.Groups(1).Value ' This returns the whole "ID:#" string, or use match.Groups(1).Value to return just the number
        Else
            Return "Unknown" ' Fallback if no ID is found
        End If
    End Function

    Private Function ExtractColor(subject As String) As String
        Dim parts As String() = subject.Split(","c)
        If parts.Length >= 3 Then
            ' Trim to remove any leading/trailing spaces
            Return parts(2).Trim()
        Else
            Return "Color not found"
        End If
    End Function

    Private Sub DataGridViews_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridView1.CellPainting, DataGridView3.CellPainting, DataGridView2.CellPainting
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            ' Fill the background
            Using backColorBrush As New SolidBrush(e.CellStyle.BackColor)
                e.Graphics.FillRectangle(backColorBrush, e.CellBounds)
            End Using

            ' Paint the content. Ensure the text uses the original forecolor.
            Dim formattedValueText As String = If(e.FormattedValue Is Nothing, String.Empty, e.FormattedValue.ToString())
            Using textBrush As New SolidBrush(e.CellStyle.ForeColor) ' Use the original ForeColor
                e.Graphics.DrawString(formattedValueText, e.CellStyle.Font, textBrush, e.CellBounds.X + 2, e.CellBounds.Y + 2)
            End Using

            ' Draw the custom border for selected cells
            If CType(sender, DataGridView).Rows(e.RowIndex).Cells(e.ColumnIndex).Selected Then
                Using borderPen As New Pen(Color.Red, 2) ' Customize the border color and thickness here
                    e.Graphics.DrawRectangle(borderPen, e.CellBounds.Left + 1, e.CellBounds.Top + 1, e.CellBounds.Width - 4, e.CellBounds.Height - 4)
                End Using
            End If

            ' Prevent the default painting process from overwriting our custom painting.
            e.Handled = True
        End If
    End Sub


    Private Sub DataGridView_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick, DataGridView1.CellClick, DataGridView2.CellClick
        Dim dgv As DataGridView = CType(sender, DataGridView)
        If e.RowIndex >= 0 Then
            Dim selectedItemText As String = dgv.Rows(e.RowIndex).Cells(0).Value.ToString()
            If selectedItemText IsNot Nothing AndAlso emailIdToMailItem.ContainsKey(selectedItemText) Then
                Dim mailItem As Outlook.MailItem = emailIdToMailItem(selectedItemText)
                Dim conversationId As String = mailItem.ConversationID

                ' Mark all emails in the conversation as read
                For Each item As KeyValuePair(Of String, Outlook.MailItem) In emailIdToMailItem
                    If item.Value.ConversationID = conversationId AndAlso item.Value.UnRead Then
                        item.Value.UnRead = False
                        item.Value.Save()
                    End If
                Next

            End If
        End If
    End Sub

    Private activeDataGridView As DataGridView

    Private Sub DataGridView_MouseDown(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDown, DataGridView3.MouseDown, DataGridView2.MouseDown
        activeDataGridView = DirectCast(sender, DataGridView)
        Dim dataGridView As DataGridView = DirectCast(sender, DataGridView)
        If e.Button = MouseButtons.Right Then
            Dim hitTestInfo As DataGridView.HitTestInfo = dataGridView.HitTest(e.X, e.Y)
            If hitTestInfo.Type = DataGridViewHitTestType.Cell Then
                dataGridView.ClearSelection()
                dataGridView.Rows(hitTestInfo.RowIndex).Cells(hitTestInfo.ColumnIndex).Selected = True
            End If
        End If
    End Sub

    Private Sub UpdateColorInFile(eid As String, newColor As String, filePath As String)
        Try
            If File.Exists(filePath) Then
                Dim lines As List(Of String) = File.ReadAllLines(filePath).ToList()
                Dim fileUpdated As Boolean = False

                ' Use a regular expression to remove all non-numeric characters from the EID
                Dim numericEID As String = Regex.Replace(eid, "[^\d]", String.Empty)

                For i As Integer = 0 To lines.Count - 1
                    ' Split the line into parts
                    Dim parts As String() = lines(i).Split(","c)
                    If parts.Length >= 3 Then
                        ' Use a regular expression to remove all non-numeric characters from the EID in the file
                        Dim numericFileEID As String = Regex.Replace(parts(1).Trim(), "[^\d]", String.Empty)

                        Console.WriteLine(numericFileEID.ToString + " ," + numericEID.ToString + " ," + String.Equals(numericFileEID, numericEID, StringComparison.OrdinalIgnoreCase).ToString)

                        ' Ensure the EID matches exactly after removing non-numeric characters
                        If String.Equals(numericFileEID, numericEID, StringComparison.OrdinalIgnoreCase) Then
                            ' Replace the color part (3rd part) with the new color
                            parts(2) = newColor
                            lines(i) = String.Join(",", parts)
                            fileUpdated = True
                            Exit For
                        End If
                    End If
                Next

                If fileUpdated Then
                    ' Only write back to the file if an update was made to avoid unnecessary file operations
                    File.WriteAllLines(filePath, lines)
                End If
            Else
                MessageBox.Show("File not found: " & filePath)
            End If
        Catch ex As Exception
            MessageBox.Show("Error updating the file: " & ex.Message)
        End Try
    End Sub

    Private Sub UpdateStatusAndColor_Click(sender As Object, e As EventArgs) Handles TestToolStripMenuItem.Click, ScheduledToolStripMenuItem.Click, InProgressToolStripMenuItem.Click, NoReplyUnavailableToolStripMenuItem.Click, CustomToolStripMenuItem.Click, NoneToolStripMenuItem.Click
        If activeDataGridView IsNot Nothing AndAlso activeDataGridView.SelectedCells.Count > 0 Then
            Dim menuItem As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
            Dim hexColor As String = menuItem.Tag.ToString()

            ' Handle custom color picker separately
            If hexColor = "Custom" Then
                Using colorDialog As New ColorDialog()
                    If colorDialog.ShowDialog() = DialogResult.OK Then
                        ' Convert the selected color to a hex code
                        hexColor = "#" & colorDialog.Color.R.ToString("X2") & colorDialog.Color.G.ToString("X2") & colorDialog.Color.B.ToString("X2")
                    Else
                        ' If the user cancels the color dialog, exit the method without making changes
                        Return
                    End If
                End Using
            End If

            Dim selectedCell As DataGridViewCell = activeDataGridView.SelectedCells(0)
            Dim rowIndex As Integer = selectedCell.RowIndex
            Dim eid As String = activeDataGridView.Rows(rowIndex).HeaderCell.Value.ToString()

            ' Update the DataGridView cell color
            UpdateDataGridViewCellColor(activeDataGridView, rowIndex, hexColor)

            ' Update the EmailIDs.txt file with the new hex code for the matching EID
            UpdateColorInFile(eid, hexColor, "\\R0327234\MRIT\EmailIDs.txt")

            ' Check and update other DataGridViews
            UpdateOtherDataGridViewsColor(eid, hexColor, activeDataGridView)
        Else
            MessageBox.Show("Please select a cell first.")
        End If

    End Sub

    Private Sub UpdateOtherDataGridViewsColor(eid As String, hexColor As String, ByVal excludeDataGridView As DataGridView)
        Dim dataGrids As DataGridView() = {DataGridView1, DataGridView2, DataGridView3}

        For Each dgv In dataGrids
            If Not dgv Is excludeDataGridView Then ' Skip the active DataGridView
                For Each row As DataGridViewRow In dgv.Rows
                    If row.HeaderCell.Value IsNot Nothing AndAlso row.HeaderCell.Value.ToString() = eid Then
                        row.DefaultCellStyle.BackColor = ColorTranslator.FromHtml(hexColor)
                        ' Optionally break here if EID is unique
                    End If
                Next
            End If
        Next
    End Sub

    Private Sub UpdateDataGridViewCellColor(ByRef dgv As DataGridView, rowIndex As Integer, hexColor As String)
        If dgv IsNot Nothing AndAlso rowIndex >= 0 Then
            dgv.Rows(rowIndex).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(hexColor)
        End If
    End Sub

    Private Sub ConsolidatedReplyClick(sender As Object, e As EventArgs) Handles btnReply1.Click, Button2.Click, Button1.Click
        ' Determine the active DataGridView based on the selected tab
        Dim activeDataGridView As DataGridView = GetActiveDataGridViewFromSelectedTab()
        If activeDataGridView Is Nothing OrElse activeDataGridView.CurrentCell Is Nothing Then
            MessageBox.Show("Please select an email to reply to.")
            Return
        End If

        ' Proceed with the reply using the selected item in the active DataGridView
        Dim selectedItemText As String = TryCast(activeDataGridView.CurrentCell.Value, String)
        If String.IsNullOrEmpty(selectedItemText) OrElse Not emailIdToMailItem.ContainsKey(selectedItemText) Then
            MessageBox.Show("Please select an email to reply to.")
            Return
        End If

        Dim mailItem As Outlook.MailItem = emailIdToMailItem(selectedItemText)

        ' Queue the operation to the ThreadPool
        Threading.ThreadPool.QueueUserWorkItem(AddressOf ReplyToEmail, mailItem)
    End Sub

    Private Sub ReplyToEmail(state As Object)
        Try
            Dim mailItem As Outlook.MailItem = TryCast(state, Outlook.MailItem)
            If mailItem Is Nothing Then Return

            Try
                ' ReplyAll to the mail item passed as the state argument
                Dim reply As Outlook.MailItem = mailItem.ReplyAll()
                reply.Display() ' This opens the reply window without freezing the UI
            Catch ex As Exception
                ' Since you're on a background thread, consider logging the exception or invoking a UI thread operation to display it
                MessageBox.Show($"An error occurred while trying to reply: {ex.Message}")
            End Try
        Catch ex As COMException
            MessageBox.Show($"COMException caught: {ex.Message}, HRESULT: {ex.ErrorCode}")
            ' Log the error details or take appropriate action
        End Try
    End Sub


    Private Sub DisplayReplyMailAsync(mailItem As Outlook.MailItem)
        If mailItem IsNot Nothing Then
            Try
                ' Create and display the reply mail item
                Dim replyMail As Outlook.MailItem = mailItem.ReplyAll()
                replyMail.Display(True)
            Catch ex As Exception
                ' Exception handling logic here, note: MessageBox.Show must be called from the UI thread
                Console.WriteLine($"Error displaying reply mail: {ex.Message}")
            End Try
        End If
    End Sub





    Private Function GetActiveDataGridViewFromSelectedTab() As DataGridView
        ' Determine the currently selected tab
        Dim selectedTabIndex As Integer = TabControl1.SelectedIndex
        Dim activeDataGridView As DataGridView = Nothing

        ' Match the selected tab index to the corresponding DataGridView
        ' Adjust these case statements based on your actual setup
        Select Case selectedTabIndex
            Case 0 ' Assuming the first tab contains DataGridView1
                activeDataGridView = DataGridView1
            Case 1 ' Assuming the second tab contains DataGridView2
                activeDataGridView = DataGridView2
            Case 2 ' Assuming the third tab contains DataGridView3
                activeDataGridView = DataGridView3
        End Select

        Return activeDataGridView
    End Function


End Class