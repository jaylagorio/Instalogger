''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' File: Report.vb
''' Author: Jay Lagorio
''' Date Changed: 31OCT2020
''' Purpose: This class queries the database using its own DBConnection (and gets the Connection
''' String from the passed Configuration object) and synthesizes an HTML-formatted email, with
''' embedded images where appropriate, and sends the report to the configured email address using
'''the configured SMTP server information.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports System.Collections.ObjectModel
Imports Instalogger.Database
Imports EASendMail
Imports System.Net
Imports SixLabors.ImageSharp
Imports SixLabors.ImageSharp.Processing
Imports System.IO

''' <summary>
''' This class queries the database using its own DBConnection (and gets the Connection String from the passed Configuration object)
''' and synthesizes an HTML-formatted email, with embedded images where appropriate, and sends the report to the configured email
''' address using the configured SMTP server information.
''' </summary>
Public Class Report
    ' This is the maximum width or height to scale the display of any profile pictures in change tracking
    Private Const MAXIMUM_RESIZED_DISPLAY_IMAGE_HEIGHT_OR_WIDTH As Integer = 250

    ' This is the maximum width or height to scale the actual size of profile pictures in headers
    Private Const MAXIMUM_PHOTO_THUMBNAIL_HEIGHT_OR_WIDTH As Integer = 125

    ' Constants used to construct the HTML body of the report
    Private Const HTML_HEAD_AND_BODY_OPENER As String = "<html><head><style>.logcell {background-color: #dddddd;margin-bottom: 3px;margin-left: 15px;}.deleted {text-decoration:line-through;color: red;}.inserted {color: green;}.modified {color: blue;}.unchanged {color: black;}.timestamp {font-size: 8pt;}</style></head><body><table border=0 width=""50%""><tr><td><center><img src=""cid:instalogger.jpg""></center><p/>"
    Private Const BODY_AND_HTML_CLOSER As String = "</td></tr></table></body></html>"
    Private Const PARAGRAPH_BREAK As String = "<p/>"
    Private Const START_SPAN_DELETED As String = "<span class=""deleted"">"
    Private Const START_SPAN_INSERTED As String = "<span class=""inserted"">"
    Private Const START_SPAN_MODIFIED As String = "<span class=""modified"">"
    Private Const START_SPAN_UNCHANGED As String = "<span class=""unchanged"">"
    Private Const CLOSE_SPAN As String = "</span>"
    Private Const START_SPAN_DELETED_WITH_AHREF As String = "<span class=""deleted""><a class=""deleted"" href="""
    Private Const START_SPAN_INSERTED_WITH_AHREF As String = "<span class=""inserted""><a class=""inserted"" href="""
    Private Const START_SPAN_MODIFIED_WITH_AHREF As String = "<span class=""modified""><a class=""modified"" href="""
    Private Const START_SPAN_UNCHANGED_WITH_AHREF As String = "<span class=""unchanged""><a class=""unchanged"" href="""
    Private Const CLOSE_AHREF_AND_SPAN As String = "</a></span>"
    Private Const START_LOGCELL_DIV As String = "<div class=""logcell"">"
    Private Const CLOSE_DIV As String = "</div>"
    Private Const START_TIMESTAMP_DIV_WITH_ITALICS As String = "<div class=""timestamp""><i>"
    Private Const CLOSE_TIMESTAMP_ITALICS As String = "</i> - "

    ' This structure is used to store the Content ID and image data used to attached changed profile photos
    Public Structure EmailImageAttachment
        Dim CID As String       ' Content ID for embedding graphics in an email
        Dim Picture() As Byte   ' The raw bytes representing the photo downloaded from Instagram
    End Structure

    ' Local copy of the Configuration object, which we need to get SMTP settings
    Private pConfiguration As Configuration

    ''' <summary>
    ''' Constructor for the Report object.
    ''' </summary>
    ''' <param name="Configuration">An instance of a Configuration object which contains the necessary SMTP settings
    ''' to send reports on to their destinations</param>
    Sub New(ByRef Configuration As Configuration)
        pConfiguration = Configuration
    End Sub


    ''' <summary>
    ''' Produces an HTML text report, with embedded images if necessary, from data in the database regarding changes
    ''' to profiles in the last day. The report is emailed to the To address specified in the Configuration object
    ''' and is from the From address configured in the Configuration object.
    ''' </summary>
    ''' <returns>Returns True if the email was sent to the SMTP server successfully, False otherwise.</returns>
    Public Function ReportChangedProfiles() As Boolean
        ' Initialize the database, set aside a Collection for EmailImageAttachments, and prepare
        ' a WebClient to download changed profile photos
        Dim Database As New Database(pConfiguration.DatabaseString)
        Dim ImageAttachments As New Collection(Of EmailImageAttachment)
        Dim WebClient As New WebClient

        ' Begin the report with the HTML, HEAD, and BODY opener tags and get the changed profiles
        Dim HTMLReport As String = HTML_HEAD_AND_BODY_OPENER
        Dim ChangedIGIDs() As String = Database.GetChangedProfilesByIGID()
        If Not ChangedIGIDs Is Nothing Then
            ' Prepare a report section for each changed user
            For i = 0 To ChangedIGIDs.Count - 1
                ' We track changes via Instagram ID because usernames can change. Get the latest username for that ID and
                ' begin the section of the report with the username and link to that profile.
                Dim Username As String = Database.GetUsernameByIGID(ChangedIGIDs(i))

                ' Retrieve, shrink, and embed the user's profile picture above their section
                Dim ProfileAttachment As New EmailImageAttachment
                ProfileAttachment.CID = "profile_" & Username & ".jpg"
                ProfileAttachment.Picture = Database.GetProfilePicData(ChangedIGIDs(i))
                If Not ProfileAttachment.Picture Is Nothing Then
                    Call ResizeImageToThumbnail(ProfileAttachment.Picture)
                    Call ImageAttachments.Add(ProfileAttachment)
                End If

                HTMLReport &= PARAGRAPH_BREAK & "<div>&nbsp;</div><center><a href=""https://instagram.com/" & Username & """><img src=""cid:" & ProfileAttachment.CID & """><br/><b>" & Username & "</b></a>:</center><p/>"

                ' Check to see if any users had changed fields. If so, the returned ChangeRecord passed to the 
                ' calculation functions will have entries in their ChangedField() array, if not the array will
                ' be empty and no HTML output will be returned.
                HTMLReport &= CalculateLogCell(Database.GetChangedBiographiesByIGID(ChangedIGIDs(i)))
                HTMLReport &= CalculateExternalURLCell(Database.GetChangedExternalURLsByIGID(ChangedIGIDs(i)))
                HTMLReport &= CalculateLogCell(Database.GetChangedFullNamesByIGID(ChangedIGIDs(i)))
                HTMLReport &= CalculateLogCell(Database.GetChangedLocationsByIGID(ChangedIGIDs(i)))
                HTMLReport &= CalculateLogCell(Database.GetChangedUsernamesByIGID(ChangedIGIDs(i)))

                ' If there are any changed profile pictures in the last day, download them, get the Filename portion
                ' of the URL to use as the Content ID, and generate HTML content that will embed them in-line the email
                ' when they're attached to the email.
                Dim ProfilePics As ChangeRecord = Database.GetChangedProfilePicsByIGID(ChangedIGIDs(i))
                If Not ProfilePics.ChangedField Is Nothing Then
                    If ProfilePics.ChangedField.Count > 0 Then
                        If Not (ProfilePics.ChangedField.Count = 1 And ProfilePics.ChangedField(0) = "") Then
                            For j = 0 To ProfilePics.ChangedField.Count - 1
                                Dim NewImageAttachment As New EmailImageAttachment
                                NewImageAttachment.CID = System.IO.Path.GetFileName((New Uri(ProfilePics.ChangedField(j))).LocalPath)
                                'NewImageAttachment.Picture = WebClient.DownloadData(ProfilePics.ChangedField(j))
                                NewImageAttachment.Picture = Database.GetProfilePicData(ProfilePics.IGID)
                                Call ImageAttachments.Add(NewImageAttachment)
                            Next
                        End If

                        ' Generate the HTML based on the picture change data and the attachments
                        HTMLReport &= CalculateChangedImage(ProfilePics, ImageAttachments)
                    End If
                End If
            Next
        Else
            HTMLReport &= "Instalogger didn't detect any changes in Followed profiles."
        End If
        HTMLReport &= BODY_AND_HTML_CLOSER

        Return SendEmail("Instalogger: Profile Changes", HTMLReport, ImageAttachments)
    End Function


    ''' <summary>
    ''' Takes a history of an image field in the profile and generates a series of cells that reflect the changes
    ''' made over time to that image field.
    ''' </summary>
    ''' <param name="ChangedData">A Database.ChangedData structure representing changes over time for a field and the times those changes were made</param>
    ''' <returns>Returns HTML content containing <divs>, <spans>, and <img> tags that style the prior and current content of the field. If there was no data passed an empty string will be returned.</returns>
    Private Function CalculateChangedImage(ByRef ChangedData As ChangeRecord, ByRef ImageAttachments As Collection(Of EmailImageAttachment)) As String
        ' If there isn't any changed data for this cell, bail and return no HTML content
        Try
            If ChangedData.ChangedField Is Nothing Then Return ""
            If ChangedData.ChangedField.Count = 0 Then Return ""
        Catch ex As Exception
            Return ""
        End Try

        ' If there is only one history record of the field and it's blank, bail and return no HTML content
        If ChangedData.ChangedField.Count = 1 And ChangedData.ChangedField(0) = "" Then Return ""

        ' Start the new HTML content and initialize a DiffPlex engine
        Dim HTMLContent As String
        Dim DiffEngine As New DiffPlex.DiffBuilder.InlineDiffBuilder()

        ' Grab the Filename component of the URL and make that the Content ID, then find the image
        ' in the array to calculate the reduction size (if it needs to be reduced)
        Dim CurrentCID As String = System.IO.Path.GetFileName((New Uri(ChangedData.ChangedField(ChangedData.ChangedField.Count - 1))).LocalPath)
        Dim CurrentCIDSizeLimit As String = ""
        For i = 0 To ImageAttachments.Count - 1
            If ImageAttachments(i).CID = CurrentCID Then
                ' Note this can come out to be an empty string in case the image doesn't need to be resized
                CurrentCIDSizeLimit = CalculateImageDisplaySize(ImageAttachments(i).Picture)
                Exit For
            End If
        Next

        ' Concatinate the <img> tag, the CID, and the size limit in case one needs to be set (plus the timestamp, etc)
        ' for the orignial image
        HTMLContent = START_LOGCELL_DIV & "<img src=""cid:" & CurrentCID & """" & " " & CurrentCIDSizeLimit & " " & ">" & CLOSE_SPAN & START_TIMESTAMP_DIV_WITH_ITALICS & ChangedData.FieldName & CLOSE_TIMESTAMP_ITALICS & ChangedData.ChangeDetected(ChangedData.ChangedField.Count - 1) & CLOSE_DIV & CLOSE_DIV

        For j = ChangedData.ChangedField.Count - 1 To 1 Step -1
            HTMLContent &= START_LOGCELL_DIV
            Dim Model As DiffPlex.DiffBuilder.Model.DiffPaneModel = DiffEngine.BuildDiffModel(ChangedData.ChangedField(j), ChangedData.ChangedField(j - 1))
            For k = 0 To Model.Lines.Count - 1
                Select Case Model.Lines(k).Type
                    Case DiffPlex.DiffBuilder.Model.ChangeType.Deleted
                        HTMLContent &= START_SPAN_DELETED
                    Case DiffPlex.DiffBuilder.Model.ChangeType.Inserted
                        HTMLContent &= START_SPAN_INSERTED
                    Case DiffPlex.DiffBuilder.Model.ChangeType.Modified
                        HTMLContent &= START_SPAN_MODIFIED
                    Case DiffPlex.DiffBuilder.Model.ChangeType.Unchanged
                        HTMLContent &= START_SPAN_UNCHANGED
                End Select

                ' Grab the Filename component of the URL and make that the Content ID, then find the image
                ' in the array to calculate the reduction size (if it needs to be reduced)
                CurrentCID = System.IO.Path.GetFileName((New Uri(ChangedData.ChangedField(k))).LocalPath)
                For i = 0 To ImageAttachments.Count - 1
                    If ImageAttachments(i).CID = CurrentCID Then
                        CurrentCIDSizeLimit = CalculateImageDisplaySize(ImageAttachments(i).Picture)
                        Exit For
                    End If
                Next

                ' Concatinate the <img> tag, the CID, and the size limit in case one needs to be set (plus the timestamp, etc)
                ' for the image in this part of the history
                HTMLContent &= "<img src=""cid:" & CurrentCID & """ " & CurrentCIDSizeLimit & " " & ">" & CLOSE_SPAN
            Next

            ' Add the timestamp, close the <div>, move on to the next one
            HTMLContent &= START_TIMESTAMP_DIV_WITH_ITALICS & ChangedData.FieldName & CLOSE_TIMESTAMP_ITALICS & ChangedData.ChangeDetected(j - 1) & CLOSE_DIV
            HTMLContent &= CLOSE_DIV
        Next

        ' Return the completed change history of this field for this user
        Return HTMLContent
    End Function


    ''' <summary>
    ''' Takes a history of a URL field in the profile and generates a series of cells that reflect the changes
    ''' made over time to that URL.
    ''' </summary>
    ''' <param name="ChangedData">A Database.ChangedData structure representing changes over time for a URL field and the times those changes were made</param>
    ''' <returns>Returns HTML content containing <divs> and <spans> that style the prior and current content of the field. If there was no data passed an empty string will be returned.</returns>
    Private Function CalculateExternalURLCell(ByRef ChangedData As ChangeRecord) As String
        ' If there isn't any changed data for this cell, bail and return no HTML content
        Try
            If ChangedData.ChangedField Is Nothing Then Return ""
            If ChangedData.ChangedField.Count = 0 Then Return ""
        Catch ex As Exception
            Return ""
        End Try

        ' If there is only one history record of the field and it's blank, bail and return no HTML content
        If ChangedData.ChangedField.Count = 1 And ChangedData.ChangedField(0) = "" Then Return ""

        ' Generating HTML here is the same as in CalculateLogCell except there's an extra <a> tag around
        ' the link content. In CSS A tags don't get things like strikethroughs and colors that a <div> or
        ' <span> will give regular text via their class so we have to add the class to the <a> tag too.
        Dim HTMLContent As String = ""
        Dim DiffEngine As New DiffPlex.DiffBuilder.InlineDiffBuilder()

        ' Only include the prior state before any changes if it wasn't empty
        If ChangedData.ChangedField(ChangedData.ChangedField.Count - 1) <> "" Then
            HTMLContent = START_LOGCELL_DIV & ChangedData.ChangedField(ChangedData.ChangedField.Count - 1) & START_TIMESTAMP_DIV_WITH_ITALICS & ChangedData.FieldName & CLOSE_TIMESTAMP_ITALICS & ChangedData.ChangeDetected(ChangedData.ChangedField.Count - 1) & CLOSE_DIV & CLOSE_DIV
        End If

        For j = ChangedData.ChangedField.Count - 1 To 1 Step -1
            HTMLContent &= START_LOGCELL_DIV
            Dim Model As DiffPlex.DiffBuilder.Model.DiffPaneModel = DiffEngine.BuildDiffModel(ChangedData.ChangedField(j), ChangedData.ChangedField(j - 1))
            For k = 0 To Model.Lines.Count - 1
                Select Case Model.Lines(k).Type
                    Case DiffPlex.DiffBuilder.Model.ChangeType.Deleted
                        HTMLContent &= START_SPAN_DELETED_WITH_AHREF
                    Case DiffPlex.DiffBuilder.Model.ChangeType.Inserted
                        HTMLContent &= START_SPAN_INSERTED_WITH_AHREF
                    Case DiffPlex.DiffBuilder.Model.ChangeType.Modified
                        HTMLContent &= START_SPAN_MODIFIED_WITH_AHREF
                    Case DiffPlex.DiffBuilder.Model.ChangeType.Unchanged
                        HTMLContent &= START_SPAN_UNCHANGED_WITH_AHREF
                End Select

                ' Add the URL to the open href attribute, close the attribute and tag start, add the link, and then close the a tag
                HTMLContent &= Model.Lines(k).Text & """>" & Model.Lines(k).Text & CLOSE_AHREF_AND_SPAN
            Next

            ' Add the timestamp the change to the External URL was made
            HTMLContent &= START_TIMESTAMP_DIV_WITH_ITALICS & ChangedData.FieldName & CLOSE_TIMESTAMP_ITALICS & ChangedData.ChangeDetected(j - 1) & CLOSE_DIV
            HTMLContent &= CLOSE_DIV
        Next
        Return HTMLContent
    End Function


    ''' <summary>
    ''' Takes a history of a regular text field in the profile and generates a series of cells that reflect the changes
    ''' made over time to that text field.
    ''' </summary>
    ''' <param name="ChangedData">A Database.ChangedData structure representing changes over time for a field and the times those changes were made</param>
    ''' <returns>Returns HTML content containing <divs> and <spans> that style the prior and current content of the field. If there was no data passed an empty string will be returned.</returns>
    Private Function CalculateLogCell(ByRef ChangedData As ChangeRecord) As String
        ' If there isn't any changed data for this cell, bail and return no HTML content
        Try
            If ChangedData.ChangedField Is Nothing Then Return ""
            If ChangedData.ChangedField.Count = 0 Then Return ""
        Catch ex As Exception
            Return ""
        End Try

        ' If there is only one history record of the field and it's blank, bail and return no HTML content
        If ChangedData.ChangedField.Count = 1 And ChangedData.ChangedField(0) = "" Then Return ""

        ' Start the new HTML content and initialize a DiffPlex engine
        Dim HTMLContent As String
        Dim DiffEngine As New DiffPlex.DiffBuilder.InlineDiffBuilder()

        ' Display the original field's content and the time it was set to that content
        HTMLContent = START_LOGCELL_DIV & ChangedData.ChangedField(ChangedData.ChangedField.Count - 1) & START_TIMESTAMP_DIV_WITH_ITALICS & ChangedData.FieldName & CLOSE_TIMESTAMP_ITALICS & ChangedData.ChangeDetected(ChangedData.ChangedField.Count - 1) & CLOSE_DIV & CLOSE_DIV
        For j = ChangedData.ChangedField.Count - 1 To 1 Step -1
            ' Open a new cell to display the added/removed/modified/unchanged content
            HTMLContent &= START_LOGCELL_DIV
            Dim Model As DiffPlex.DiffBuilder.Model.DiffPaneModel = DiffEngine.BuildDiffModel(ChangedData.ChangedField(j), ChangedData.ChangedField(j - 1))
            For k = 0 To Model.Lines.Count - 1
                Select Case Model.Lines(k).Type
                    Case DiffPlex.DiffBuilder.Model.ChangeType.Deleted
                        ' This text was deleted, add a span for the deleted style
                        HTMLContent &= START_SPAN_DELETED
                    Case DiffPlex.DiffBuilder.Model.ChangeType.Inserted
                        ' This text was inserted, add a span for the inserted style
                        HTMLContent &= START_SPAN_INSERTED
                    Case DiffPlex.DiffBuilder.Model.ChangeType.Modified
                        ' This text was modified, add a span for the modified style
                        HTMLContent &= START_SPAN_MODIFIED
                    Case DiffPlex.DiffBuilder.Model.ChangeType.Unchanged
                        ' This text was unchanged, add a span for the unchanged style
                        HTMLContent &= START_SPAN_UNCHANGED
                End Select

                ' Close the span we just added after adding the text
                HTMLContent &= Model.Lines(k).Text & CLOSE_SPAN
            Next

            ' Add a timespan indicating when the modification was made and then close the whole cell
            HTMLContent &= START_TIMESTAMP_DIV_WITH_ITALICS & ChangedData.FieldName & CLOSE_TIMESTAMP_ITALICS & ChangedData.ChangeDetected(j - 1) & CLOSE_DIV
            HTMLContent &= CLOSE_DIV
        Next
        Return HTMLContent
    End Function


    ''' <summary>
    ''' Attempts to send an email to the preconfigured inbox with the content of the changed profile report.
    ''' </summary>
    ''' <param name="EmailSubject">The email subject</param>
    ''' <param name="EmailBody">The HTML content of the email body</param>
    ''' <param name="Images">A collection of Content ID/Picture data to embed in the email</param>
    ''' <returns>Returns True if the SMTP Server accepted the message, False otherwise.</returns>
    Public Function SendEmail(ByVal EmailSubject As String, ByVal EmailBody As String, ByVal Images As Collection(Of EmailImageAttachment)) As Boolean
        ' Create a new SMTP message
        Dim SMTPMessage As New SmtpMail("TryIt")
        SMTPMessage.From = pConfiguration.SmtpFromAddress
        SMTPMessage.To = pConfiguration.SmtpToAddress
        SMTPMessage.Subject = EmailSubject
        SMTPMessage.HtmlBody = EmailBody

        ' Add any images we might have from our collection, name them by Content ID
        Dim InstaLoggerJpg As New EmailImageAttachment()
        InstaLoggerJpg.CID = "instalogger.jpg"
        InstaLoggerJpg.Picture = IO.File.ReadAllBytes("instalogger.jpg")
        Call Images.Add(InstaLoggerJpg)
        For i = 0 To Images.Count - 1
            Call SMTPMessage.AddAttachment(Images(i).CID, Images(i).Picture)
        Next

        ' Create an SMTPServer object to communicate with the server, configure it with our Configuration object
        Dim SMTPServer As New SmtpServer(pConfiguration.SmtpMailServer, pConfiguration.SmtpPort)
        If pConfiguration.SmtpUsername <> "" Then
            SMTPServer.User = pConfiguration.SmtpUsername
            SMTPServer.Password = pConfiguration.SmtpPassword
        End If
        If pConfiguration.SmtpUseTLS Then
            SMTPServer.ConnectType = SmtpConnectType.ConnectTryTLS
        End If

        ' Try to send the message, return True if it was successful and False otherwise
        Dim SMTPClient As New SmtpClient
        Try
            Call SMTPClient.SendMail(SMTPServer, SMTPMessage)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function


    ''' <summary>
    ''' This function resizes the image to be now wider or taller than MAXIMUM_PHOTO_THUMBNAIL_HEIGHT_OR_WIDTH. The
    ''' underlying image data is changed upon return of the function.
    ''' </summary>
    ''' <param name="ImageData">This is a byte array representing the raw image data. When the function returns it
    ''' contains raw image data if the image had to be resized.</param>
    Private Sub ResizeImageToThumbnail(ByRef ImageData() As Byte)
        Using Image As Image = Image.Load(ImageData)
            ' If the image isn't taller or wider than the maximum we can bail
            If Image.Height < MAXIMUM_PHOTO_THUMBNAIL_HEIGHT_OR_WIDTH And Image.Width < MAXIMUM_PHOTO_THUMBNAIL_HEIGHT_OR_WIDTH Then
                Return
            End If

            Dim Scale As Double
            ' Calculate the scale to resize the image by
            If Image.Height > Image.Width Then
                Scale = MAXIMUM_PHOTO_THUMBNAIL_HEIGHT_OR_WIDTH / Image.Height
            ElseIf Image.Width > Image.Height Then
                Scale = MAXIMUM_PHOTO_THUMBNAIL_HEIGHT_OR_WIDTH / Image.Width
            Else
                ' Must be the same
                Scale = MAXIMUM_PHOTO_THUMBNAIL_HEIGHT_OR_WIDTH / Image.Height
            End If

            ' Modify the image
            Call Image.Mutate(Sub(c)
                                  c.Resize(Image.Width * Scale, Image.Height * Scale)
                              End Sub)

            ' Shoot the output via the passed array
            Dim MemStream As New MemoryStream()
            Call Image.SaveAsJpeg(MemStream)
            ReDim ImageData(MemStream.Length - 1)
            Call MemStream.Seek(0, SeekOrigin.Begin)
            Call MemStream.Read(ImageData, 0, ImageData.Length)
        End Using

        Return
    End Sub


    ''' <summary>
    ''' Calculates and returns HTML width= and height= fields, ensuring the image is not displayed wider
    ''' or taller than MAXIMUM_IMAGE_HEIGHT_OR_WIDTH. This function does not change the underlying image data.
    ''' </summary>
    ''' <param name="ImageData">A byte array with the raw image data</param>
    ''' <returns>Returns a string that can be plugged right into an <img> tag to scale the
    ''' display of the image down if necessary. If the image is already small enough it returns
    ''' an empty string.</returns>
    Private Function CalculateImageDisplaySize(ByRef ImageData() As Byte) As String
        ' Create a rendered version of the image data
        Dim Image As Image = Image.Load(ImageData)

        ' If the picture was neither taller nor wider than the limit don't return any size data
        If Image.Height < MAXIMUM_RESIZED_DISPLAY_IMAGE_HEIGHT_OR_WIDTH And Image.Width < MAXIMUM_RESIZED_DISPLAY_IMAGE_HEIGHT_OR_WIDTH Then
            Return ""
        End If

        ' If the picture is either taller or wider than 250 pixes, we're going to shrink it
        Dim Scale As Double
        If Image.Height > Image.Width Then
            Scale = MAXIMUM_RESIZED_DISPLAY_IMAGE_HEIGHT_OR_WIDTH / Image.Height
        ElseIf Image.Width > Image.Height Then
            Scale = MAXIMUM_RESIZED_DISPLAY_IMAGE_HEIGHT_OR_WIDTH / Image.Width
        Else
            ' Must be the same
            Scale = MAXIMUM_PHOTO_THUMBNAIL_HEIGHT_OR_WIDTH / Image.Height
        End If

        Return "width=" & (Image.Width * Scale) & " height=" & (Image.Height * Scale)
    End Function
End Class
