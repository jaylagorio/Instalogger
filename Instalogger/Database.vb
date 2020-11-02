''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' File: Database.vb
''' Author: Jay Lagorio
''' Date Changed: 31OCT2020
''' Purpose: Interfaces with the backing Microsoft SQL database to store data about the Followed
''' users. Each instance of this class contains one SqlConnection so multiple connections will 
'''require multiple instances of the class.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports DiffPlex.DiffBuilder

''' <summary>
''' Interfaces with the backing Microsoft SQL database to store data about the Followed users. Each instance of this
''' class contains one SqlConnection so multiple connections will require multiple instances of the class.
''' </summary>
Public Class Database
    ' The number of times to try any given transaction on the SQL Server
    Private Const DATABASE_COMMS_RETRIES As Integer = 10

    ' The number of milliseconds to wait between tries when retrying transactions
    Private Const DATABASE_COMMS_RETRY_DELTA_IN_MILLISECONDS As Integer = 5

    ' SELECT statements to tell if a profile has been collected in the past, either by Instagram ID or by username
    Private Const SELECT_PROFILE_EXISTS_IGID As String = "SELECT TOP 1 Username FROM UsernameChanges WHERE IGID = @IGID"
    Private Const SELECT_PROFILE_EXISTS_USERNAME As String = "SELECT TOP 1 Username FROM UsernameChanges WHERE Username = @Username"

    ' SELECT statements looking for the most recent entry for that type of data in each table, by username
    Private Const SELECT_RECENT_BIOGRAPHY_USERNAME As String = "SELECT TOP 1 Biography, UpdatedTime FROM BiographyChanges WHERE Username = @Username ORDER BY UpdatedTime DESC"
    Private Const SELECT_RECENT_FULLNAME_USERNAME As String = "SELECT TOP 1 FullName, UpdatedTime FROM FullNameChanges WHERE Username = @Username ORDER BY UpdatedTime DESC"
    Private Const SELECT_RECENT_EXTERNALURL_USERNAME As String = "SELECT TOP 1 ExternalURL, UpdatedTime FROM ExternalURLChanges WHERE Username = @Username ORDER BY UpdatedTime DESC"
    Private Const SELECT_RECENT_LOCATION_USERNAME As String = "SELECT TOP 1 Location, UpdatedTime FROM LocationChanges WHERE Username = @Username ORDER BY UpdatedTime DESC"
    Private Const SELECT_RECENT_PROFILEPICURL_USERNAME As String = "SELECT TOP 1 ProfilePicURL, ProfilePic, UpdatedTime FROM ProfilePicChanges WHERE Username = @Username ORDER BY UpdatedTime DESC"
    Private Const SELECT_RECENT_IGID_USERNAME As String = "SELECT TOP 1 IGID, Username, UpdatedTime FROM UsernameChanges WHERE Username = @Username ORDER BY UpdatedTime DESC"

    ' SELECT statements looking for the most recent entry for that type of data in each table, by IGID
    Private Const SELECT_RECENT_BIOGRAPHY_IGID As String = "SELECT TOP 1 Biography, UpdatedTime FROM BiographyChanges WHERE IGID = @IGID ORDER BY UpdatedTime DESC"
    Private Const SELECT_RECENT_FULLNAME_IGID As String = "SELECT TOP 1 FullName, UpdatedTime FROM FullNameChanges WHERE IGID = @IGID ORDER BY UpdatedTime DESC"
    Private Const SELECT_RECENT_EXTERNALURL_IGID As String = "SELECT TOP 1 ExternalURL, UpdatedTime FROM ExternalURLChanges WHERE IGID = @IGID ORDER BY UpdatedTime DESC"
    Private Const SELECT_RECENT_LOCATION_IGID As String = "SELECT TOP 1 Location, UpdatedTime FROM LocationChanges WHERE IGID = @IGID ORDER BY UpdatedTime DESC"
    Private Const SELECT_RECENT_PROFILEPICURL_IGID As String = "SELECT TOP 1 ProfilePicURL, ProfilePic, UpdatedTime FROM ProfilePicChanges WHERE IGID = @IGID ORDER BY UpdatedTime DESC"
    Private Const SELECT_RECENT_IGID_IGID As String = "SELECT TOP 1 IGID, Username, UpdatedTime FROM UsernameChanges WHERE IGID = @IGID ORDER BY UpdatedTime DESC"

    ' INSERT and UPDATE statements to add newly collected data about an Instagram user, by IGID
    Private Const INSERT_UPDATED_BIOGRAPHY_IGID As String = "INSERT INTO BiographyChanges (IGID, Username, Biography, UpdatedTime) VALUES (@IGID, @Username, @Biography, @UpdatedTime)"
    Private Const INSERT_UPDATED_FULLNAME_IGID As String = "INSERT INTO FullNameChanges (IGID, Username, FullName, UpdatedTime) VALUES (@IGID, @Username, @FullName, @UpdatedTime)"
    Private Const INSERT_UPDATED_EXTERNALURL_IGID As String = "INSERT INTO ExternalURLChanges (IGID, Username, ExternalURL, UpdatedTime) VALUES (@IGID, @Username, @ExternalURL, @UpdatedTime)"
    Private Const INSERT_UPDATED_LOCATION_IGID As String = "INSERT INTO LocationChanges (IGID, Username, Location, UpdatedTime) VALUES (@IGID, @Username, @Location, @UpdatedTime)"
    Private Const INSERT_UPDATED_PROFILEPICURL_IGID As String = "INSERT INTO ProfilePicChanges (IGID, Username, ProfilePicURL, UpdatedTime) VALUES (@IGID, @Username, @ProfilePicURL, @UpdatedTime)"
    Private Const INSERT_UPDATED_USERNAME_IGID As String = "INSERT INTO UsernameChanges (IGID, Username, UpdatedTime) VALUES (@IGID, @Username, @UpdatedTime)"
    Private Const INSERT_UPDATED_CHANGELOG_IGID As String = "INSERT INTO ChangeLog (IGID, UpdatedTime) VALUES (@IGID, @UpdatedTime)"
    Private Const UPDATE_PROFILEPICURL_IGID As String = "UPDATE ProfilePicChanges SET ProfilePicURL = @ProfilePicURL WHERE ProfilePicURL = @OldProfilePicURL"
    Private Const UPDATE_UPDATED_PHOTO_ID As String = "UPDATE ProfilePicChanges SET ProfilePic = @ProfilePic WHERE ID = @ID"

    ' SELECT statements to find the most recent changes for an Instagram user, by IGID
    Private Const SELECT_CHANGELOG_UPDATES_IGID_YESTERDAY As String = "SELECT DISTINCT IGID FROM ChangeLog WHERE (UpdatedTime > CONVERT(DATE, DATEADD(DAY, -1, GETDATE()), 101))"
    Private Const SELECT_BIOGRAPHY_UPDATES As String = "SELECT TOP 50 * FROM BiographyChanges WHERE IGID = @IGID ORDER BY UpdatedTime DESC"
    Private Const SELECT_FULLNAME_UPDATES As String = "SELECT TOP 50 * FROM FullNameChanges WHERE IGID = @IGID ORDER BY UpdatedTime DESC"
    Private Const SELECT_EXTERNALURL_UPDATES As String = "SELECT TOP 50 * FROM ExternalURLChanges WHERE IGID = @IGID ORDER BY UpdatedTime DESC"
    Private Const SELECT_LOCATION_UPDATES As String = "SELECT TOP 50 * FROM LocationChanges WHERE IGID = @IGID ORDER BY UpdatedTime DESC"
    Private Const SELECT_PROFILEPIC_UPDATES As String = "SELECT TOP 50 * FROM ProfilePicChanges WHERE IGID = @IGID ORDER BY UpdatedTime DESC"
    Private Const SELECT_USERNAME_UPDATES As String = "SELECT TOP 50 * FROM UsernameChanges WHERE IGID = @IGID ORDER BY UpdatedTime DESC"

    ' An enumeration of fields that may have changed during a profile's
    ' processing. These values can be OR'd together to represent that
    ' more than one changed at once.
    Public Enum ProfileField
        None = 0
        Username = 1
        FullName = 2
        Location = 4
        ExternalURL = 8
        ProfilePicURL = 16
        Biography = 32
    End Enum

    ' Holds the most recent state of an Instagram profile, as it is known to the database
    Public Structure CompiledProfile
        Dim IGID As String                          ' The Instagram ID of the user
        Dim Username As String                      ' The username
        Dim UsernameLastChanged As DateTime         ' The date/time the username last changed
        Dim FullName As String                      ' The user's full name
        Dim FullNameLastChanged As DateTime         ' The date/time the full name last changed
        Dim Location As String                      ' The user's location
        Dim LocationLastChanged As DateTime         ' The date/time the location last changed
        Dim ExternalUrl As String                   ' The user's external URL
        Dim ExternalUrlLastChanged As DateTime      ' The date/time the external URL last changed
        Dim ProfilePicUrl As String                 ' The URL of the user's profile picture
        Dim ProfilePicData() As Byte                ' The raw data representing that picture (JPEG, PNG, etc)
        Dim ProfilePicUrlLastChanged As DateTime    ' The date/time the profile picture was changed
        Dim Biography As String                     ' The user's biography field
        Dim BiographyLastChanged As DateTime        ' The date/time the biography last changed
        Dim CollectionDate As DateTime              ' The date the above representation was collected
    End Structure

    ' Holds records of changes for a given field
    Public Structure ChangeRecord
        Dim IGID As String                  ' The IGID of the user in question
        Dim Username As String              ' The username
        Dim FieldName As String             ' The name of the field with the changes
        Dim ChangedField() As String        ' The content of the changed field, most-recent state first
        Dim ChangeDetected() As DateTime    ' The dates and times the changes were detected, corresponding to the same slot in ChangedField()
    End Structure

    ' The SqlConnection backing the connection to the SQL server. This connection is
    ' long-lived and requires our application to be single-threaded.
    Private pDBConnection As SqlConnection

    ''' <summary>
    ''' The Constructor for the Database class.
    ''' </summary>
    ''' <param name="DatabaseString">The database string used to connect to the SQL Server backing the data</param>
    Sub New(ByVal DatabaseString As String)
        pDBConnection = New SqlConnection(DatabaseString)

        For i = 1 To DATABASE_COMMS_RETRIES
            Try
                Call pDBConnection.Open()
                Return
            Catch ex As Exception
                If i = DATABASE_COMMS_RETRIES Then
                    ' If we've tried all possible times to connect to the database and it's failed
                    ' the last retry then throw the actual exception upward instead of eating it
                    ' as part of the retry process.
                    Throw ex
                End If

                Call Threading.Thread.Sleep(DATABASE_COMMS_RETRY_DELTA_IN_MILLISECONDS)
            End Try
        Next
    End Sub


    ''' <summary>
    ''' Get the most recent data collection about a user for all fields, by username, as they are currently known.
    ''' </summary>
    ''' <param name="Username">The Instagram username to query for</param>
    ''' <returns>Returns a CompiledProfile of fields for the user including when those fields were last changed.</returns>
    Public Function GetProfileMostRecentFieldsByUsername(ByVal Username As String) As CompiledProfile
        ' For each table and data type, get the most recent entry
        Dim BiographyDT As DataTable = GetSingleEntry(SELECT_RECENT_BIOGRAPHY_USERNAME, "@Username", Username)
        Dim FullNameDT As DataTable = GetSingleEntry(SELECT_RECENT_FULLNAME_USERNAME, "@Username", Username)
        Dim ExternalURLDT As DataTable = GetSingleEntry(SELECT_RECENT_EXTERNALURL_USERNAME, "@Username", Username)
        Dim LocationDT As DataTable = GetSingleEntry(SELECT_RECENT_LOCATION_USERNAME, "@Username", Username)
        Dim ProfilePicURLDT As DataTable = GetSingleEntry(SELECT_RECENT_PROFILEPICURL_USERNAME, "@Username", Username)
        Dim UsernameDT As DataTable = GetSingleEntry(SELECT_RECENT_IGID_USERNAME, "@Username", Username)

        ' Merge those entries into one CompiledProfile and return them
        Return MergeProfileDataTables(BiographyDT, FullNameDT, ExternalURLDT, LocationDT, ProfilePicURLDT, UsernameDT)
    End Function


    ''' <summary>
    ''' Get the most recent data collection about a user for all fields, by Instagram ID, as they are currently known.
    ''' </summary>
    ''' <param name="IGID">The Instagram ID to query for</param>
    ''' <returns>Returns a CompiledProfile of fields for the user including when those fields were last changed.</returns>
    Public Function GetProfileMostRecentFieldsByIGID(ByVal IGID As String) As CompiledProfile
        ' For each table and data type, get the most recent entry
        Dim BiographyDT As DataTable = GetSingleEntry(SELECT_RECENT_BIOGRAPHY_IGID, "@IGID", IGID)
        Dim FullNameDT As DataTable = GetSingleEntry(SELECT_RECENT_FULLNAME_IGID, "@IGID", IGID)
        Dim ExternalURLDT As DataTable = GetSingleEntry(SELECT_RECENT_EXTERNALURL_IGID, "@IGID", IGID)
        Dim LocationDT As DataTable = GetSingleEntry(SELECT_RECENT_LOCATION_IGID, "@IGID", IGID)
        Dim ProfilePicURLDT As DataTable = GetSingleEntry(SELECT_RECENT_PROFILEPICURL_IGID, "@IGID", IGID)
        Dim UsernameDT As DataTable = GetSingleEntry(SELECT_RECENT_IGID_IGID, "@IGID", IGID)

        ' Merge those entries into one CompiledProfile and return them
        Return MergeProfileDataTables(BiographyDT, FullNameDT, ExternalURLDT, LocationDT, ProfilePicURLDT, UsernameDT)
    End Function


    ''' <summary>
    ''' Merges the results of all the passed DataTables into one common data structure accounting for any null fields.
    ''' </summary>
    ''' <param name="BiographyDT">A DataTable representing the BiographyChanges table</param>
    ''' <param name="FullNameDT">A DataTable representing the FullNameChanges table</param>
    ''' <param name="ExternalURLDT">A DataTable representing the ExternalURLChanges table</param>
    ''' <param name="LocationDT">A DataTable representing the LocationChanges table</param>
    ''' <param name="ProfilePicURLDT">A DataTable representing the ProfilePicChanges table</param>
    ''' <param name="UsernameDT">A DataTable representing the UsernameChanges table</param>
    ''' <returns>Returns a CompiledProfile with the contents of all the DataTables passed in. Some fields may be empty if they were empty in that respective DataTable.</returns>
    Private Function MergeProfileDataTables(ByRef BiographyDT As DataTable, ByRef FullNameDT As DataTable, ByRef ExternalURLDT As DataTable, ByRef LocationDT As DataTable, ByRef ProfilePicURLDT As DataTable, ByRef UsernameDT As DataTable) As CompiledProfile
        Dim CombinedProfile As New CompiledProfile

        ' Fill out each field with the values from their respective tables. Some fields can be empty
        ' and those are taken into account as well.

        If BiographyDT.Rows.Count > 0 Then
            CombinedProfile.BiographyLastChanged = BiographyDT.Rows(0).Item("UpdatedTime")
            If BiographyDT.Rows(0).IsNull("Biography") Then
                CombinedProfile.Biography = ""
            Else
                CombinedProfile.Biography = BiographyDT.Rows(0).Item("Biography")
            End If
        Else
            CombinedProfile.Biography = ""
            CombinedProfile.BiographyLastChanged = DateTime.MinValue
        End If

        If FullNameDT.Rows.Count > 0 Then
            CombinedProfile.FullNameLastChanged = FullNameDT.Rows(0).Item("UpdatedTime")
            If FullNameDT.Rows(0).IsNull("FullName") Then
                CombinedProfile.FullName = ""
            Else
                CombinedProfile.FullName = FullNameDT.Rows(0).Item("FullName")
            End If
        Else
            CombinedProfile.FullName = ""
            CombinedProfile.FullNameLastChanged = DateTime.MinValue
        End If

        If ExternalURLDT.Rows.Count > 0 Then
            CombinedProfile.ExternalUrlLastChanged = ExternalURLDT.Rows(0).Item("UpdatedTime")
            If ExternalURLDT.Rows(0).IsNull("ExternalURL") Then
                CombinedProfile.ExternalUrl = ""
            Else
                CombinedProfile.ExternalUrl = ExternalURLDT.Rows(0).Item("ExternalURL")
            End If
        Else
            CombinedProfile.ExternalUrl = ""
            CombinedProfile.ExternalUrlLastChanged = DateTime.MinValue
        End If

        If LocationDT.Rows.Count > 0 Then
            CombinedProfile.LocationLastChanged = LocationDT.Rows(0).Item("UpdatedTime")
            If LocationDT.Rows(0).IsNull("Location") Then
                CombinedProfile.Location = ""
            Else
                CombinedProfile.Location = LocationDT.Rows(0).Item("Location")
            End If
        End If

        If ProfilePicURLDT.Rows.Count > 0 Then
            CombinedProfile.ProfilePicUrlLastChanged = ProfilePicURLDT.Rows(0).Item("UpdatedTime")
            CombinedProfile.ProfilePicUrl = ProfilePicURLDT.Rows(0).Item("ProfilePicURL")
        Else
            CombinedProfile.ProfilePicUrl = ""
            CombinedProfile.ProfilePicUrlLastChanged = DateTime.MinValue
        End If

        If UsernameDT.Rows.Count > 0 Then
            CombinedProfile.UsernameLastChanged = UsernameDT.Rows(0).Item("UpdatedTime")
            CombinedProfile.Username = UsernameDT.Rows(0).Item("Username")
        Else
            CombinedProfile.Username = ""
            CombinedProfile.UsernameLastChanged = DateTime.MinValue
        End If

        ' Return the content of the combined profile
        Return CombinedProfile
    End Function


    ''' <summary>
    ''' Update any fields in the passed CompiledProfile in the database, but only if those fields
    ''' have changed since the most recent entry in the respective database table.
    ''' </summary>
    ''' <param name="Profile">A CompiledProfile with all of the updated fields filled in and the CollectionDate
    ''' field indicating the time the data was collected.</param>
    ''' <returns>Returns True if all fields were updated as necessary, False if any errors occur.</returns>
    Public Function UpdateProfile(ByVal Profile As CompiledProfile) As Boolean
        Dim Success As Boolean = True
        If ProfileExistsByIGID(Profile.IGID) Then
            Dim ProfileFieldModified As ProfileField = ProfileField.None
            Dim ExistingProfileData As CompiledProfile = GetProfileMostRecentFieldsByIGID(Profile.IGID)
            Dim DiffBuilder As New InlineDiffBuilder

            ' Compare the data. If it's different we'll want to save it.
            If DiffBuilder.BuildDiffModel(ExistingProfileData.Biography, Profile.Biography).HasDifferences Then
                If Not InsertModificationRecord(INSERT_UPDATED_BIOGRAPHY_IGID, Profile.IGID, Profile.Username, "@Biography", Profile.Biography, Profile.CollectionDate) Then
                    Success = False
                Else
                    ProfileFieldModified = ProfileFieldModified Or ProfileField.Biography
                End If
            End If

            If DiffBuilder.BuildDiffModel(ExistingProfileData.FullName, Profile.FullName).HasDifferences Then
                If Not InsertModificationRecord(INSERT_UPDATED_FULLNAME_IGID, Profile.IGID, Profile.Username, "@FullName", Profile.FullName, Profile.CollectionDate) Then
                    Success = False
                Else
                    ProfileFieldModified = ProfileFieldModified Or ProfileField.FullName
                End If
            End If

            If ExistingProfileData.ExternalUrl <> "" And Profile.ExternalUrl <> "" Then
                If DiffBuilder.BuildDiffModel(ExistingProfileData.ExternalUrl, Profile.ExternalUrl).HasDifferences Then
                    If Not InsertModificationRecord(INSERT_UPDATED_EXTERNALURL_IGID, Profile.IGID, Profile.Username, "@ExternalURL", Profile.ExternalUrl, Profile.CollectionDate) Then
                        Success = False
                    Else
                        ProfileFieldModified = ProfileFieldModified Or ProfileField.ExternalURL
                    End If
                End If
            ElseIf ExistingProfileData.ExternalUrl = "" And Profile.ExternalUrl <> "" Then
                If Not InsertModificationRecord(INSERT_UPDATED_EXTERNALURL_IGID, Profile.IGID, Profile.Username, "@ExternalURL", Profile.ExternalUrl, Profile.CollectionDate) Then
                    Success = False
                Else
                    ProfileFieldModified = ProfileFieldModified Or ProfileField.ExternalURL
                End If
            ElseIf ExistingProfileData.ExternalUrl <> "" And Profile.ExternalUrl = "" Then
                If Not InsertModificationRecord(INSERT_UPDATED_EXTERNALURL_IGID, Profile.IGID, Profile.Username, "@ExternalURL", Profile.ExternalUrl, Profile.CollectionDate) Then
                    Success = False
                Else
                    ProfileFieldModified = ProfileFieldModified Or ProfileField.ExternalURL
                End If
            End If

            If ExistingProfileData.Location <> "" And Profile.Location <> "" Then
                If DiffBuilder.BuildDiffModel(ExistingProfileData.Location, Profile.Location).HasDifferences Then
                    If Not InsertModificationRecord(INSERT_UPDATED_LOCATION_IGID, Profile.IGID, Profile.Username, "@Location", Profile.Location, Profile.CollectionDate) Then
                        Success = False
                    Else
                        ProfileFieldModified = ProfileFieldModified Or ProfileField.Location
                    End If
                End If
            ElseIf ExistingProfileData.Location <> "" And Profile.Location = "" Then
                If Not InsertModificationRecord(INSERT_UPDATED_LOCATION_IGID, Profile.IGID, Profile.Username, "@Location", Profile.Location, Profile.CollectionDate) Then
                    Success = False
                Else
                    ProfileFieldModified = ProfileFieldModified Or ProfileField.Location
                End If
            ElseIf ExistingProfileData.Location = "" And Profile.Location <> "" Then
                If Not InsertModificationRecord(INSERT_UPDATED_LOCATION_IGID, Profile.IGID, Profile.Username, "@Location", Profile.Location, Profile.CollectionDate) Then
                    Success = False
                Else
                    ProfileFieldModified = ProfileFieldModified Or ProfileField.Location
                End If
            End If

            If DiffBuilder.BuildDiffModel(ExistingProfileData.Username, Profile.Username).HasDifferences Then
                If Not InsertModificationRecord(INSERT_UPDATED_USERNAME_IGID, Profile.IGID, Profile.Username, "@Username", Profile.Username, Profile.CollectionDate) Then
                    Success = False
                Else
                    ProfileFieldModified = ProfileFieldModified Or ProfileField.Username
                End If
            End If

            ' For the profile picture URLs we compare the filename because sometimes the images, though the same, change hosting servers
            ' or Instagram throws access tokens and tags into the query string, and you can't download the picture without them so we 
            ' have to save them but for the purposes of change detection the host through the filename components are what we care about.
            ' Alternatively, if the photo hasn't changed but the query string has we won't be able to download the photo later so keep
            ' track of the query string too but don't mark the photo as changed.
            If ExistingProfileData.ProfilePicUrl = "" Then
                ' No profile picture has been saved before this run
                If Not InsertModificationRecord(INSERT_UPDATED_PROFILEPICURL_IGID, Profile.IGID, Profile.Username, "@ProfilePicURL", Profile.ProfilePicUrl, Profile.CollectionDate) Then
                    Success = False
                Else
                    If Not InsertModifiedProfilePicture(ReturnLastInsertID(), Profile.ProfilePicData) Then
                        Success = False
                    Else
                        ProfileFieldModified = ProfileFieldModified Or ProfileField.ProfilePicURL
                    End If
                End If
            Else
                Dim ExistingProfilePicFilename As String = Path.GetFileName(New Uri(ExistingProfileData.ProfilePicUrl).LocalPath)
                Dim NewProfilePicFilename As String = Path.GetFileName(New Uri(Profile.ProfilePicUrl).LocalPath)
                If ExistingProfilePicFilename <> NewProfilePicFilename Then
                    ' The image has changed, store the new one and mark it changed
                    If Not InsertModificationRecord(INSERT_UPDATED_PROFILEPICURL_IGID, Profile.IGID, Profile.Username, "@ProfilePicURL", Profile.ProfilePicUrl, Profile.CollectionDate) Then
                        Success = False
                    Else
                        If Not InsertModifiedProfilePicture(ReturnLastInsertID(), Profile.ProfilePicData) Then
                            Success = False
                        Else
                            ProfileFieldModified = ProfileFieldModified Or ProfileField.ProfilePicURL
                        End If
                    End If
                ElseIf ExistingProfileData.ProfilePicUrl.Substring(ExistingProfileData.ProfilePicUrl.IndexOf("?")) <> Profile.ProfilePicUrl.Substring(Profile.ProfilePicUrl.IndexOf("?")) Then
                    '    ' We don't need to store the photo because the photo is the same, just the host holding the file or the access querystring changed.
                    If Not UpdateProfilePictureURL(ExistingProfileData.ProfilePicUrl, Profile.ProfilePicUrl) Then
                        Success = False
                    End If
                    '    End If
                End If
            End If

            ' Check to see that any of the profile fields were modified, and depending on which ones were log to the
            ' console and also insert a record into the ChangeLog table to show it was updated recently
            If Not ProfileFieldModified = ProfileField.None Then
                Dim Fields As String = ""
                Dim Comparison As String = ""
                If ProfileFieldModified And ProfileField.Biography Then
                    Fields &= "Biography, "
                    Comparison &= vbTab & "Old: " & ExistingProfileData.Biography & vbCrLf & vbTab & "New: " & Profile.Biography & vbCrLf
                End If
                If ProfileFieldModified And ProfileField.FullName Then
                    Fields &= "FullName, "
                    Comparison &= vbTab & "Old: " & ExistingProfileData.FullName & vbCrLf & vbTab & "New: " & Profile.FullName & vbCrLf
                End If
                If ProfileFieldModified And ProfileField.ExternalURL Then
                    Fields &= "ExternalURL, "
                    Comparison &= vbTab & "Old: " & ExistingProfileData.ExternalUrl & vbCrLf & vbTab & "New: " & Profile.ExternalUrl & vbCrLf
                End If
                If ProfileFieldModified And ProfileField.Location Then
                    Fields &= "Location, "
                    Comparison &= vbTab & "Old: " & ExistingProfileData.Location & vbCrLf & vbTab & "New: " & Profile.Location & vbCrLf
                End If
                If ProfileFieldModified And ProfileField.ProfilePicURL Then
                    Fields &= "ProfilePic, "
                    Comparison &= vbTab & "Old: " & ExistingProfileData.ProfilePicUrl & vbCrLf & vbTab & "New: " & Profile.ProfilePicUrl & vbCrLf
                End If
                If ProfileFieldModified And ProfileField.Username Then
                    Fields &= "Username"
                    Comparison &= vbTab & "Old: " & ExistingProfileData.Username & vbCrLf & vbTab & "New: " & Profile.Username & vbCrLf
                End If
                Call LogConsole("[*] Modified trigger: " & Profile.FullName & " (" & Profile.Username & "): " & Fields & vbCrLf & Comparison)
                If Not InsertChangelogRecord(Profile.IGID, Profile.CollectionDate) Then
                    Success = False
                End If
            End If
        Else
            ' This entire profile is new, so insert all of the fields into the database
            If Not InsertModificationRecord(INSERT_UPDATED_BIOGRAPHY_IGID, Profile.IGID, Profile.Username, "@Biography", Profile.Biography, Profile.CollectionDate) Then
                Success = False
            End If
            If Not InsertModificationRecord(INSERT_UPDATED_FULLNAME_IGID, Profile.IGID, Profile.Username, "@FullName", Profile.FullName, Profile.CollectionDate) Then
                Success = False
            End If
            If Not InsertModificationRecord(INSERT_UPDATED_EXTERNALURL_IGID, Profile.IGID, Profile.Username, "@ExternalURL", Profile.ExternalUrl, Profile.CollectionDate) Then
                Success = False
            End If
            If Not InsertModificationRecord(INSERT_UPDATED_LOCATION_IGID, Profile.IGID, Profile.Username, "@Location", Profile.Location, Profile.CollectionDate) Then
                Success = False
            End If
            If Not InsertModificationRecord(INSERT_UPDATED_USERNAME_IGID, Profile.IGID, Profile.Username, "@Username", Profile.Username, Profile.CollectionDate) Then
                Success = False
            End If
            If Not InsertModificationRecord(INSERT_UPDATED_PROFILEPICURL_IGID, Profile.IGID, Profile.Username, "@ProfilePicURL", Profile.ProfilePicUrl, Profile.CollectionDate) Then
                Success = False
            Else
                If Not InsertModifiedProfilePicture(ReturnLastInsertID(), Profile.ProfilePicData) Then
                    Success = False
                End If
            End If
            If Not InsertChangelogRecord(Profile.IGID, Profile.CollectionDate) Then
                Success = False
            End If
        End If

        ' Return whether there were any failures
        Return Success
    End Function


    ''' <summary>
    ''' Inserts a changelog record into the database
    ''' </summary>
    ''' <param name="IGID">The IGID of the profile that had some aspect change</param>
    ''' <param name="UpdatedTime">The collection date/time of the change</param>
    ''' <returns>True if the record was inserted successfully, False otherwise.</returns>
    Private Function InsertChangelogRecord(ByVal IGID As String, ByVal UpdatedTime As DateTime) As Boolean
        Dim SQLCommand As New SqlCommand(INSERT_UPDATED_CHANGELOG_IGID, pDBConnection)
        Call SQLCommand.Parameters.Add(New SqlParameter("@IGID", IGID))
        Call SQLCommand.Parameters.Add(New SqlParameter("@UpdatedTime", UpdatedTime))

        For i = 0 To DATABASE_COMMS_RETRIES
            Try
                Return (SQLCommand.ExecuteNonQuery() > 0)
            Catch ex As Exception
                Threading.Thread.Sleep(DATABASE_COMMS_RETRY_DELTA_IN_MILLISECONDS)
            End Try
        Next

        Return False
    End Function


    ''' <summary>
    ''' Inserts a record into the database where a given variable name/variable value have been identified.
    ''' </summary>
    ''' <param name="SQLStatement">The SQL statement to use for the INSERT statement</param>
    ''' <param name="IGID">The IGID of the profile</param>
    ''' <param name="Username">The username of the profile</param>
    ''' <param name="VariableName">The variable name to insert, like @Biography or @Location</param>
    ''' <param name="VariableValue">The value of the variable to insert</param>
    ''' <param name="UpdatedTime">The collection time of the record</param>
    ''' <returns>Returns True if the INSERT completed successfully, False otherwise.</returns>
    Private Function InsertModificationRecord(ByVal SQLStatement As String, ByVal IGID As String, ByVal Username As String, ByVal VariableName As String, ByVal VariableValue As String, ByVal UpdatedTime As DateTime) As Boolean
        Dim SQLCommand As New SqlCommand(SQLStatement, pDBConnection)
        Call SQLCommand.Parameters.Add(New SqlParameter("@IGID", IGID))
        Call SQLCommand.Parameters.Add(New SqlParameter("@Username", Username))
        Call SQLCommand.Parameters.Add(New SqlParameter("@UpdatedTime", UpdatedTime))

        ' Sometimes the modified record needs to be a Username, but we can see above that that's already included
        ' in most other tables. This checks to make sure that the variable name being inserted hasn't already been
        ' inserted above.
        If SQLCommand.Parameters.IndexOf(VariableName) = -1 Then
            ' Add the parameter to the collection, making sure to switch out Nothing or an empty string for DBNULL
            If VariableValue = "" Or VariableValue Is Nothing Then
                Call SQLCommand.Parameters.Add(New SqlParameter(VariableName, DBNull.Value))
            Else
                Call SQLCommand.Parameters.Add(New SqlParameter(VariableName, VariableValue))
            End If
        End If

        For i = 0 To DATABASE_COMMS_RETRIES
            Try
                Return (SQLCommand.ExecuteNonQuery() > 0)
            Catch ex As Exception
                Threading.Thread.Sleep(DATABASE_COMMS_RETRY_DELTA_IN_MILLISECONDS)
            End Try
        Next

        Return False
    End Function


    ''' <summary>
    ''' Inserts a the profile picture data into a table row based on the row's ID.
    ''' </summary>
    ''' <param name="ID">The ID of the table row</param>
    ''' <param name="Image">A Byte() of the raw image data</param>
    ''' <returns>True if the image was inserted, False otherwise.</returns>
    Private Function UpdateProfilePictureURL(ByVal OldURL As String, ByVal NewURL As String) As Boolean
        Dim SQLCommand As New SqlCommand(UPDATE_PROFILEPICURL_IGID, pDBConnection)
        Call SQLCommand.Parameters.Add(New SqlParameter("@OldProfilePicURL", OldURL))
        Call SQLCommand.Parameters.Add(New SqlParameter("@ProfilePicURL", NewURL))

        For i = 0 To DATABASE_COMMS_RETRIES
            Try
                Return (SQLCommand.ExecuteNonQuery() > 0)
            Catch ex As Exception
                Threading.Thread.Sleep(DATABASE_COMMS_RETRY_DELTA_IN_MILLISECONDS)
            End Try
        Next

        Return False
    End Function


    ''' <summary>
    ''' Inserts a the profile picture data into a table row based on the row's ID.
    ''' </summary>
    ''' <param name="ID">The ID of the table row</param>
    ''' <param name="Image">A Byte() of the raw image data</param>
    ''' <returns>True if the image was inserted, False otherwise.</returns>
    Private Function InsertModifiedProfilePicture(ByVal ID As Integer, ByRef Image() As Byte) As Boolean
        Dim SQLCommand As New SqlCommand(UPDATE_UPDATED_PHOTO_ID, pDBConnection)
        Call SQLCommand.Parameters.Add(New SqlParameter("@ProfilePic", Image))
        Call SQLCommand.Parameters.Add(New SqlParameter("@ID", ID))

        For i = 0 To DATABASE_COMMS_RETRIES
            Try
                Return (SQLCommand.ExecuteNonQuery() > 0)
            Catch ex As Exception
                Threading.Thread.Sleep(DATABASE_COMMS_RETRY_DELTA_IN_MILLISECONDS)
            End Try
        Next

        Return False
    End Function


    ''' <summary>
    ''' Query the database for a profile with the passed username. Note that usernames can change over time.
    ''' </summary>
    ''' <param name="Username">The username to query for</param>
    ''' <returns>True if the username has ever been collected, False if not.</returns>
    Private Function ProfileExistsByUsername(ByVal Username As String) As Boolean
        Dim DataTable As DataTable = GetSingleEntry(SELECT_PROFILE_EXISTS_USERNAME, "@Username", Username)

        ' If even one record comes back it exists
        If DataTable.Rows.Count > 0 Then
            Return True
        End If

        Return False
    End Function


    ''' <summary>
    ''' Queries the database to see if there's a profile that has previously been collected based on IGID.
    ''' </summary>
    ''' <param name="IGID">The Instagram ID to check for</param>
    ''' <returns>True if the IGID has been collected at least once, False if not.</returns>
    Private Function ProfileExistsByIGID(ByVal IGID As String) As Boolean
        Dim DataTable As DataTable = GetSingleEntry(SELECT_PROFILE_EXISTS_IGID, "@IGID", IGID)

        ' If even one record comes back then it exists
        If DataTable.Rows.Count > 0 Then
            Return True
        End If

        Return False
    End Function


    ''' <summary>
    ''' Looks up a username by the given Instagram ID
    ''' </summary>
    ''' <param name="IGID">The Instagram ID to query for</param>
    ''' <returns>Returns the username belonging to the IGID, if it hasn't ever been collected it returns an empty string.</returns>
    Public Function GetUsernameByIGID(ByVal IGID As String) As String
        Dim DataTable As DataTable = GetSingleEntry(SELECT_PROFILE_EXISTS_IGID, "@IGID", IGID)
        If DataTable.Rows.Count > 0 Then
            Return DataTable.Rows(0).Item("Username")
        End If

        Return ""
    End Function


    ''' <summary>
    ''' Queries the database for all Instagram IDs with any fields that changed since midnight yesterday.
    ''' </summary>
    ''' <returns>Returns an array of strings of IGIDs, unless there are no accounts that
    ''' changed since yesterday and then it returns Nothing.</returns>
    Public Function GetChangedProfilesByIGID() As String()
        ' Fill a DataTable with the result of the query below
        Dim DataTable As New DataTable
        Dim SQLCommand As New SqlCommand(SELECT_CHANGELOG_UPDATES_IGID_YESTERDAY, pDBConnection)
        Dim SQLDataAdaptor As New SqlDataAdapter
        SQLDataAdaptor.SelectCommand = SQLCommand
        Call FillDataTableWithRetries(SQLDataAdaptor, DataTable)

        ' Return Nothing if there were no records, otherwise fill a string array
        ' with each of the IGIDs returned from the query.
        If DataTable.Rows.Count = 0 Then Return Nothing
        Dim IGIDs(DataTable.Rows.Count - 1) As String
        For i = 0 To DataTable.Rows.Count - 1
            IGIDs(i) = DataTable.Rows(i).Item("IGID")
        Next

        Return IGIDs
    End Function


    ''' <summary>
    ''' Retrieves the change history of the Biography field by Instagram ID.
    ''' </summary>
    ''' <param name="IGID">Instagram ID to look up</param>
    ''' <returns>Returns a ChangeRecord with the IGID, Username, the field name, and the change history/date and time of change. If no
    ''' data was found it returns a New ChangeRecord that's completely blank.</returns>
    Public Function GetChangedBiographiesByIGID(ByVal IGID As String) As ChangeRecord
        Dim Yesterday As Date = (New Date(Now.Year, Now.Month, Now.Day)).Subtract(New TimeSpan(1, 0, 0, 0))

        ' Fill the DataTable to query data by IGID
        Dim DataTable As New DataTable
        Dim SQLCommand As New SqlCommand(SELECT_BIOGRAPHY_UPDATES, pDBConnection)
        Call SQLCommand.Parameters.Add(New SqlParameter("@IGID", IGID))
        Dim SQLDataAdaptor As New SqlDataAdapter
        SQLDataAdaptor.SelectCommand = SQLCommand
        Call FillDataTableWithRetries(SQLDataAdaptor, DataTable)

        ' If there wasn't any data found just return a blank ChangeRecord
        If DataTable.Rows.Count = 0 Then Return New ChangeRecord

        Dim ChangeRecord As New ChangeRecord
        ChangeRecord.IGID = IGID
        ChangeRecord.Username = DataTable.Rows(0).Item("Username")
        ChangeRecord.FieldName = "Biography"

        ' The data is going to come back ordered recent first -> earliest last
        Dim Biographies(DataTable.Rows.Count - 1) As String
        For i = 0 To DataTable.Rows.Count - 1
            If DataTable.Rows(i).IsNull("Biography") Then
                Biographies(i) = ""
            Else
                Biographies(i) = DataTable.Rows(i).Item("Biography")
            End If

            ' If this is the first entry that came back that's later than yesterday we'll
            ' add it to the array and stop including data here.
            If DataTable.Rows(i).Item("UpdatedTime") < Yesterday Then
                If i < DataTable.Rows.Count - 1 Then
                    ReDim Preserve Biographies(i)
                    Exit For
                End If
            End If
        Next
        ChangeRecord.ChangedField = Biographies

        ' Add all the date/times the changes were made to the array at a corresponding time, such that
        ' ChangedField(i) was changed on ChangesDetected(i), ChangedField(i + 1) was changed
        ' on ChangesDetected(i + 1), etc
        Dim ChangesDetected(Biographies.Count - 1) As DateTime
        For i = 0 To ChangesDetected.Count - 1
            ChangesDetected(i) = DataTable.Rows(i).Item("UpdatedTime")
        Next
        ChangeRecord.ChangeDetected = ChangesDetected

        ' Don't return any records if the only one we have is earlier than yesterday
        If ChangeRecord.ChangedField.Count = 1 Then
            If ChangeRecord.ChangeDetected(0) < Yesterday Then
                Return New ChangeRecord
            End If
        End If

        Return ChangeRecord
    End Function


    ''' <summary>
    ''' Retrieves the change history of the FullName field by Instagram ID.
    ''' </summary>
    ''' <param name="IGID">Instagram ID to look up</param>
    ''' <returns>Returns a ChangeRecord with the IGID, Username, the field name, and the change history/date and time of change. If no
    ''' data was found it returns a New ChangeRecord that's completely blank.</returns>
    Public Function GetChangedFullNamesByIGID(ByVal IGID As String) As ChangeRecord
        Dim Yesterday As Date = (New Date(Now.Year, Now.Month, Now.Day)).Subtract(New TimeSpan(1, 0, 0, 0))

        ' Fill the DataTable to query data by IGID
        Dim DataTable As New DataTable
        Dim SQLCommand As New SqlCommand(SELECT_FULLNAME_UPDATES, pDBConnection)
        Call SQLCommand.Parameters.Add(New SqlParameter("@IGID", IGID))
        Dim SQLDataAdaptor As New SqlDataAdapter
        SQLDataAdaptor.SelectCommand = SQLCommand
        Call FillDataTableWithRetries(SQLDataAdaptor, DataTable)

        ' If there wasn't any data found just return a blank ChangeRecord
        If DataTable.Rows.Count = 0 Then Return New ChangeRecord

        Dim ChangeRecord As New ChangeRecord
        ChangeRecord.IGID = IGID
        ChangeRecord.Username = DataTable.Rows(0).Item("Username")
        ChangeRecord.FieldName = "Full Name"

        ' The data is going to come back ordered recent first -> earliest last
        Dim FullNames(DataTable.Rows.Count - 1) As String
        For i = 0 To DataTable.Rows.Count - 1
            If DataTable.Rows(i).IsNull("FullName") Then
                FullNames(i) = ""
            Else
                FullNames(i) = DataTable.Rows(i).Item("FullName")
            End If
            ' If this is the first entry that came back that's later than yesterday we'll
            ' add it to the array and stop including data here.
            If DataTable.Rows(i).Item("UpdatedTime") < Yesterday Then
                If i < DataTable.Rows.Count - 1 Then
                    ReDim Preserve FullNames(i)
                    Exit For
                End If
            End If
        Next
        ChangeRecord.ChangedField = FullNames

        ' Add all the date/times the changes were made to the array at a corresponding time, such that
        ' ChangedField(i) was changed on ChangesDetected(i), ChangedField(i + 1) was changed
        ' on ChangesDetected(i + 1), etc
        Dim ChangesDetected(FullNames.Count - 1) As DateTime
        For i = 0 To ChangesDetected.Count - 1
            ChangesDetected(i) = DataTable.Rows(i).Item("UpdatedTime")
        Next
        ChangeRecord.ChangeDetected = ChangesDetected

        ' Don't return any records if the only one we have is earlier than yesterday
        If ChangeRecord.ChangedField.Count = 1 Then
            If ChangeRecord.ChangeDetected(0) < Yesterday Then
                Return New ChangeRecord
            End If
        End If

        Return ChangeRecord
    End Function


    ''' <summary>
    ''' Retrieves the change history of the ExternalURLs field by Instagram ID.
    ''' </summary>
    ''' <param name="IGID">Instagram ID to look up</param>
    ''' <returns>Returns a ChangeRecord with the IGID, Username, the field name, and the change history/date and time of change. If no
    ''' data was found it returns a New ChangeRecord that's completely blank.</returns>
    Public Function GetChangedExternalURLsByIGID(ByVal IGID As String) As ChangeRecord
        Dim Yesterday As Date = (New Date(Now.Year, Now.Month, Now.Day)).Subtract(New TimeSpan(1, 0, 0, 0))

        ' Fill the DataTable to query data by IGID
        Dim DataTable As New DataTable
        Dim SQLCommand As New SqlCommand(SELECT_EXTERNALURL_UPDATES, pDBConnection)
        Call SQLCommand.Parameters.Add(New SqlParameter("@IGID", IGID))
        Dim SQLDataAdaptor As New SqlDataAdapter
        SQLDataAdaptor.SelectCommand = SQLCommand
        Call FillDataTableWithRetries(SQLDataAdaptor, DataTable)

        ' If there wasn't any data found just return a blank ChangeRecord
        If DataTable.Rows.Count = 0 Then Return New ChangeRecord

        Dim ChangeRecord As New ChangeRecord
        ChangeRecord.IGID = IGID
        ChangeRecord.Username = DataTable.Rows(0).Item("Username")
        ChangeRecord.FieldName = "External URL"

        ' The data is going to come back ordered recent first -> earliest last
        Dim ExternalURLs(DataTable.Rows.Count - 1) As String
        For i = 0 To DataTable.Rows.Count - 1
            If DataTable.Rows(i).IsNull("ExternalURL") Then
                ExternalURLs(i) = ""
            Else
                ExternalURLs(i) = DataTable.Rows(i).Item("ExternalURL")
            End If
            ' If this is the first entry that came back that's later than yesterday we'll
            ' add it to the array and stop including data here.
            If DataTable.Rows(i).Item("UpdatedTime") < Yesterday Then
                If i < DataTable.Rows.Count - 1 Then
                    ReDim Preserve ExternalURLs(i)
                    Exit For
                End If
            End If
        Next
        ChangeRecord.ChangedField = ExternalURLs

        ' Add all the date/times the changes were made to the array at a corresponding time, such that
        ' ChangedField(i) was changed on ChangesDetected(i), ChangedField(i + 1) was changed
        ' on ChangesDetected(i + 1), etc
        Dim ChangesDetected(ExternalURLs.Count - 1) As DateTime
        For i = 0 To ChangesDetected.Count - 1
            ChangesDetected(i) = DataTable.Rows(i).Item("UpdatedTime")
        Next
        ChangeRecord.ChangeDetected = ChangesDetected

        ' Don't return any records if the only one we have is earlier than yesterday
        If ChangeRecord.ChangedField.Count = 1 Then
            If ChangeRecord.ChangeDetected(0) < Yesterday Then
                Return New ChangeRecord
            End If
        End If

        Return ChangeRecord
    End Function


    ''' <summary>
    ''' Retrieves the change history of the Locations field by Instagram ID.
    ''' </summary>
    ''' <param name="IGID">Instagram ID to look up</param>
    ''' <returns>Returns a ChangeRecord with the IGID, Username, the field name, and the change history/date and time of change. If no
    ''' data was found it returns a New ChangeRecord that's completely blank.</returns>
    Public Function GetChangedLocationsByIGID(ByVal IGID As String) As ChangeRecord
        Dim Yesterday As Date = (New Date(Now.Year, Now.Month, Now.Day)).Subtract(New TimeSpan(1, 0, 0, 0))

        ' Fill the DataTable to query data by IGID
        Dim DataTable As New DataTable
        Dim SQLCommand As New SqlCommand(SELECT_LOCATION_UPDATES, pDBConnection)
        Call SQLCommand.Parameters.Add(New SqlParameter("@IGID", IGID))
        Dim SQLDataAdaptor As New SqlDataAdapter
        SQLDataAdaptor.SelectCommand = SQLCommand
        Call FillDataTableWithRetries(SQLDataAdaptor, DataTable)

        ' If there wasn't any data found just return a blank ChangeRecord
        If DataTable.Rows.Count = 0 Then Return New ChangeRecord

        Dim ChangeRecord As New ChangeRecord
        ChangeRecord.IGID = IGID
        ChangeRecord.Username = DataTable.Rows(0).Item("Username")
        ChangeRecord.FieldName = "Location"

        ' The data is going to come back ordered recent first -> earliest last
        Dim Locations(DataTable.Rows.Count - 1) As String
        For i = 0 To DataTable.Rows.Count - 1
            If DataTable.Rows(i).IsNull("Location") Then
                Locations(i) = ""
            Else
                Locations(i) = DataTable.Rows(i).Item("Location")
            End If
            ' If this is the first entry that came back that's later than yesterday we'll
            ' add it to the array and stop including data here.
            If DataTable.Rows(i).Item("UpdatedTime") < Yesterday Then
                If i < DataTable.Rows.Count - 1 Then
                    ReDim Preserve Locations(i)
                    Exit For
                End If
            End If
        Next
        ChangeRecord.ChangedField = Locations

        ' Add all the date/times the changes were made to the array at a corresponding time, such that
        ' ChangedField(i) was changed on ChangesDetected(i), ChangedField(i + 1) was changed
        ' on ChangesDetected(i + 1), etc
        Dim ChangesDetected(Locations.Count - 1) As DateTime
        For i = 0 To ChangesDetected.Count - 1
            ChangesDetected(i) = DataTable.Rows(i).Item("UpdatedTime")
        Next
        ChangeRecord.ChangeDetected = ChangesDetected

        ' Don't return any records if the only one we have is earlier than yesterday
        If ChangeRecord.ChangedField.Count = 1 Then
            If ChangeRecord.ChangeDetected(0) < Yesterday Then
                Return New ChangeRecord
            End If
        End If

        Return ChangeRecord
    End Function


    ''' <summary>
    ''' Retrieves the change history of the ProfilePicURL field by Instagram ID.
    ''' </summary>
    ''' <param name="IGID">Instagram ID to look up</param>
    ''' <returns>Returns a ChangeRecord with the IGID, Username, the field name, and the change history/date and time of change. If no
    ''' data was found it returns a New ChangeRecord that's completely blank.</returns>
    Public Function GetChangedProfilePicsByIGID(ByVal IGID As String) As ChangeRecord
        Dim Yesterday As Date = (New Date(Now.Year, Now.Month, Now.Day)).Subtract(New TimeSpan(1, 0, 0, 0))

        ' Fill the DataTable to query data by IGID
        Dim DataTable As New DataTable
        Dim SQLCommand As New SqlCommand(SELECT_PROFILEPIC_UPDATES, pDBConnection)
        Call SQLCommand.Parameters.Add(New SqlParameter("@IGID", IGID))
        Dim SQLDataAdaptor As New SqlDataAdapter
        SQLDataAdaptor.SelectCommand = SQLCommand
        Call FillDataTableWithRetries(SQLDataAdaptor, DataTable)

        ' If there wasn't any data found just return a blank ChangeRecord
        If DataTable.Rows.Count = 0 Then Return New ChangeRecord

        Dim ChangeRecord As New ChangeRecord
        ChangeRecord.IGID = IGID
        ChangeRecord.Username = DataTable.Rows(0).Item("Username")
        ChangeRecord.FieldName = "Profile Picture"

        ' The data is going to come back ordered recent first -> earliest last
        Dim ProfilePicURLs(DataTable.Rows.Count - 1) As String
        For i = 0 To DataTable.Rows.Count - 1
            If DataTable.Rows(i).IsNull("ProfilePicURL") Then
                ProfilePicURLs(i) = ""
            Else
                ProfilePicURLs(i) = DataTable.Rows(i).Item("ProfilePicURL")
            End If
            ' If this is the first entry that came back that's later than yesterday we'll
            ' add it to the array and stop including data here.
            If DataTable.Rows(i).Item("UpdatedTime") < Yesterday Then
                If i < DataTable.Rows.Count - 1 Then
                    ReDim Preserve ProfilePicURLs(i)
                    Exit For
                End If
            End If
        Next
        ChangeRecord.ChangedField = ProfilePicURLs

        ' Add all the date/times the changes were made to the array at a corresponding time, such that
        ' ChangedField(i) was changed on ChangesDetected(i), ChangedField(i + 1) was changed
        ' on ChangesDetected(i + 1), etc
        Dim ChangesDetected(ProfilePicURLs.Count - 1) As DateTime
        For i = 0 To ChangesDetected.Count - 1
            ChangesDetected(i) = DataTable.Rows(i).Item("UpdatedTime")
        Next
        ChangeRecord.ChangeDetected = ChangesDetected

        ' Don't return any records if the only one we have is earlier than yesterday
        If ChangeRecord.ChangedField.Count = 1 Then
            If ChangeRecord.ChangeDetected(0) < Yesterday Then
                Return New ChangeRecord
            End If
        End If

        Return ChangeRecord
    End Function


    ''' <summary>
    ''' Returns the most recent previously stored file representation of the profile image from Instagram.
    ''' </summary>
    ''' <param name="IGID">The Instagram ID of the latest stored profile picture to retrieve</param>
    ''' <returns>A Byte array with the stored image data if one is found, otherwise returns Nothing</returns>
    Public Function GetProfilePicData(ByVal IGID As String) As Byte()
        Dim ProfilePicDT As DataTable = GetSingleEntry(SELECT_RECENT_PROFILEPICURL_IGID, "@IGID", IGID)
        If ProfilePicDT.Rows.Count = 0 Then
            Return Nothing
        End If

        If ProfilePicDT.Rows(0).IsNull("ProfilePic") Then
            Return Nothing
        End If

        Return ProfilePicDT.Rows(0).Item("ProfilePic")
    End Function


    ''' <summary>
    ''' Retrieves the change history of the Username field by Instagram ID.
    ''' </summary>
    ''' <param name="IGID">Instagram ID to look up</param>
    ''' <returns>Returns a ChangeRecord with the IGID, Username, the field name, and the change history/date and time of change. If no
    ''' data was found it returns a New ChangeRecord that's completely blank.</returns>
    Public Function GetChangedUsernamesByIGID(ByVal IGID As String) As ChangeRecord
        Dim Yesterday As Date = (New Date(Now.Year, Now.Month, Now.Day)).Subtract(New TimeSpan(1, 0, 0, 0))

        ' Fill the DataTable to query data by IGID
        Dim DataTable As New DataTable
        Dim SQLCommand As New SqlCommand(SELECT_USERNAME_UPDATES, pDBConnection)
        Call SQLCommand.Parameters.Add(New SqlParameter("@IGID", IGID))
        Dim SQLDataAdaptor As New SqlDataAdapter
        SQLDataAdaptor.SelectCommand = SQLCommand
        Call FillDataTableWithRetries(SQLDataAdaptor, DataTable)

        ' If there wasn't any data found just return a blank ChangeRecord
        If DataTable.Rows.Count = 0 Then Return New ChangeRecord

        Dim ChangeRecord As New ChangeRecord
        ChangeRecord.IGID = IGID
        ChangeRecord.Username = DataTable.Rows(0).Item("Username")
        ChangeRecord.FieldName = "Username"

        ' The data is going to come back ordered recent first -> earliest last
        Dim Usernames(DataTable.Rows.Count - 1) As String
        For i = 0 To DataTable.Rows.Count - 1
            If DataTable.Rows(i).IsNull("Username") Then
                Usernames(i) = ""
            Else
                Usernames(i) = DataTable.Rows(i).Item("Username")
            End If
            ' If this is the first entry that came back that's later than yesterday we'll
            ' add it to the array and stop including data here.
            If DataTable.Rows(i).Item("UpdatedTime") < Yesterday Then
                If i < DataTable.Rows.Count - 1 Then
                    ReDim Preserve Usernames(i)
                    Exit For
                End If
            End If
        Next
        ChangeRecord.ChangedField = Usernames

        Dim ChangesDetected(Usernames.Count - 1) As DateTime
        For i = 0 To ChangesDetected.Count - 1
            ChangesDetected(i) = DataTable.Rows(i).Item("UpdatedTime")
        Next
        ChangeRecord.ChangeDetected = ChangesDetected

        ' Don't return any records if the only one we have is earlier than yesterday
        If ChangeRecord.ChangedField.Count = 1 Then
            If ChangeRecord.ChangeDetected(0) < Yesterday Then
                Return New ChangeRecord
            End If
        End If

        Return ChangeRecord
    End Function


    ''' <summary>
    ''' Returns a DataTable filled with the passed SQL Statement and the passed variable name/variable value.
    ''' </summary>
    ''' <param name="SQLStatement">A SQL statement to run against the database</param>
    ''' <param name="VariableName">A variable name to pass with the SQL Statement</param>
    ''' <param name="VariableValue">The value to assign the passed variable name</param>
    ''' <returns>A DataTable with data returned from the passed query. If the query didn't return any results the DataTable will be empty.</returns>
    Private Function GetSingleEntry(ByVal SQLStatement As String, ByVal VariableName As String, ByVal VariableValue As String) As DataTable
        ' Fill a DataTable with the result of the query below
        Dim DataTable As New DataTable
        Dim SQLCommand As New SqlCommand(SQLStatement, pDBConnection)
        Call SQLCommand.Parameters.Add(New SqlParameter(VariableName, VariableValue))
        Dim SQLDataAdaptor As New SqlDataAdapter
        SQLDataAdaptor.SelectCommand = SQLCommand
        Call FillDataTableWithRetries(SQLDataAdaptor, DataTable)

        Return DataTable
    End Function


    ''' <summary>
    ''' A quick way to abstract filling a DataTable with a SqlDataAdaptor with retries built in.
    ''' </summary>
    ''' <param name="SQLDataAdaptor">The SqlDataAdaptor you want to use to fill the DataTable</param>
    ''' <param name="DataTable">The DataTable to fill with the SqlDataAdaptor</param>
    ''' <returns>Returns True if the SqlDataAdaptor successfully filled the DataTable without throwing an exception, False otherwise.</returns>
    Private Function FillDataTableWithRetries(ByRef SQLDataAdaptor As SqlDataAdapter, ByRef DataTable As DataTable) As Boolean
        ' Try several times to fill the DataTable, retrying after a momentary wait on failure up to
        ' a DATABASE_COMMS_RETRIES number of times. Returns True as soon as it's successful.
        For i = 1 To DATABASE_COMMS_RETRIES
            Try
                Call SQLDataAdaptor.Fill(DataTable)
                Return True
            Catch ex As Exception
                Call Threading.Thread.Sleep(DATABASE_COMMS_RETRY_DELTA_IN_MILLISECONDS)
            End Try
        Next

        Return False
    End Function


    ''' <summary>
    ''' Returns the @@IDENTITY result from the DBConnection
    ''' </summary>
    ''' <returns>An Integer with the @@IDENTITY result</returns>
    Private Function ReturnLastInsertID() As Integer
        Dim SQLCommand As New SqlCommand("SELECT @@IDENTITY", pDBConnection)
        For i = 0 To DATABASE_COMMS_RETRIES
            Try
                Return SQLCommand.ExecuteScalar()
            Catch ex As Exception
                Threading.Thread.Sleep(DATABASE_COMMS_RETRY_DELTA_IN_MILLISECONDS)
            End Try
        Next

        Return 0
    End Function


    ''' <summary>
    ''' Logs a text line to the console, optionally you can specify the color to log with. The console text color
    ''' is automatically set back to the previous color on return.
    ''' </summary>
    ''' <param name="LogLine">Text to log</param>
    ''' <param name="Color">ConsoleColor to use, if specified</param>
    Private Sub LogConsole(ByVal LogLine As String, Optional ByVal Color As ConsoleColor = Nothing)
        Dim PrevColor As ConsoleColor = Console.ForegroundColor
        If Not Color = Nothing Then
            Console.ForegroundColor = Color
        End If
        Call Console.WriteLine(LogLine)
        Console.ForegroundColor = PrevColor
    End Sub
End Class
