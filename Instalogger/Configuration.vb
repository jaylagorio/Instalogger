''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' File: Configuration.vb
''' Author: Jay Lagorio
''' Date Changed: 31OCT2020
''' Purpose: This class holds the configuration data needed to communicate with the SQL Server
''' database, Instagram, and Twilio. It handles serialization and deserialization to JSON to save
''' and retrieve these settings between executions.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports System.IO
Imports System.Text.Json

''' <summary>
''' This class holds the configuration data needed to communicate with the SQL Server database, Instagram,
''' and Twilio. It handles serialization and deserialization to JSON to save and retrieve these settings between
''' executions.
''' </summary>
Public Class Configuration
    ' The default filename in the current application directory for the configuation file
    Private Const CONFIGURATION_FILENAME As String = "config.json"

    ' Stores a Microsoft SQL Server database connection string
    Public Property DatabaseString As String

    ' Username for Instagram
    Public Property InstagramUsername As String

    ' Password to log into the Instagram account
    Public Property InstagramPassword As String

    ' The private Base32 two factor secret data given by Instagram to log into the account
    Public Property InstagramTwoFactorAuthCode As String

    ' The Account SID used for Twilio SMS checking
    Public Property TwilioAccountSid As String

    ' The Authentication Token used for Twilio
    Public Property TwilioAuthToken As String

    ' The two factor phone number used to follow the SMS two factor flow
    Public Property TwilioPhoneNumber As String

    ' The SMTP server used for sending reports to the reporting email address
    Public Property SmtpMailServer As String

    ' The username to use to send mail from the SMTP server, if necessary
    Public Property SmtpUsername As String

    ' The password to use to send mail from the SMTP server, if necessary
    Public Property SmtpPassword As String

    ' The port to use to send mail from the SMTP server, defaults to 25
    Public Property SmtpPort As Integer = 25

    ' Indicates whether TLS is required for the chosen SMTP port, defaults to False
    Public Property SmtpUseTLS As Boolean = False

    ' The email address that the completed reports will be sent from
    Public Property SmtpFromAddress As String

    ' The email address that will receive the completed reports
    Public Property SmtpToAddress As String

    ' The last time the report was sent to the corresponding SmtpToAddress
    Public Property LastReportingTime As DateTime

    ' Any serialized login session data for Instagram
    Public Property PreviousSessionJSON As String

    ''' <summary>
    ''' Loads the JSON configuration file data into a new Configuration instance and returns it.
    ''' </summary>
    ''' <returns>Returns a new Configuration instance if the file exists, Nothing if the file doesn't exist.</returns>
    Public Shared Function LoadConfiguration() As Configuration
        ' Check that the file exists
        If File.Exists(CONFIGURATION_FILENAME) Then
            ' Returns a deserialized instance of the Configuration file contents
            Return JsonSerializer.Deserialize(File.ReadAllText(CONFIGURATION_FILENAME), (New Configuration).GetType())
        Else
            ' If the file doesn't exist we return Nothing to cause an error
            Return Nothing
        End If
    End Function


    ''' <summary>
    ''' Writes all the Configuration data to the file in the application directory, erasing any existing file data.
    ''' </summary>
    Public Sub SaveConfiguration()
        ' Serialize this class data and write it to the file
        Call File.WriteAllText(CONFIGURATION_FILENAME, JsonSerializer.Serialize(Me))
    End Sub
End Class
