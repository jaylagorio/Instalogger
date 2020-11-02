''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' File: Instagram.vb
''' Author: Jay Lagorio
''' Date Changed: 31OCT2020
''' Purpose: The interface between the InstagramAPISharp library and the Instagram service itself.
''' This simplifies the very complex login process and gets information about Followed users.
''' Each instance of this class represents an Instagram account being logged into so one will need
''' to be instantiated to log into multiple accounts.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports Instalogger.Database
Imports InstagramApiSharp.API
Imports InstagramApiSharp.Classes
Imports InstagramApiSharp.Classes.Models
Imports System.Collections.ObjectModel
Imports System.Net.Http
Imports TwoFactorAuthNet

''' <summary>
''' The interface between the InstagramAPISharp library and the Instagram service itself. This simplifies the very complex
''' login process and gets information about Followed users. Each instance of this class represents an Instagram account
''' being logged into so one will need to be instantiated to log into multiple accounts.
''' </summary>
Public Class Instagram
    ' The amount of time we pause just after downloading a Follower's profile so we don't get ratelimited
    ' as fast. Randomization may also deter the bot detector.
    Private Const PROFILE_DOWNLOAD_MINIMUM_PAUSE_IN_MILLISECONDS = 2000
    Private Const PROFILE_DOWNLOAD_MAXIMUM_PAUSE_IN_MILLISECONDS = 5000

    ' The first back off period for when we get ratelimited. This gets increased until ratelimiting ends,
    ' at which point it resets until the next time we get rate limited.
    Private Const DEFAULT_RATELIMIT_BACKOFF_PERIOD_IN_MILLISECONDS As Integer = 100000

    ' When we hit the rate limiter the second time for the same call it's not likely that leaving it the same
    ' length as it was in the first call is going to work. We make it longer each time we get stopped for the same
    ' request until the request goes through.
    Private Const RATE_LIMIT_MULTIPLIER As Double = 1.5

    ' Used to write updates to the console, but only after this increment of percentage
    Private Const PERCENTAGE_UPDATE_MODULUS As Integer = 25

    ' The period of validity for any given TOTP code generated using the TOTP secret key
    Private Const TOTP_CODE_VALIDITY_PERIOD_IN_MILLISECONDS As Integer = 30000

    ' Configuration object to hold authentication data, saved session state, etc
    Private pConfiguration As Configuration

    ' Interface to the Instagram API
    Private pApi As IInstaApi

    ''' <summary>
    ''' Instantiates a new interface to Instagram using the passed Configuration instance.
    ''' </summary>
    ''' <param name="Configuration">An instance of a loaded Configuration that includes Instagram credentials. Optionally it
    ''' includes Twilio credentials and/or Instagram Two Factor Authentication private key material.</param>
    Sub New(ByRef Configuration As Configuration)
        pConfiguration = Configuration
    End Sub


    ''' <summary>
    ''' A synchronous method to log into Instagram. Calls LoginAsync and passes its return value to the caller.
    ''' </summary>
    ''' <returns>Returns True if we logged in successfully, False if not.</returns>
    Public Function Login() As Boolean
        Return LoginAsync().ConfigureAwait(False).GetAwaiter().GetResult()
    End Function


    ''' <summary>
    ''' Uses the authentication information provided through the Configuration object provided to the
    ''' class constructor to log into Instagram. This will involve the username and password and may
    ''' involve saved session data from a prior login, Twilio authentication for SMS Two Factor login
    ''' or will use a Base32 key provided for "Authenticator App" Two Factor login.
    ''' </summary>
    ''' <returns>True if a fully logged in session was successful, False otherwise</returns>
    Public Async Function LoginAsync() As Task(Of Boolean)
        ' Construct the instance of the Instagram API and pass it credentials to login
        Dim B As Builder.InstaApiBuilder = InstagramApiSharp.API.Builder.InstaApiBuilder.CreateBuilder()
        Dim UserCreds As New UserSessionData
        UserCreds.UserName = pConfiguration.InstagramUsername
        UserCreds.Password = pConfiguration.InstagramPassword
        Call B.SetUser(UserCreds)
        pApi = B.Build()
        Call pApi.SetDevice(Android.DeviceInfo.AndroidDeviceGenerator.GetByName(Android.DeviceInfo.AndroidDevices.SAMSUNG_NOTE3))

        ' Check to see if there's a previously saved and logged in session in the Configuration
        ' that was passed in. If so load the API interface with that and return True after calling
        ' UserProcessor.GetCurrentUserAsync().
        If pConfiguration.PreviousSessionJSON <> "" Then
            Call pApi.LoadStateDataFromString(pConfiguration.PreviousSessionJSON)
            If pApi.IsUserAuthenticated Then
                Dim CurrentUserRes As IResult(Of InstaCurrentUser) = Await pApi.UserProcessor.GetCurrentUserAsync()
                If CurrentUserRes.Succeeded Then
                    If CurrentUserRes.Value.UserName.ToLower = pConfiguration.InstagramUsername.ToLower Then
                        Call LogConsole("[+] Session restored from serialized state!", ConsoleColor.Green)
                        Return True
                    End If
                End If
            End If
        End If

        ' Attempt the initial login using the Instagram username and password provided in the Configuration
        Call LogConsole("[*] Initial Login attempt...")
        Dim LoginRes As IResult(Of InstaLoginResult)
        LoginRes = Await pApi.LoginAsync()

        ' Check that the call completed successfully
        If LoginRes.Succeeded Then
            ' Check that login completed successfully
            If LoginRes.Value = InstaLoginResult.Success Then
                ' The most trivial login attempt succeeded! Save the state in the Configuration and write the
                ' Configuration and saved state to disk.
                Call LogConsole("[+] Login succeeded!", ConsoleColor.Green)
                pConfiguration.PreviousSessionJSON = pApi.GetStateDataAsString()
                Call pConfiguration.SaveConfiguration()
                Return True
            Else
                Call LogConsole("[-] The Login call succeeded but did not return a valid session: " & LoginRes.Value, ConsoleColor.Red)
                Return False
            End If
        Else
            ' There are several ways that the call to Login actually went fine but that there's more to do to successfully
            ' log into the account. Three of those ways include Two Factor challenges (SMS or TOTP) or a challenge where you
            ' have to specify the account is yours via an emailed code. The emailed code requires manual input.
            If LoginRes.Value = InstaLoginResult.TwoFactorRequired Then
                ' Get the style of Two Factor the service is expecting
                Dim TwoFactorInfo As IResult(Of InstaTwoFactorLoginInfo) = Await pApi.GetTwoFactorInfoAsync()
                Dim TwoFactorLoginRes As IResult(Of InstaLoginTwoFactorResult)

                If TwoFactorInfo.Succeeded Then
                    Dim TFACode As String
                    Select Case True
                        Case TwoFactorInfo.Value.SmsTwoFactorOn
                            ' Make sure Twilio credentials are configured before trying to use them
                            If pConfiguration.TwilioAccountSid <> "" And pConfiguration.TwilioAuthToken <> "" And pConfiguration.TwilioPhoneNumber <> "" Then
                                ' Create the TwilioHelper and grab the code from the most recent (or wait for a new, incoming) SMS message
                                Dim TwilioHelper As New TwilioHelper(pConfiguration.TwilioAccountSid, pConfiguration.TwilioAuthToken, pConfiguration.TwilioPhoneNumber)
                                TFACode = ProcessInstagram2FAText(Await TwilioHelper.GetMostRecentSMSAsync(True))

                                ' Throw the SMS 2FA code over to the service for login
                                TwoFactorLoginRes = Await pApi.TwoFactorLoginAsync(TFACode, True, InstagramApiSharp.Enums.InstaTwoFactorVerifyOptions.SmsCode)
                                If TwoFactorLoginRes.Succeeded Then
                                    Call LogConsole("[+] Login succeeded (2FA SMS Flow)!", ConsoleColor.Green)

                                    ' Serialize the logged in state of the service, save it with the configuration, and
                                    ' delete any stray SMS messages that might be left in the account
                                    pConfiguration.PreviousSessionJSON = pApi.GetStateDataAsString()
                                    Call pConfiguration.SaveConfiguration()
                                    Await TwilioHelper.DeleteAllSMSAsync()
                                    Return True
                                Else
                                    Call LogConsole("[-] Login failure (2FA SMS Flow)", ConsoleColor.Red)
                                    Return False
                                End If
                            Else
                                Call LogConsole("[-] Login failure (2FA SMS Flow requested, Twilio credentials not loaded)", ConsoleColor.Red)
                                Return False
                            End If
                        Case TwoFactorInfo.Value.ToTpTwoFactorOn
                            ' Check that a TOTP secret key was configured before using it
                            If pConfiguration.InstagramTwoFactorAuthCode <> "" Then
                                Dim TFANet As New TwoFactorAuth()

                                ' Get a TOTP Code based on the passed Configuration TOTP code and throw it to the service
                                TFACode = TFANet.GetCode(pConfiguration.InstagramTwoFactorAuthCode)
                                Call LogConsole("[*] Using TOTP 2FA Code: " & TFACode)
                                TwoFactorLoginRes = Await pApi.TwoFactorLoginAsync(TFACode, True, InstagramApiSharp.Enums.InstaTwoFactorVerifyOptions.AuthenticationApp)

                                ' If the code was expired or invalid we'll wait the length of the validity period (30 seconds)
                                ' and try again with the next valid code.
                                Do While (TwoFactorLoginRes.Value = InstaLoginTwoFactorResult.CodeExpired Or TwoFactorLoginRes.Value = InstaLoginTwoFactorResult.InvalidCode)
                                    Call LogConsole("[!] Code invalid, trying another...", ConsoleColor.Yellow)
                                    Await Task.Delay(TOTP_CODE_VALIDITY_PERIOD_IN_MILLISECONDS)
                                    TFACode = TFANet.GetCode(pConfiguration.InstagramTwoFactorAuthCode)
                                    Call LogConsole("[*] Using TOTP 2FA Code: " & TFACode)
                                    TwoFactorLoginRes = Await pApi.TwoFactorLoginAsync(TFACode, True, InstagramApiSharp.Enums.InstaTwoFactorVerifyOptions.AuthenticationApp)
                                Loop

                                If TwoFactorLoginRes.Succeeded Then
                                    Call LogConsole("[+] Login succeeded (TOTP 2FA flow)!", ConsoleColor.Green)

                                    ' Save the logged in session data to the Configuration, save the Configuration
                                    pConfiguration.PreviousSessionJSON = pApi.GetStateDataAsString()
                                    Call pConfiguration.SaveConfiguration()
                                    Return True
                                Else
                                    ' The TOTP code must have been valid but login failed.
                                    Call LogConsole("[-] Login failed with TOTP 2FA flow, encountered error response " & TwoFactorLoginRes.Info.ResponseType, ConsoleColor.Red)
                                    Return False
                                End If
                            Else
                                Call LogConsole("[-] Login failure (2FA TOTP Flow requested, TOTP Private Key not loaded)", ConsoleColor.Red)
                                Return False
                            End If
                    End Select
                Else
                    Call LogConsole("[-] Failed to retrieve Two Factor login parameters", ConsoleColor.Red)
                    Return False
                End If
            ElseIf LoginRes.Value = InstaLoginResult.ChallengeRequired Then
                ' The account is in challenge mode, Instagram recognizes this login attempt
                ' to be abnormal. Get the challenge verification method.
                Dim ChallengeRes As IResult(Of InstaChallengeRequireVerifyMethod) = Await pApi.GetChallengeRequireVerifyMethodAsync()
                If ChallengeRes.Succeeded Then
                    Call LogConsole("[*] Email Challenge flow")
                    ' Request an email to be sent to the account email address
                    Dim EmailChallengeRes As IResult(Of InstaChallengeRequireEmailVerify) = Await pApi.RequestVerifyCodeToEmailForChallengeRequireAsync()
                    If EmailChallengeRes.Succeeded Then
                        ' When the user receives the email they have to open it, get the code, and type it in interactively
                        ' to have it sent back to the service.
                        Console.Write("[?] Please enter email verification code: ")
                        Dim ChallengeCode As String = Console.ReadLine()
                        LoginRes = Await pApi.VerifyCodeForChallengeRequireAsync(ChallengeCode)
                        If LoginRes.Succeeded Then
                            If LoginRes.Value = InstaLoginResult.Success Then
                                Call LogConsole("[+] Login success!", ConsoleColor.Green)

                                ' Store the logged in session to the Configuration object and save it to the file system
                                pConfiguration.PreviousSessionJSON = pApi.GetStateDataAsString()
                                Call pConfiguration.SaveConfiguration()
                                Return True
                            ElseIf LoginRes.Value = InstaLoginResult.ChallengeRequired Then
                                ' Get the challenge data
                                Dim ChallengeInfoRes As IResult(Of InstaLoggedInChallengeDataInfo) = Await pApi.GetLoggedInChallengeDataInfoAsync
                                If ChallengeInfoRes.Succeeded Then
                                    ' Accept the login challenge 
                                    Dim AcceptChallengeRes As IResult(Of Boolean) = Await pApi.AcceptChallengeAsync()
                                    If AcceptChallengeRes.Succeeded Then
                                        If AcceptChallengeRes.Value Then
                                            Call LogConsole("[+] Login success (Accept Challenge flow)!", ConsoleColor.Green)

                                            ' Save session state in the Configuration object, then flush to disk
                                            pConfiguration.PreviousSessionJSON = pApi.GetStateDataAsString()
                                            Call pConfiguration.SaveConfiguration()
                                            Return True
                                        End If
                                    End If
                                End If

                                Call LogConsole("[-] Login temporarily disabled, wait and try again later.", ConsoleColor.Red)
                                Return False
                            Else
                                Call LogConsole("[-] Unknown login failure, wait and try again later.", ConsoleColor.Red)
                                Return False
                            End If
                        Else
                            Call LogConsole("[-] Unknown login failure, wait and try again later.", ConsoleColor.Red)
                            Return False
                        End If
                    Else
                        Call LogConsole("[-] Request for email challenge failed, wait and try again later.", ConsoleColor.Red)
                        Return False
                    End If
                Else
                    Call LogConsole("[-] Request for challenge type failed, wait and try again later.", ConsoleColor.Red)
                    Return False
                End If
            ElseIf LoginRes.Value = InstaLoginResult.LimitError Then
                Call LogConsole("[!] Login rate limit encountered, wait try again later...", ConsoleColor.Yellow)
                Return False
            ElseIf LoginRes.Value = InstaLoginResult.Exception Then
                Call LogConsole("[-] Login has been disabled for this account. Use the Instagram client to complete the Captcha.", ConsoleColor.Red)
                Return False
            Else
                Call LogConsole("[-] Unknown login failure, wait and try again later.", ConsoleColor.Red)
                Return False
            End If
        End If

        Return False
    End Function


    ''' <summary>
    ''' A synchronous version of LogoutAsync().
    ''' </summary>
    ''' <returns>Returns True if Logout succeeded, False otherwise.</returns>
    Public Function Logout() As Boolean
        Return LogoutAsync().ConfigureAwait(False).GetAwaiter().GetResult()
    End Function


    ''' <summary>
    ''' Logs the user out of the Instagram account. In doing so it invalidates any saved session state in the
    ''' Configuration object. Attempting to restore old session state to log back into the account will fail
    ''' after this call.
    ''' </summary>
    ''' <returns>True if the user was logged out successfully, False otherwise.</returns>
    Public Async Function LogoutAsync() As Task(Of Boolean)
        If (Await pApi.LogoutAsync()).Succeeded Then
            Return True
        End If

        Return False
    End Function


    ''' <summary>
    ''' A synchronous version of GetFollowingAsync().
    ''' </summary>
    ''' <returns>Returns a Collection(Of CompiledProfile) containing data about the accounts the logged in
    ''' user follows. If an error occurs it returns Nothing.</returns>
    Public Function GetFollowing() As Collection(Of CompiledProfile)
        Return GetFollowingAsync().ConfigureAwait(False).GetAwaiter().GetResult()
    End Function


    ''' <summary>
    ''' Queries the list of Followed users and enumerates through each of them to get profile data.
    ''' </summary>
    ''' <returns>Returns a Collection(Of CompiledProfile) containing data about the accounts the logged in
    ''' user follows. If an error occurs getting the Followed users list it returns Nothing, if an error occurs
    ''' trying to get data on an individual Followed user it will continue to try to retrieve that data and
    ''' will return the most amount of data available.</returns>
    Public Async Function GetFollowingAsync() As Task(Of Collection(Of CompiledProfile))

        ' Set the rate limit to the default until they start ratelimiting us, set the percentage displayed to
        ' less than zero to make sure the first console write is displayed
        Dim RateLimitDelayMilliseconds As Integer = DEFAULT_RATELIMIT_BACKOFF_PERIOD_IN_MILLISECONDS
        Dim LastPercentDisplayed As Integer = -1

        Call LogConsole("[*] Enumerating Followed Users...")

        ' Try to get up to 100 pages of users, then get the users the logged in user is following
        Dim Paging As InstagramApiSharp.PaginationParameters = InstagramApiSharp.PaginationParameters.MaxPagesToLoad(100)
        Dim FollowingResult As IResult(Of InstaUserShortList) = Await pApi.UserProcessor.GetUserFollowingAsync(pApi.GetLoggedUser().UserName, Paging)
        If FollowingResult.Succeeded Then
            ' Prep the list of profiles, a collection to return, and an HttpClient to download profile pictures
            Dim Following As InstaUserShortList = FollowingResult.Value
            Dim Profiles As New Collection(Of CompiledProfile)
            Dim HttpClient As New HttpClient

            If Not Following Is Nothing Then
                ' Loop through all of the users individually to get their profile information
                For i = 0 To Following.Count - 1
                    Dim UIResult As IResult(Of InstaUserInfo) = Await pApi.UserProcessor.GetUserInfoByIdAsync(Following(i).Pk)
                    If UIResult.Succeeded Then
                        ' If the call succeeded make sure we reset the backoff period if we were previously ratelimited
                        RateLimitDelayMilliseconds = DEFAULT_RATELIMIT_BACKOFF_PERIOD_IN_MILLISECONDS

                        ' Declare the profile structure to add to the collection and fill it out. Some of these
                        ' fields can be blank so take that into account.
                        Dim Profile As New CompiledProfile
                        Dim UserInfo As InstaUserInfo = UIResult.Value
                        Profile.IGID = UserInfo.Pk
                        Profile.Username = UserInfo.UserName
                        If UserInfo.FullName = "" Or UserInfo.FullName Is Nothing Then
                            Profile.FullName = ""
                        Else
                            Profile.FullName = UserInfo.FullName
                        End If
                        If UserInfo.CityName = "" Or UserInfo.CityName Is Nothing Then
                            Profile.Location = ""
                        Else
                            Profile.Location = UserInfo.CityName
                        End If
                        If UserInfo.ExternalUrl = "" Or UserInfo.ExternalUrl Is Nothing Then
                            Profile.ExternalUrl = ""
                        Else
                            Profile.ExternalUrl = UserInfo.ExternalUrl
                        End If
                        If UserInfo.Biography = "" Or UserInfo.Biography Is Nothing Then
                            Profile.Biography = ""
                        Else
                            Profile.Biography = UserInfo.Biography
                        End If

                        ' There are HD profile pictures and regular profile pictures. Try for the HD profile picture
                        ' first, then if that doesn't pan out get the smaller picture.
                        If UserInfo.HdProfilePicUrlInfo.Uri.ToString <> "" Then
                            Profile.ProfilePicUrl = UserInfo.HdProfilePicUrlInfo.Uri.ToString
                        Else
                            Profile.ProfilePicUrl = UserInfo.ProfilePicUrl
                        End If

                        ' If a URL was found for a picture download the picture
                        If Profile.ProfilePicUrl <> "" Then
                            Try
                                Profile.ProfilePicData = Await HttpClient.GetByteArrayAsync(Profile.ProfilePicUrl)
                            Catch ex As Exception
                                Call LogConsole("[-] Failure to download profile picture for " & Following(i).UserName & ": " & Profile.ProfilePicUrl, ConsoleColor.Red)
                            End Try
                        End If
                        Profile.CollectionDate = Now

                        Call Profiles.Add(Profile)

                        ' Introduce a random delay between profile downloads
                        Await Task.Delay(GetRandomDelayInMilliseconds(PROFILE_DOWNLOAD_MINIMUM_PAUSE_IN_MILLISECONDS, PROFILE_DOWNLOAD_MAXIMUM_PAUSE_IN_MILLISECONDS))
                    Else
                        ' For some reason the call to get user data was unsuccessful
                        If UIResult.Info.ResponseType = ResponseType.RequestsLimit Then
                            ' If we encounter the ratelimit on the next attempt it's not likely that waiting the same amount
                            ' of time is going to satisfly the ratelimit. Make it longer for in hopes that will work. If this is
                            ' the first time we're hitting the rate limit the lower limit will be DEFAULT_RATELIMIT_BACKOFF_PERIOD_IN_MILLISECONDS
                            ' and then the upper maximum is 1.5 times that constant. The floor rises every time we hit the rate limit without
                            ' a successful download in between and resets to the default once a success occurs.
                            RateLimitDelayMilliseconds = GetRandomDelayInMilliseconds(RateLimitDelayMilliseconds, DEFAULT_RATELIMIT_BACKOFF_PERIOD_IN_MILLISECONDS * 1.5)

                            ' We're being rate limited. Log the name of the user we were on and how long we wait (cut the decimal off too).
                            Call LogConsole("[!] Ratelimiting in GetUserInfoByUsernameAsync for " & Following(i).UserName & ", trying again in " & (CInt(RateLimitDelayMilliseconds / 1000)) & " seconds...", ConsoleColor.Yellow)

                            ' Back the iterator up by one to make sure we don't skip this user when we have the next go around
                            i -= 1

                            ' Delay for the wait time
                            Await Task.Delay(RateLimitDelayMilliseconds)
                        ElseIf UIResult.Info.ResponseType = ResponseType.NetworkProblem Or UIResult.Info.ResponseType = ResponseType.UnExpectedResponse Then
                            ' Something weird happened, do a quick gentleman's delay and then try again
                            Await Task.Delay(5)
                        Else
                            ' Something else weird happened, log it and move on
                            Call LogConsole("[-] Failure to download profile picture for " & Following(i).UserName & ": " & UIResult.Info.ResponseType, ConsoleColor.Red)
                        End If
                    End If

                    ' Log the percentage through user enumeration we are to the console at the intervals defined
                    Dim ProgressPercent As Integer = CInt(FormatNumber((CDbl(i / Following.Count) * 100), 0))
                    If ProgressPercent Mod PERCENTAGE_UPDATE_MODULUS = 0 And ProgressPercent <> LastPercentDisplayed Then
                        Call LogConsole("[*] Completed " & ProgressPercent & "%...")
                        LastPercentDisplayed = ProgressPercent
                    End If
                Next

                Call LogConsole("[*] Enumeration complete.")
                Return Profiles
            End If
        End If

        ' Enumerating the user failed, display status as to why that might have happened
        If FollowingResult.Info.ResponseType = ResponseType.ChallengeRequired Then
            ' If the account requires a challenge that's probably a phone number/SMS entry request. This process
            ' isn't completely clear yet.
            Call LogConsole("[-] GetFollowing failed: Challenge Required", ConsoleColor.Red)
        Else
            ' The call failed, log the repsonse type as to why to figure out later.
            Call LogConsole("[-] GetFollowing failed: " & FollowingResult.Info.ResponseType, ConsoleColor.Red)
        End If
        Return Nothing
    End Function


    ''' <summary>
    ''' This function processes the body text of the Instagram SMS two factor code and returns the challenge code.
    ''' </summary>
    ''' <param name="TextBody">The full body text of the Instagram SMS two factor message</param>
    ''' <returns>The return value is the 6 digit Two Factor Code.</returns>
    Private Function ProcessInstagram2FAText(ByVal TextBody As String) As String
        ' At the time of this writing, all Instagram Two Factor SMS messages start out as:
        ' ### ### is your Instagram code
        ' Where # above is one of the six digits to enter. This function removes all spaces (including the one
        ' between the first And second group of three digits) And then returns the first six characters of the
        ' message which should constitute the Two Factor code

        TextBody = TextBody.Replace(" ", "")
        Return TextBody.Substring(0, 6)
    End Function


    ''' <summary>
    ''' Returns a random number between a minimum and a maximum value.
    ''' </summary>
    ''' <param name="MinimumTime">The minimum value to return</param>
    ''' <param name="MaximumTime">The maximum value to return</param>
    ''' <returns>A random Integer between the two parameters passed to the function.</returns>
    Private Function GetRandomDelayInMilliseconds(ByVal MinimumTime As Integer, ByVal MaximumTime As Integer) As Integer
        Call Randomize()
        Return (Rnd() * (MaximumTime - MinimumTime + 1) + MinimumTime)
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
