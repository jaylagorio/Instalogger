''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' File: Program.vb
''' Author: Jay Lagorio
''' Date Changed: 31OCT2020
''' Purpose: Drives the proof of concept. It loads the configuration from the JSON in the
''' configuration file.
''' 
''' 1. It gets the list of Followed users and queries the profile data from each in turn.
''' 2. Queries the database for changes between what it downloaded and what's stored in the
''' database and if there are differences it timestamps and stores them.
''' 3. If this is the first time the loop has gone through today it calculates the changes made
''' today and sends areport email using the configured server to the configured address.
''' 
''' Once all of this is done it waits a random amount of time between 
''' MINIMUM_DELAY_TIME_IN_MILLISECONDS and MAXIMUM_DELAY_TIME_IN_MILLISECONDS and then does it
''' all again.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports Instalogger.Database
Imports System.Collections.ObjectModel

Module Program
    ' Used to delay the thread for this amount of time between profile checks
    Private Const MINIMUM_DELAY_TIME_IN_MILLISECONDS As Long = 64800000 ' 18 hours
    Private Const MAXIMUM_DELAY_TIME_IN_MILLISECONDS As Long = 86400000 ' 24 hours

    ' Used to write updates to the console, but only after this increment of percentage
    Private Const PERCENTAGE_UPDATE_MODULUS As Integer = 25

    Sub Main(args As String())
        Console.ForegroundColor = ConsoleColor.Magenta
        Call Console.WriteLine()
        Call Console.WriteLine("  8888888                   888             888                                             ")
        Call Console.WriteLine("    888                     888             888                                             ")
        Call Console.WriteLine("    888                     888             888                                             ")
        Call Console.WriteLine("    888   88888b.  .d8888b  888888  8888b.  888  .d88b.   .d88b.   .d88b.   .d88b.  888d888 ")
        Call Console.WriteLine("    888   888 ""88b 88K      888        ""88b 888 d88""""88b d88P""88b d88P""88b d8P  Y8b 888P""   ")
        Call Console.WriteLine("    888   888  888 ""Y8888b. 888    .d888888 888 888  888 888  888 888  888 88888888 888     ")
        Call Console.WriteLine("    888   888  888      X88 Y88b.  888  888 888 Y88..88P Y88b 888 Y88b 888 Y8b.     888     ")
        Call Console.WriteLine("  8888888 888  888  88888P'  ""Y888 ""Y888888 888  ""Y88P""   ""Y88888  ""Y88888  ""Y8888  888     ")
        Call Console.WriteLine("                                                              888      888                  ")
        Call Console.WriteLine("                                                         Y8b d88P Y8b d88P                  ")
        Call Console.WriteLine("                                                          ""Y88P""   ""Y88P""              ")
        Call Console.WriteLine()
        Console.ForegroundColor = ConsoleColor.Gray

        ' Load the configuration from the JSON file. Includes all settings and may also include serialized login
        ' state from the Instagram library for stored login cookies, etc.
        Dim Configuration As Configuration = Configuration.LoadConfiguration()
        If Configuration Is Nothing Then
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("Configuration not loaded, config.json not present.")
            Exit Sub
        End If

        ' Instantiate SQL Server and Instagram interfaces with the loaded configuration
        Dim Database As New Database(Configuration.DatabaseString)

        ' Until the process is killed check for changes to Followed profiles, store in the database, generate a report, repeat
        While 1
            ' Attempt to log into Instagram
            Dim Instagram As New Instagram(Configuration)
            If Instagram.Login() Then
                ' Get the Followed profiles
                Dim Profiles As Collection(Of CompiledProfile) = Instagram.GetFollowing()
                If Not Profiles Is Nothing Then
                    Dim LastPercentDisplayed As Integer = -1
                    Call Console.WriteLine("[*] Updating database...")
                    For i = 0 To Profiles.Count - 1
                        ' Update the database based on the downloaded profile data
                        Call Database.UpdateProfile(Profiles(i))

                        ' Only show a progress percentage update every 25%
                        Dim ProgressPercent As Integer = CInt(FormatNumber((CDbl(i / Profiles.Count) * 100), 0))
                        If ProgressPercent Mod PERCENTAGE_UPDATE_MODULUS = 0 And ProgressPercent <> LastPercentDisplayed Then
                            Call Console.WriteLine("[*] Completed " & ProgressPercent & "%...")
                            LastPercentDisplayed = ProgressPercent
                        End If
                    Next

                    Call LogConsole("[+] Database successfully updated.", ConsoleColor.Green)

                    ' Generate the report of the changed profiles if a report hasn't been generated today
                    Dim Today As New DateTime(Now.Year, Now.Month, Now.Day)
                    Dim ReportDate As New DateTime(Configuration.LastReportingTime.Year, Configuration.LastReportingTime.Month, Configuration.LastReportingTime.Day)
                    If Configuration.LastReportingTime < Today Then
                        Dim Reporting As New Report(Configuration)
                        If Reporting.ReportChangedProfiles() Then
                            Configuration.LastReportingTime = Now
                            Call Configuration.SaveConfiguration()
                            Call LogConsole("[+] Database changes calculated and written.", ConsoleColor.Green)
                        Else
                            Call LogConsole("[-] Database change detection failed.", ConsoleColor.Red)
                        End If
                    Else
                        Call LogConsole("[*] Report last execution time: " & FormatNumber(Now.Subtract(Configuration.LastReportingTime).TotalHours, 2) & " hours ago.")
                    End If
                Else
                    Call Console.WriteLine("[-] Failed.")
                End If

                ' These don't go out of scope when we wait so destroying this frees up memory. We reinstantiate later.
                Profiles = Nothing
                Instagram = Nothing
            Else
                Exit While
            End If

            ' Wait a random time between the minimum wait time and maximum wait time and do it all again
            Dim TotalWaitTime As Long = GetRandomDelayInMilliseconds(MINIMUM_DELAY_TIME_IN_MILLISECONDS, MAXIMUM_DELAY_TIME_IN_MILLISECONDS)
            Dim NextRunTime As DateTime = Now.AddSeconds(TotalWaitTime / 1000)
            Call LogConsole("Waiting until " & NextRunTime & " for the next cycle...", ConsoleColor.Green)
            While TotalWaitTime > 0
                Dim WaitTime As UInteger
                If TotalWaitTime > Integer.MaxValue Then
                    WaitTime = UInteger.MaxValue
                Else
                    WaitTime = TotalWaitTime
                End If
                Call Threading.Thread.Sleep(WaitTime)
                TotalWaitTime -= WaitTime
            End While
        End While
    End Sub


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
End Module
