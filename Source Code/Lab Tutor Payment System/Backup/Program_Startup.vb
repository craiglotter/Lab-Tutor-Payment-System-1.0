Module Program_Startup
    Sub main()
        Try
            Dim main_screen As New Main_Screen
            main_screen.ShowDialog()
            main_screen.Dispose()
            Application.Exit()
        Catch ex As Exception
            MsgBox("Sorry, but it would appear that Lab Tutor Payment System has encountered an error situation. The following error has been encountered: " & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Lab Tutor Payment System")
            Application.Exit()
        End Try
    End Sub
End Module
