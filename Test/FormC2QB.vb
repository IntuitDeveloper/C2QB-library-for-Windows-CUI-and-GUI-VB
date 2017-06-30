
Imports System.ComponentModel
Imports System.Configuration

Public Class FormC2QB

    Private Sub C2QB_Click(sender As Object, e As EventArgs) Handles C2QB.Click
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object,
                                         ByVal e As System.ComponentModel.DoWorkEventArgs) _
                                         Handles BackgroundWorker1.DoWork
        ' Do some time-consuming work on this thread.
        System.Threading.Thread.Sleep(1000)

        Dim config As New QbConfig()
        config.ApplicationToken = ConfigurationSettings.AppSettings("applicationToken")
        config.ConsumerKey = ConfigurationSettings.AppSettings("consumerKey")
        config.ConsumerSecret = ConfigurationSettings.AppSettings("consumerSecret")
        config.OauthRequestTokenEndpoint = ConfigurationSettings.AppSettings("oauthRequestTokenEndpoint")
        config.OauthAccessTokenEndpoint = ConfigurationSettings.AppSettings("oauthAccessTokenEndpoint")
        config.OauthBaseUrl = ConfigurationSettings.AppSettings("oauthBaseUrl")
        config.OauthUserAuthUrl = ConfigurationSettings.AppSettings("oauthUserAuthUrl")
        config.SelectedMode = QbConfig.AppMode.Console
        Dim result As QbResponse = Nothing
        Dim connectToQuickBooks As New QbConnect()
        connectToQuickBooks.IConnect_Connect(config, result)
    End Sub

End Class


