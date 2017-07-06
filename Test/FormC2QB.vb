
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
        config.ApplicationToken = ConfigurationManager.AppSettings("applicationToken")
        config.ConsumerKey = ConfigurationManager.AppSettings("consumerKey")
        config.ConsumerSecret = ConfigurationManager.AppSettings("consumerSecret")
        config.OauthRequestTokenEndpoint = ConfigurationManager.AppSettings("oauthRequestTokenEndpoint")
        config.OauthAccessTokenEndpoint = ConfigurationManager.AppSettings("oauthAccessTokenEndpoint")
        config.OauthBaseUrl = ConfigurationManager.AppSettings("oauthBaseUrl")
        config.OauthUserAuthUrl = ConfigurationManager.AppSettings("oauthUserAuthUrl")
        config.SelectedMode = QbConfig.AppMode.Console
        Dim result As QbResponse = Nothing
        Dim connectToQuickBooks As New QbConnect()
        connectToQuickBooks.IConnect_Connect(config, result)
    End Sub

End Class


