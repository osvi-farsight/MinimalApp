using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Foundation.Collections;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace MinimalApp
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        Settings settings = Settings.LoadSettings();

        public MainWindow()
        {
            this.InitializeComponent();
        }

        private void myButton_Click(object sender, RoutedEventArgs e)
        {
            myButton.Content = "Clicked";
        }

        private void ButtonGreetUserAsync_Click(object sender, RoutedEventArgs e)
        {
            _ = GreetUserAsync();
        }

        private void ButtonDisplayAccessTokenAsync_Click(object sender, RoutedEventArgs e)
        {
            _ = DisplayAccessTokenAsync();
        }

        private void ButtonListInboxAsync_Click(object sender, RoutedEventArgs e)
        {
            _ = ListInboxAsync();
        }

        private void ButtonSendMailAsync_Click(object sender, RoutedEventArgs e)
        {
            _ = SendMailAsync();
        }

        void InitializeGraph(Settings settings)
        {
            GraphHelper.InitializeGraphForUserAuth(settings,
                (info, cancel) =>
                {
                    // Display the device code message to
                    // the user. This tells them
                    // where to go to sign in and provides the
                    // code to use.
                    Debug.WriteLine(info.Message);
                    return Task.FromResult(0);
                });
        }

        async Task GreetUserAsync()
        {
            try
            {
                var user = await GraphHelper.GetUserAsync();
                Debug.WriteLine($"Hello, {user?.DisplayName}!");
                // For Work/school accounts, email is in Mail property
                // Personal accounts, email is in UserPrincipalName
                Debug.WriteLine($"Email: {user?.Mail ?? user?.UserPrincipalName ?? ""}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting user: {ex.Message}");
            }
        }

        async Task DisplayAccessTokenAsync()
        {
            try
            {
                var userToken = await GraphHelper.GetUserTokenAsync();
                Debug.WriteLine($"User token: {userToken}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting user access token: {ex.Message}");
            }
        }

        async Task ListInboxAsync()
        {
            try
            {
                var messagePage = await GraphHelper.GetInboxAsync();

                if (messagePage?.Value == null)
                {
                    Debug.WriteLine("No results returned.");
                    return;
                }

                // Output each message's details
                foreach (var message in messagePage.Value)
                {
                    Debug.WriteLine($"Message: {message.Subject ?? "NO SUBJECT"}");
                    Debug.WriteLine($"  From: {message.From?.EmailAddress?.Name}");
                    Debug.WriteLine($"  Status: {(message.IsRead!.Value ? "Read" : "Unread")}");
                    Debug.WriteLine($"  Received: {message.ReceivedDateTime?.ToLocalTime().ToString()}");
                }

                // If NextPageRequest is not null, there are more messages
                // available on the server
                // Access the next page like:
                // var nextPageRequest = new MessagesRequestBuilder(messagePage.OdataNextLink, _userClient.RequestAdapter);
                // var nextPage = await nextPageRequest.GetAsync();
                var moreAvailable = !string.IsNullOrEmpty(messagePage.OdataNextLink);

                Debug.WriteLine($"\nMore messages available? {moreAvailable}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting user's inbox: {ex.Message}");
            }
        }

        async Task SendMailAsync()
        {
            try
            {
                // Send mail to the signed-in user
                // Get the user for their email address
                var user = await GraphHelper.GetUserAsync();

                var userEmail = user?.Mail ?? user?.UserPrincipalName;

                if (string.IsNullOrEmpty(userEmail))
                {
                    Debug.WriteLine("Couldn't get your email address, canceling...");
                    return;
                }

                await GraphHelper.SendMailAsync("Testing Microsoft Graph",
                    "Hello world!", userEmail);

                Debug.WriteLine("Mail sent.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error sending mail: {ex.Message}");
            }
        }
    }
}
