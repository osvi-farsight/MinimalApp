using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Windows.Data.Xml.Dom;
using Windows.System;

using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Me.SendMail;
using User = Microsoft.Graph.Models.User;

namespace MinimalApp
{
    internal class GraphHelper
    {
        // Settings object
        private static Settings? _settings;
        // User auth token credential
        private static DeviceCodeCredential? _deviceCodeCredential;
        // Client configured with user authentication
        private static GraphServiceClient? _userClient;

        private static string? _token;

        public static void InitializeGraphForUserAuth(Settings settings,
            Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
        {
            _settings = settings;

            var options = new DeviceCodeCredentialOptions
            {
                ClientId = settings.ClientId,
                TenantId = settings.TenantId,
                DeviceCodeCallback = deviceCodePrompt,
            };

            _deviceCodeCredential = new DeviceCodeCredential(options);

            _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
        }

        public static async Task<string> GetUserTokenAsync()
        {
            if (_token != null) return _token;

            // Ensure credential isn't null
            _ = _deviceCodeCredential ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            // Ensure scopes isn't null
            _ = _settings?.GraphUserScopes ?? throw new System.ArgumentNullException("Argument 'scopes' cannot be null");

            // Request token with given scopes
            var context = new TokenRequestContext(_settings.GraphUserScopes);
            var response = await _deviceCodeCredential.GetTokenAsync(context);
            _token = response.Token;
            return response.Token;
        }

        public static Task<User?> GetUserAsync()
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            return _userClient.Me.GetAsync((config) =>
            {
                // Only request specific properties
                config.QueryParameters.Select = new[] { "displayName", "mail", "userPrincipalName" };
                if (_token != null) config.Headers.Add("Authorization", $"bearer {_token}");
            });
        }

        public static Task<MessageCollectionResponse?> GetInboxAsync()
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            return _userClient.Me
                // Only messages from Inbox folder
                .MailFolders["Inbox"]
                .Messages
                .GetAsync((config) =>
                {
                    // Only request specific properties
                    config.QueryParameters.Select = new[] { "from", "isRead", "receivedDateTime", "subject" };
                    // Get at most 25 results
                    config.QueryParameters.Top = 25;
                    // Sort by received time, newest first
                    config.QueryParameters.Orderby = new[] { "receivedDateTime DESC" };
                    if (_token != null) config.Headers.Add("Authorization", $"bearer {_token}");
                });
        }

        public static async Task SendMailAsync(string subject, string body, string recipient)
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            // Create a new message
            var message = new Message
            {
                Subject = subject,
                Body = new ItemBody
                {
                    Content = body,
                    ContentType = BodyType.Text
                },
                ToRecipients = new List<Recipient>
                {
                new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = recipient
                    }
                }
            }
            };

            // Send the message
            await _userClient.Me
                .SendMail
                .PostAsync(new SendMailPostRequestBody
                {
                    Message = message
                }, (config) =>
                {
                    if (_token != null) config.Headers.Add("Authorization", $"bearer {_token}");
                });
        }
    }
}
