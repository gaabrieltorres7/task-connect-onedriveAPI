using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;

class GraphHelper
{
    // Settings object
    private static Settings? _settings;
    // User auth token credential
    private static DeviceCodeCredential? _deviceCodeCredential;
    // Client configured with user authentication
    private static GraphServiceClient? _userClient;

    public static void InitializeGraphForUserAuth(Settings settings,
        Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
    {
        _settings = settings;

        _deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
            settings.TenantId, settings.ClientId);

        _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
    }

    public static async Task<string> GetUserTokenAsync()
    {
        // Ensure credential isn't null
        _ = _deviceCodeCredential ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        // Ensure scopes isn't null
        _ = _settings?.GraphUserScopes ?? throw new System.ArgumentNullException("Argument 'scopes' cannot be null");

        // Request token with given scopes
        var context = new TokenRequestContext(_settings.GraphUserScopes);
        var response = await _deviceCodeCredential.GetTokenAsync(context);
        return response.Token;
    }

    public static Task<User> GetUserAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me
            .Request()
            .Select(u => new
            {
                // Only request specific properties
                u.DisplayName,
                u.Mail,
                u.UserPrincipalName
            })
            .GetAsync();
    }

    public static Task<IMailFolderMessagesCollectionPage> GetInboxAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me
            // Only messages from Inbox folder
            .MailFolders["Inbox"]
            .Messages
            .Request()
            .Select(m => new
            {
                // Only request specific properties
                m.From,
                m.IsRead,
                m.ReceivedDateTime,
                m.Subject
            })
            // Get at most 25 results
            .Top(25)
            // Sort by received time, newest first
            .OrderBy("ReceivedDateTime DESC")
            .GetAsync();
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
            ToRecipients = new Recipient[]
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
            .SendMail(message)
            .Request()
            .PostAsync();
    }

    public async static Task<IDriveItemChildrenCollectionPage> MakeGraphCallAsync()
    {

        // Ensure client isn't null
        _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return await _userClient.Me.Drive.Root.Children.Request().GetAsync();
    }

    public async static Task<string> GetUserIdByEmail(string email)
    {
        try
        {
            var user = await _userClient.Users[email].Request().GetAsync();
            return user.Id;
        }
        catch (ServiceException ex)
        {
            Console.WriteLine("Error getting user ID: " + ex.Message);
            return null;
        }
    }

    public async static Task shareFolderWithSomeone(string fileId)
    {
        try
        {
            var recipients = new List<DriveRecipient>()
        {
            new DriveRecipient()
        {
            Email = "vsiqueira@seisistemas.com.br",
        }
    };
            var roles = new List<String>()
    {
        "write"
    };
            var message = "Testando link compartilhado!";
            var requireSignIn = true;
            var sendInvitation = true;


            await _userClient.Me.Drive.Items[fileId].Invite(
            recipients,
            requireSignIn,
            roles,
            sendInvitation,
            message)
            .Request()
            .PostAsync();

            Console.WriteLine("Success!");
        }
        catch (ServiceException ex)
        {
            Console.WriteLine("Erro grant to permission: " + ex.Message);
        }
    }

}