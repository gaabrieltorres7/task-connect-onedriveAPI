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

    public async static Task CreateFolder(string emailID)
    {

        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        //Creating folder
        var folder = new DriveItem
        {
            Name = "folder test",
            Folder = new Folder(),
            AdditionalData = new Dictionary<string, object>
                {
                    { "@microsoft.graph.conflictBehavior", "rename" }
                }
        };

        try
        {
            var createdFolder = await _userClient
                .Drives[emailID]
                .Items
                .Request()
                .AddAsync(folder);

            Console.WriteLine("Folder created successfully");
            Console.WriteLine($"Folder ID: {createdFolder.Id}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating subfolder: {ex.Message}");
        }


    }

    public async static Task CreateSubFolder()
    {

        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");


        var userID = "818f2a17-095d-4e9f-887c-fcd535cf9445";
        var parentFolderId = "01P4IMMET4XJMVXGRL2FFLTWXFCJ3KJUZK"; //POWER APPS
        //Creating folder
        var subfolder = new DriveItem
        {
            Name = "subfolder test",
            Folder = new Folder(),
            AdditionalData = new Dictionary<string, object>
                {
                    { "@microsoft.graph.conflictBehavior", "rename" }
                }
        };

        try
        {
            var createdSubfolder = await _userClient.Drives[userID].Items[parentFolderId].Children.Request().AddAsync(subfolder);

            Console.WriteLine("Subfolder created successfully");
            Console.WriteLine($"Subfolder ID: {createdSubfolder.Id}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating subfolder: {ex.Message}");
        }


    }

    public async static Task CreateASharingLink()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        //var emailAddress = "gtorres@seisistemas.com.br";
        // Create a sharing link for the item
        var link = new Permission
        {
            Link = new SharingLink
            {
                Type = "view",
                Scope = "organization"
            },
            Roles = new List<string> { "read" }
        };

        try
        {
            //testando folder
            var response = await _userClient.Me.Drive.Items["01P4IMMERKF6RKTOF6ZZF3TGDM6RHKEG3Q"].CreateLink("view", "organization").Request().PostAsync();
            // The response will contain the created sharing link, which you can use to share the item with others
            var sharingLink = response.Link.WebUrl;
            Console.WriteLine(sharingLink);

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating link: {ex.Message}");
        }

    }

}