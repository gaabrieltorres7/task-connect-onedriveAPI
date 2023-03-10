var settings = Settings.LoadSettings();

// Initialize Graph
InitializeGraph(settings);

// Greet the user by name
await GreetUserAsync();

int choice = -1;

while (choice != 0)
{
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Display access token");
    Console.WriteLine("2. List my inbox");
    Console.WriteLine("3. Send mail");
    Console.WriteLine("4. List my folders");
    Console.WriteLine("5. Get email ID");
    Console.WriteLine("6. Share folder with someone");
    Console.WriteLine("7. Create folder");
    Console.WriteLine("8. Create Subfolder");
    Console.WriteLine("9. Create a sharing link for a driveItem");

    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (System.FormatException)
    {
        // Set to invalid value
        choice = -1;
    }

    switch (choice)
    {
        case 0:
            // Exit the program
            Console.WriteLine("Goodbye...");
            break;
        case 1:
            // Display access token
            await DisplayAccessTokenAsync();
            break;
        case 2:
            // List emails from user's inbox
            await ListInboxAsync();
            break;
        case 3:
            // Send an email message
            await SendMailAsync();
            break;
        case 4:
            // Run any Graph code
            await MakeGraphCallAsync();
            break;
        case 5:
            // Run any Graph code
            await GetUserIdByEmail();
            break;
        case 6:
            // Run any Graph code
            await shareFolderWithSomeone();
            break;
        case 7:
            // Run any Graph code
            await CreateFolder();
            break;
        case 8:
            // Run any Graph code
            await CreateSubFolder();
            break;
        case 9:
            // Run any Graph code
            await CreateASharingLink();
            break;
        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }
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
            Console.WriteLine(info.Message);
            return Task.FromResult(0);
        });
}

async Task GreetUserAsync()
{
    try
    {
        var user = await GraphHelper.GetUserAsync();
        Console.WriteLine($"Hello, {user?.DisplayName}!");
        // For Work/school accounts, email is in Mail property
        // Personal accounts, email is in UserPrincipalName
        Console.WriteLine($"Email: {user?.Mail ?? user?.UserPrincipalName ?? ""}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user: {ex.Message}");
    }
}

async Task DisplayAccessTokenAsync()
{
    try
    {
        var userToken = await GraphHelper.GetUserTokenAsync();
        Console.WriteLine($"User token: {userToken}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user access token: {ex.Message}");
    }
}

async Task ListInboxAsync()
{
    try
    {
        var messagePage = await GraphHelper.GetInboxAsync();

        // Output each message's details
        foreach (var message in messagePage.CurrentPage)
        {
            Console.WriteLine($"Message: {message.Subject ?? "NO SUBJECT"}");
            Console.WriteLine($"  From: {message.From?.EmailAddress?.Name}");
            Console.WriteLine($"  Status: {(message.IsRead!.Value ? "Read" : "Unread")}");
            Console.WriteLine($"  Received: {message.ReceivedDateTime?.ToLocalTime().ToString()}");
        }

        // If NextPageRequest is not null, there are more messages
        // available on the server
        // Access the next page like:
        // messagePage.NextPageRequest.GetAsync();
        var moreAvailable = messagePage.NextPageRequest != null;

        Console.WriteLine($"\nMore messages available? {moreAvailable}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user's inbox: {ex.Message}");
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
            Console.WriteLine("Couldn't get your email address, canceling...");
            return;
        }

        await GraphHelper.SendMailAsync("Testing Microsoft Graph",
            "Hello world!", userEmail);

        Console.WriteLine("Mail sent.");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error sending mail: {ex.Message}");
    }
}

async Task MakeGraphCallAsync()
{
    try
    {
        var items = await GraphHelper.MakeGraphCallAsync();
        foreach (var item in items.CurrentPage)
        {
            Console.WriteLine("\n==========================");
            Console.WriteLine(item.Id);
            Console.WriteLine(item.Name);
            Console.WriteLine(item.Size);
            Console.WriteLine(item.LastModifiedDateTime);
            Console.WriteLine("==========================\n");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting folders: {ex.Message}");
    }
}

async Task<string> GetUserIdByEmail()
{
    var emailID = await GraphHelper.GetUserIdByEmail("gtorres@seisistemas.com.br");
    Console.WriteLine(emailID);
    return emailID;
}

async Task shareFolderWithSomeone()
{
    await GraphHelper.shareFolderWithSomeone("01P4IMMEXS4NEHSACETVEJXT34IJF2NF4W");

}

async Task CreateFolder()
{
    await GraphHelper.CreateFolder("818f2a17-095d-4e9f-887c-fcd535cf9445");

}

async Task CreateSubFolder()
{
    await GraphHelper.CreateSubFolder();

}

async Task CreateASharingLink()
{
    await GraphHelper.CreateASharingLink();
}