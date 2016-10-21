# Office365-Resume-Builder-Addin
This is a sample project and requires a change in clientcontext url and credential. 
```C#
# Download the solution change the url in this line
 using (ClientContext context = new ClientContext("https://contoso.sharepoint.com"))
# Password and Admin ID in he below line
foreach (char c in ("P@ssW0rd!").ToCharArray()) Password.AppendChar(c);
context.Credentials = new SharePointOnlineCredentials("admin@contoso.onmicrosoft.com", Password);

# Sample Output
[![mutt dark](https://github.com/ChendrayanV/Office365-Resume-Builder-Addin/blob/master/ResumeBuilder/Images/2016-10-21_12-19-12.png)](https://github.com/ChendrayanV/Office365-Resume-Builder-Addin/blob/master/ResumeBuilder/Images/2016-10-21_12-19-12.png))