# Office365-Resume-Builder-Addin
This is a sample project and requires a change in clientcontext url and credential. 
### Download the solution change the url in this line
```C#
using (ClientContext context = new ClientContext("https://contoso.sharepoint.com"))
### Password and Admin ID in the below line
foreach (char c in ("P@ssW0rd!").ToCharArray()) Password.AppendChar(c);
context.Credentials = new SharePointOnlineCredentials("admin@contoso.onmicrosoft.com", Password);
```
# Sample Output
[![mutt dark](https://github.com/ChendrayanV/Office365-Resume-Builder-Addin/blob/master/ResumeBuilder/Images/2016-10-21_12-19-12.png)](https://github.com/ChendrayanV/Office365-Resume-Builder-Addin/blob/master/ResumeBuilder/Images/2016-10-21_12-19-12.png))

# Notes
This is very basic code and to proper exception handling in place and the information are retrieved from SharePoint Online User Profile Properties. Do check the code and append your
feedback.
### Work In Progress
1. Export to PDF. 
2. Image manipulation.
3. More styling 
4. One Pager and Detailed resume.     