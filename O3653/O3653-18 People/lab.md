#Microsoft Graph People API

In this lab, you will use Microsoft Graph to integrate the Office365 people api’s with an ASP.NET MVC application.

## Prerequisites
  1. You must have an Office 365 tenant and Microsoft Azure subscription to complete this lab. If you do not have one, the lab for O3651-7 Setting up your Developer environment in Office 365 shows you how to obtain a trial.
  2. You must have Visual Studio 2015 with Update 1 installed.
  3. You must have the Graph AAD Auth v2 Started Project template installed.
  4. You must have the OData v4 Client Code Generator addin installed.

## Exercise 1: Create a new project using Azure Active Directory v2 authentication

In this first step, you will create a new ASP.NET MVC project using the
**Graph AAD Auth v2 Start Project** template, register a new application
in the developer portal, and log in to your app and generate access tokens
for calling the Graph API.

1. Launch Visual Studio 2015 and select **New**, **Project**.
  1. Search the installed templates for **Graph** and select the **Graph AAD Auth v2 Starter Project** template.
  2. Name the new project **PeopleGraphWeb** and click **OK**.
  3. Open the **Web.config** file and find the **appSettings** element. This is where you will need to add your appId and app secret you will generate in the next step.
2. Launch the [Application Registration Portal](https://apps.dev.microsoft.com) to register a new application.
      1. Sign into the portal using your Office 365 username and password.
      2. Click **Add an App** and type **Graph Files Quick Start** for the application name.
      3. Copy the **Application Id** and paste it into the value for **ida:AppId** in your project's **web.config** file.
      3. Under **Application Secrets** click **Generate New Password** to create a new client secret for your app.
      4. Copy the displayed app password and paste it into the value for **ida:AppSecret** in your project's **web.config** file.
      5. Modify the **ida:AppScopes** value to include the required **https://graph.microsoft.com/files.readwrite** scope.

```xml
<configuration>
  <appSettings>
    <!-- ... -->
    <add key="ida:AppId" value="paste application id here" />
    <add key="ida:AppSecret" value="paste application password here" />
    <!-- ... -->
    <!-- Specify scopes in this value. Multiple values should be comma separated. -->
    <add key="ida:AppScopes" value="https://graph.microsoft.com/people.read" />
  </appSettings>
  <!-- ... -->
</configuration>
```
3. Add a redirect URL to enable testing on your localhost.
   1. Right click on **PeopleGraphWeb** and click on **Properties** to open the project properties.
   2. Click on **Web** in the left navigation.
   3. Copy the **Project Url** value.
   4. Back on the Application Registration Portal page, click **Add Platform** and then **Web**.
   5. Paste the value of **Project Url** into the **Redirect URIs** field.
   6. Scroll to the bottom of the page and click **Save**.

4. Press F5 to compile and launch your new application in the default browser.
   1. Once the Graph and AAD v2 Auth Endpoint Starter page appears, click **Sign in** and login to your Office 365 account.
   2. Review the permissions the application is requesting, and click **Accept**.
   3. Now that you are signed into your application, exercise 1 is complete!

## Exercise 2: Add a reference to the Graph API beta namespace

1. Right-click the project and select **add item**
   1. Select Visual C# -> Code -> Odata Client
   2. Name the file Graph.tt and click Add
2. Edit the Graph.tt file
   1. Edit MetadataDocumentUri to be "https://graph.microsoft.com/beta/$metadata"
   2. Edit NamespacePrefix to be "PeopleGraphWeb.Service"
3. Build the project


## Exercise 3: Add the people controller and call the people api
1. Right click the **Controllers** folder and select **Add...** -> Controller... 
   1. Select **MVC5 Controller - Empty** and click Add
   2. Name the controller **PeopleController** and click Add
2. Edit the usings:
```C#
using System;
using System.Web.Mvc;
using PeopleGraphWeb.TokenStorage;
using System.Configuration;
using System.Threading.Tasks;
using System.Security.Claims;
using PeopleGraphWeb.Auth;
using System.Linq;
using System.Collections.Generic;
```

3. Add the following helper functions to the PeopleController class:

```C#
        public async Task<string> GetToken()
        {
            string userObjId = AuthHelper.GetUserId(ClaimsPrincipal.Current);

            SessionTokenCache tokenCache = new SessionTokenCache(userObjId, HttpContext);

            string authority = string.Format(ConfigurationManager.AppSettings["ida:AADInstance"], "common/", "");

            AuthHelper authHelper = new AuthHelper(authority, ConfigurationManager.AppSettings["ida:AppId"],
               ConfigurationManager.AppSettings["ida:AppSecret"], tokenCache);

            return await authHelper.GetUserAccessToken(Url.Action("Index", "Home", null, Request.Url.Scheme));
        }

        public async Task<Service.GraphService> GetService()
        {
            string token = await GetToken();
            Service.GraphService service = new Service.GraphService(new Uri("https://graph.microsoft.com/beta/"));
            service.BuildingRequest += (sender, e) => e.Headers.Add("Authorization", "Bearer " + token);
            return service;
        }

        private IEnumerable<T> Search<T>(Microsoft.OData.Client.DataServiceContext dataServiceContext, Microsoft.OData.Client.DataServiceQuery<T> path, string searchString)
        {
            return dataServiceContext.Execute<T>(new Uri(path.RequestUri, "?$search=\"" + searchString + "\""));
        }
```

4. Add the index action that will list the relevant people for the logged in user

```c#
        [Authorize]
        public async Task<ActionResult> Index()
        {
            var service = await GetService();
            return View(service.Me.People.ToList());
        }
```

5. Add the view for the index controller 
   1. Right-click the views folder and click add -> new folder
   2. Rename the folder to People
   3. Right-click People and select **Add** -> **View...**
   4. Name the view "Index" and select Template "Empty"
   5. Set the contents of the file to the following:

```asp
@model IEnumerable<PeopleGraphWeb.Service.Person>

@{
    ViewBag.Title = "People";
}

<table class="table">
    <tr>
        <th>
            @Html.DisplayName("Display Name");
        </th>
        <th></th>
    </tr>
    
@foreach (var item in Model.People) {
    <tr>
        <td>
            @Html.DisplayFor(modelItem => item.DisplayName)
        </td>
        <td>
            @Html.ActionLink("Details", "Details", new { id=item.Id }) 
        </td>
    </tr>
}
    
</table>

```

### Edit the default layout to point to our new controller
1. Locate the **Views/Shared** folder in the project.
2. Open the **_Layout.cshtml** file found in the **Views/Shared** folder.
    1. Locate the part of the file that includes a few links at the top of the
       page. It should look similar to the following code:

```asp
  <ul class="nav navbar-nav">
      <li>@Html.ActionLink("Home", "Index", "Home")</li>
      <li>@Html.ActionLink("About", "About", "Home")</li>
      <li>@Html.ActionLink("Contact", "Contact", "Home")</li>
      <li>@Html.ActionLink("Graph API", "Graph", "Home")</li>
  </ul>
```

    2. Update that navigation to replace the "Graph API" link with "People"
       and connect this to the controller you just created.

```asp
  <ul class="nav navbar-nav">
      <li>@Html.ActionLink("Home", "Index", "Home")</li>
      <li>@Html.ActionLink("About", "About", "Home")</li>
      <li>@Html.ActionLink("Contact", "Contact", "Home")</li>
      <li>@Html.ActionLink("People", "Index", "People")</li>
  </ul>
```

### Verify the project works
1. In **Visual Studio**, hit **F5** to begin debugging.
2. When prompted, log in with your Office 365 Account.
3. Click the link **People** on the top of the home page.
4. Verify that your application displays the top relevant people for the current logged in user

## Exercise 4: Add support for people search and the details page
1. Add the search and details controllers:
```c#
        [Authorize]
        public async Task<ActionResult> Search(string searchText, string topic)
        {
            var service = await GetService();
            var searchString = string.IsNullOrWhiteSpace(topic) ? searchText : searchText + " topic:" + topic;
            return View("Index", Search(service, service.Me.People, searchString));
        }

        [Authorize]
        public async Task<ActionResult> Details(string id)
        {
            var service = await GetService();
            return View(service.Me.People.Where(p => p.Id == id).Single());
        }
```

2. Add the details view and update the index view to support search
   1. Right-click on the People folder under views and select Add -> View… 
      1. Set **View Name** to "Details"
      2. Set **Template** to "Details"
      3. Set **Model Class** to "Person (PersonGraphWeb.Service)"
   2. Edit the **Views/People/Index.cshtml**
      1. Locate the table element and add the following code right above it:

```asp
@{ using (Html.BeginForm("Search", "People", FormMethod.Get))
    {
        @Html.Label("Seach:")
        @Html.TextBox("searchText")
        @Html.Label("Topic:")
        @Html.TextBox("topic")
        <input type="submit" value="Search" />
    }
}
```
This will allow the user to enter search strings that will be passed to the search controller

3. Verify the search and details features work
  1. In **Visual Studio**, hit **F5** to begin debugging.
  2. When prompted, log in with your Office 365 Account.
  3. Click the link **People** on the top of the home page.
  4. Verify that your application displays the top relevant people for the current logged in user
  5. Click details to and verify the details are shown
  6. Go back to the index and enter a search term into the search field then click Search
    For example: 
      * Search with the text: "Dennis Dehin" and see the fuzzy matched result Denis Dehenne is returned
      * Search with the text: "Azis Hasoneh" and see the fuzzy matched result Aziz Hassouneh is returned
      * Search with the topic: XT2000

## Exercise 5: Add support for working with
1. Edit the **web.config** file to request the "https://graph.microsoft.com/user.readbasic.all" scope, it should now look like:
```xml
<configuration>
  <appSettings>
    <!-- ... -->
    <add key="ida:AppScopes" value="https://graph.microsoft.com/people.read,https://graph.microsoft.com/user.readbasic.all" />
  </appSettings>
  <!-- ... -->
</configuration>
```
2. Add the following method to the people controller:
```c#
        [Authorize]
        public async Task<ActionResult> RelatedPeople(string id)
        {
            var service = await GetService();
            return View("Index", service.Users.Where(p => p.Id == id).Single().People);
        }
```
Notice the code re-uses the index view to display the results so another view is not needed

3. Edit the People index view and add a new column to the table that links to the related people action:
```asp
        <td>
            @Html.ActionLink("Related People", "RelatedPeople", new { id=item.Id }) 
        </td> 
```
4. Verify the search and details features work
   1. In **Visual Studio**, hit **F5** to begin debugging.
   2. When prompted, log in with your Office 365 Account.
   3. Click the link **People** on the top of the home page.
   4. Selected **RelatedPeople** for a user and verify the related contacts are shown
