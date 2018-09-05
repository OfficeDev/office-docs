# Build an Outlook add-in

Outlook add-ins are web applications built using standard web technologies and loaded within the Outlook client. In this hands-on lab, you will use our new JavaScript APIs to build an event-driven room booking add-in. The add-in that you build will:

- Be launched when the user clicks a button in the Outlook ribbon.
- Run in a task pane that's displayed to the right of an appointment in compose mode.
- Display the email address of the appointment organizer.
- Alert the user when recipients are changed.
- Alert the user when appointment time is changed.

## Install prerequisites

**IMPORTANT** Do this as part of VM image. (Remove from lab)

Begin this lab by installing the tools that you'll use to create your add-in project: [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install the latest version of these tools globally, open the command prompt and run the following command.

```
npm install -g yo generator-office
```

## Create the add-in project

Next, complete the following steps to create the add-in project by using the **Yeoman generator for Office Add-ins**.

1. Create a folder on your local drive and name it `my-outlook-addin`. This is where you'll create the files for your add-in.

1. Navigate to your new folder by running the following command from teh command prompt.

    ```
    cd my-outlook-addin
    ```

1. Use the Yeoman generator to create an Outlook Add-in project. Run the following command from the command prompt and then answer the prompts as shown below:

    ```
    yo office
    ```

    - **Choose a project type:** `Office Add-in project using Jquery framework`
    - **Choose a script type:** `Typescript`
    - **What do you want to name your add-in?:** `My Outlook Add-in`
    - **Which Office client application would you like to support?:** `Outlook`
    
    ![A screenshot of the prompts and answers for the Yeoman generator](images/quick-start-yo-prompts.PNG)
    
    After you complete the wizard, the generator will create the project and install supporting Node components.

## Update the code

At this point, the **Yeoman generator for Office Add-ins** has created a very basic add-in project that you can use as a starting point for building your Outlook add-in. Update the code as described in this section to customize the functionality of your add-in.

### Step 1: Customize the Manifest

The add-in manifest define's the add-in's settings and capabilities. 

1. Open the file **my-outlook-add-in-manifest.xml** file.

1. The `ProviderName` element has a placeholder value. Replace it with your name.

1. The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with `Room Validator`.

1. The `DefaultValue` attribute of the `SupportUrl` element has a placeholder. Replace it with `https://localhost:3000` and save the file.

    ```xml
    ...
    <ProviderName>Jane Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Outlook Add-in" />
    <Description DefaultValue="Room Validator Add-in"/>

    <!-- Icon for your add-in. Used on installation screens and the add-ins dialog. -->
    <IconUrl DefaultValue="https://localhost:3000/assets/icon-32.png" />
    <HighResolutionIconUrl DefaultValue="https://localhost:3000/assets/hi-res-icon.png"/>

    <!--If you plan to submit this add-in to the Office Store, uncomment the SupportUrl element below-->
    <SupportUrl DefaultValue="https://localhost:3000" />
    ...
    ```

1. Replace the entire contents of the `ExtensionPoint` element with the following XML markup. TODO: Change the type of extension point, since we want the button to be displayed on the ribbon in appointment compose mode. (By default, Yo Office creates a "Message Read" add-in.)

    ```xml
    <!-- Appointment Organizer -->
    <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
       <!-- Use the default tab of the ExtensionPoint or create your own with <CustomTab id="myTab"> -->
      <OfficeTab id="TabDefault">
        <!-- Up to 6 Groups added per Tab -->
        <Group id="apptComposeGroup">
          <Label resid="groupLabel" />
          <!-- Launch the add-in : task pane button -->
          <Control xsi:type="Button" id="apptComposeOpenPaneButton">
            <Label resid="apptComposeButtonLabel" />
            <Supertip>
              <Title resid="apptComposeSuperTipTitle" />
              <Description resid="apptComposeSuperTipDescription" />
            </Supertip>
            <Icon>
              <bt:Image size="16" resid="icon16" />
              <bt:Image size="32" resid="icon32" />
              <bt:Image size="80" resid="icon80" />
            </Icon>
            <Action xsi:type="ShowTaskpane">
              <SourceLocation resid="apptComposeTaskPaneUrl" />
            </Action>
          </Control>
          <!-- Go to http://aka.ms/ButtonCommands to learn how to add more Controls: ExecuteFunction and Menu -->
        </Group>
      </OfficeTab>
    </ExtensionPoint>    
    ```

1. Replace the entire contents of the `Resources` element with the following XML markup.

    ```xml
    <Resources>
      <bt:Images>
        <bt:Image id="icon16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="icon32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="icon80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="functionFile" DefaultValue="https://localhost:3000/function-file/function-file.html"/>
        <bt:Url id="apptComposeTaskPaneUrl" DefaultValue="https://localhost:3000/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="groupLabel" DefaultValue="My Add-in Group"/>
        <bt:String id="customTabLabel"  DefaultValue="My Add-in Tab"/>
        <bt:String id="apptComposeButtonLabel" DefaultValue="Room Validator"/>
        <bt:String id="apptComposeSuperTipTitle" DefaultValue="Validate the choice of meeting room"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="apptComposeSuperTipDescription" DefaultValue="Opens a pane which validates that the selected meeting room is available at the chosen time and can accommodate the number of invited attendees."/>
      </bt:LongStrings>
    </Resources>
    ```

### Step 2: Customize the HTML

Open the file **index.html** to specify the HTML for the add-in. Replace the generated `main` tag with the following markup, and save the file.

```
<div id="content-main">
    <div class="padding">
        <p>Choose the button below to set the color of the selected range to green.</p>
        <br />
        <h3>Try it out</h3>
        <button class="ms-Button" id="set-color">Set color</button>
    </div>
</div>
```

### Step 3: Customize the CSS

Open the file **app.css** to specify the custom styles for the add-in. Replace the entire contents with the following code and save the file.

```
#content-header {
    background: #2a8dd4;
    color: #fff;
    position: absolute;
    top: 0;
    left: 0;
    width: 100%;
    height: 80px; 
    overflow: hidden;
}

#content-main {
    background: #fff;
    position: fixed;
    top: 80px;
    left: 0;
    right: 0;
    bottom: 0;
    overflow: auto; 
}

.padding {
    padding: 15px;
}
```

### Step 4: Customize the script

Open the file **src\index.js** to specify the script for the add-in. Replace the entire contents with the following code and save the file.

...

## Sideload the manifest

1. In your command prompt/shell, make sure you are in the root directory of your project, and enter `npm start`. This will start a web server at `https://localhost:3000` and open your default browser to that address.

1. If your browser indicates that the site's certificate is not trusted, you will need to add the certificate as a trusted certificate. Outlook will not load add-ins if the site is not trusted. See [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for details.

1. After your browser loads the add-in page without any certificate errors, follow the instructions in [Sideload Outlook Add-ins for testing](sideload-outlook-add-ins-for-testing.md) to sideload the **my-office-add-in-manifest.xml** file.

## Try it out

1. Once you've sideloaded the manifest, open an appointment in a new window in Outlook.

1. On the **Appointment** tab , locate the add-in's **Display all properties** button.

    ![A screenshot of a message window in Outlook with the add-in button highlighted](images/quick-start-button.PNG)

1. Click the button to open the add-in's taskpane.

    ![A screenshot of the add-in's taskpane displaying message properties](images/quick-start-task-pane.PNG)

## Congratulations!

Congratulations, you've successfully created an Outlook add-in! To learn more about creating Outlook add-ins, checkout the Outlook add-ins developer documentation at [https://aka.ms/outlook-add-ins-docs](https://aka.ms/outlook-add-ins-docs).
