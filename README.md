# Office Addins with Fluent UI React version 8

> **Important Note**: This repo is archived because we do not have the staff to update it. The samples in this repo may have security alerts. These vulnerabilities do not pose a problem for running and modifying this sample on your development computer, but do not use the vulnerable libraries in a production add-in.

This repository contains sample Office add-ins for developers who need to support Office versions that run add-ins in the Trident (IE) webview. If you don't need to support Trident in your add-in, we recommend that you create React-based Office add-ins with the quickstart found at [Use Fluent UI React in Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/quickstarts/fluent-react-quickstart). But projects created with the quickstart use Fluent UI React version 9 which doesn't support Trident. The samples in this repo use Fluent UI React version **8**. For information about which Office versions use Trident, see [Browsers and webview controls used by Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins).

The following table shows the main differences between the projects created with the quickstart and the projects in this repo.

| Characteristic         | This repository     | The quickstart |
|--------------|-----------|------------|
| React version | 17      | 18        |
| Fluent UI React verson      | 8  | 9       |
| Type of React components | class | function |
| Handling of React state changes | lifecycle events | hooks |
| Styling | CSS files | Fluent UI version 9 `makeStyles` API |

There are twelve samples in this repo, including a TypeScript and a JavaScript sample for Excel, OneNote, Outlook, PowerPoint, Project, and Word.

## Try it out

To run any of the samples:

1. Clone or download this repo.
1. In a command prompt, navigate to the sample you want to work with; for example, the sample in the /JavaScript/Excel folder. 
1. Run `npm install`.
1. If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.

    ```command&nbsp;line
    npm run dev-server
    ``` 

1. Run the following command in the root directory of the sample. This starts the local web server, it it isn't already running, and opens the Office host application with your add-in loaded.

    ```command&nbsp;line
    npm start
    ```
    > **Note*:* If this is the first time that you have sideloaded an Office add-in on your computer (or the first time in over a month), you're prompted first to delete an old certificate and then to install a new one. Agree to both prompts.

1. The Office application opens.

1. If the "My Office Add-in" task pane isn't already open, choose the **Home** tab (or **Message** tab in Outlook), and then choose the **Show Taskpane** button on the ribbon to open the add-in task pane.

1. A **WebView Stop On Load** prompt appears. Select **OK**.

## Questions and comments

Questions about Office Add-ins development in general should be posted to [Microsoft Q&A](https://docs.microsoft.com/answers/questions/185087/questions-about-office-add-ins.html). If your question is about the Office JavaScript APIs, make sure it's tagged with [office-js-dev]. For support for this repo, see the file SUPPORT.md in the root of the repo.

## Join the Microsoft 365 Developer Program
Get a free sandbox, tools, and other resources you need to build solutions for the Microsoft 365 platform.
- [Free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Get a free, renewable 90-day Microsoft 365 E5 developer subscription.
- [Sample data packs](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Automatically configure your sandbox by installing user data and content to help you build your solutions.
- [Access to experts](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Access community events to learn from Microsoft 365 experts.
- [Personalized recommendations](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Find developer resources quickly from your personalized dashboard.

## Join the Microsoft 365 Developer Program
Get a free sandbox, tools, and other resources you need to build solutions for the Microsoft 365 platform.
- [Free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Get a free, renewable 90-day Microsoft 365 E5 developer subscription.
- [Sample data packs](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Automatically configure your sandbox by installing user data and content to help you build your solutions.
- [Access to experts](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Access community events to learn from Microsoft 365 experts.
- [Personalized recommendations](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Find developer resources quickly from your personalized dashboard.

## Additional resources

* [Office Add-ins documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* More Office Add-ins samples at [OfficeDev on Github](https://github.com/officedev)

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Copyright

Copyright (c) 2023 Microsoft Corporation. All rights reserved.
