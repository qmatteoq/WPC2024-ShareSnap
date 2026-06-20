# ShareSnap

**ShareSnap** is the end-to-end demo presented at **WPC 2024**. It turns a document into
a ready-to-share social media post by orchestrating a team of AI agents, and surfaces the
experience directly inside Word through an Office Add-in.

The solution combines **Semantic Kernel** multi-agent orchestration, **Microsoft Graph**
and an **Office Add-in** front end.

## Architecture

| Project | Stack | Responsibility |
| --- | --- | --- |
| `ShareSnapAPI` | ASP.NET Core (.NET), Semantic Kernel | Back-end API that runs a multi-agent group chat to summarize content, draft a social post, review it and share it. |
| `ShareSnapAddin` | TypeScript / React (Office Add-in task pane) | Word task-pane add-in that sends the active document to the API and shows the generated post. |

### The agents

The API coordinates a group of Semantic Kernel agents (each defined as a YAML prompt under
`ShareSnapAPI/Prompts`):

- **SummarizerAgent** — summarizes the source document.
- **SocialNetworkExpertAgent** — writes an engaging social media post from the summary.
- **SocialNetworkReviewerAgent** — reviews and refines the post.
- **MailShareAgent** — shares the final post by e-mail to a recipient via the **Microsoft Graph** plugin (`MsGraphPlugin`).

Scenarios (e.g. `BlogAuthorScenario`) wire these agents together into a coordinated workflow.
Both **OpenAI** and **Azure OpenAI** back ends are supported (`KernelCreator`). A
[Dev Proxy](https://aka.ms/devproxy) configuration (`devproxyrc.json`) is included for local
testing.

## Prerequisites

- [.NET SDK](https://dotnet.microsoft.com/download)
- [Node.js](https://nodejs.org/) (LTS) for the add-in
- An **OpenAI** or **Azure OpenAI** endpoint and key
- A Microsoft 365 tenant + an Entra ID app registration with Microsoft Graph **Mail.Send** permission (for the mail-sharing agent)

## Getting started

### API

1. Configure your model and Graph settings in `ShareSnapAPI/ShareSnapAPI/appsettings.json`
   (use user-secrets for keys — do not commit them).
2. Run the API:
   ```bash
   cd ShareSnapAPI/ShareSnapAPI
   dotnet run
   ```

### Office Add-in

1. Install dependencies and start the dev server:
   ```bash
   cd ShareSnapAddin
   npm install
   npm start
   ```
2. Sideload the add-in into Word and open the **ShareSnap** task pane.

## License

This sample is provided **as-is**, for demonstration purposes. No explicit license has
been specified.
