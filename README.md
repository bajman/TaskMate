Here is a GitHub repository description for your project **TaskMate**:

---

## TaskMate

TaskMate is an AI-powered assistant that integrates with Microsoft 365 to help users manage their calendars seamlessly. It allows users to create, update, delete, and search for calendar events, all through intuitive and interactive prompts. TaskMate leverages OpenAI’s GPT models alongside Microsoft Graph API to provide powerful calendar management functionalities directly from the terminal.

### ⚠️ Beta Disclaimer
TaskMate is currently in **beta**. While many features are functional, there may be bugs or unexpected behaviors. New features and enhancements are planned, including support for additional Microsoft 365 services like To-Do, Email, and SharePoint.

### Features
- **Create Calendar Events**: Easily set up single or recurring events with details like time, date, title, description, location, and reminders.
- **Update Events**: Modify existing events with new details, such as adjusting the time, location, or attendees.
- **Delete Events**: Remove events from your calendar with confirmation to avoid accidental deletions.
- **List and Search Events**: View your events based on specific filters, such as date range, categories, or subject.
- **Manage Categories**: Create, update, or delete event categories to organize your schedule effectively.
- **Interactive Prompts**: Use intuitive, multiline command prompts with suggestions and completions for a streamlined user experience.

### Future Plans
- **Microsoft To-Do Support**: Manage tasks and lists with the same ease as calendar events.
- **Email Integration**: Send, receive, and manage emails from your Microsoft 365 account.
- **SharePoint Integration**: Manage files and collaborate within SharePoint sites.
- **More Microsoft Graph API Endpoints**: Additional support for Microsoft 365 services will be added over time.

### Installation and Usage
TaskMate requires the following environment variables to function:
- `OPENAI_API_KEY`: Your OpenAI API key to enable the assistant’s intelligence.
- `CLIENT_ID`, `TENANT_ID`: Your Microsoft 365 credentials to access the Graph API.

To install and run TaskMate, clone the repository and run:
```bash
git clone <repo-url>
cd taskmate
pip install -r requirements.txt
python taskmate.py
```

### Contribution
This project is under active development, and contributions are welcome! If you encounter any issues or have ideas for improvements, please open a pull request or issue.
