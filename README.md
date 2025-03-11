# VBA_Trilium

**Outlook to Trilium Notes**

**Description**
This macro automates the creation of notes in Trilium when marking emails as "Important" in Outlook. The macro monitors the "Inbox" folder and creates a new note in Trilium with the email's subject and content when the "Important" flag is set.

**Setup/Configuration**
Copy the code: Paste the provided code into a new VBA module in Outlook to section ThisOutlookSession. 

![image](https://github.com/user-attachments/assets/af3ff690-2ade-4e05-82ea-4e3c57f3d329)

**Configure constants:**
TRILIUM_API_URL: Specify your Trilium API URL.

API_TOKEN: Enter your Trilium API token for authorization.

Enable macros: Ensure that macros are enabled in Outlook settings.
Restart Outlook


**How It Works**
Monitor the "Inbox" folder: The macro automatically tracks changes in your "Inbox".
Flag check: When the "Important" flag is set on an email, the macro checks this event.
Create a note: If the condition is met, a new note is created in Trilium with the email's subject and content.

**Notes**
This is the first version of the macro, so there may be errors or inconsistencies.
Your feedback and suggestions for improvement are welcome!


**Acknowledgments**

Trilium https://github.com/zadam/trilium 

TriliumNext/Notes https://github.com/TriliumNext/Notes 

Thanks to everyone who contributes to these projects!


🔑 License
This program is free software: you can redistribute it and/or modify it under the terms of the GNU Affero General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
