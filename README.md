This VBA (Visual Basic for Applications) script is designed to automatically generate tasks in Microsoft Outlook based on certain criteria found in the subject or body of an email. Specifically, it identifies emails containing the term "AR" (which likely stands for Action Required) and extracts relevant information to create tasks.

How It Works
When an email is about to be sent (ItemSend event), the script checks if the item is a mail item. If so, it proceeds to analyze the email's subject and body.

Key Features
Identification of AR Items: The script searches for occurrences of "AR" in both the subject and body of the email to identify Action Required items.
Task Creation: For each identified AR item, the script creates a task in Outlook. The task's subject includes details extracted from the email, and its due date is calculated based on the work week and day specified in the email body.
Email Modification: After processing, the script adds a notification to the email body indicating that automatic tasks were generated for the identified AR items.
Usage Instructions
Integration with Outlook: This script should be added to the Outlook VBA editor to take effect. It can be added as a macro in Outlook.
Trigger: The script is triggered when an email is about to be sent.
Email Formatting: To ensure proper detection and extraction of information, emails containing AR items should follow a specific format. The email's body should include a clear delineation indicating the start of new text (commonly marked by "From:"). Additionally, AR items should be clearly labeled within the body text.
Customization
Thread Marker: The script assumes the presence of a thread marker ("From:") to identify the start of new text in the email body. This can be adjusted based on the email client's specific thread marker.
Date Calculation: The script calculates task due dates based on the work week and day specified in the email body. Adjustments may be needed depending on the organization's work week structure.
Function Explanation
CalculateDateFromWWAndDay: This function calculates the due date for a task based on the specified work week and day of the week. It takes into account the start of the year and adds the appropriate number of days to determine the due date.
Disclaimer
This script is provided as-is and may require customization to suit specific organizational needs or email formats. Users should ensure proper testing and validation before deploying it in a production environment.

Note: This README provides an overview of the script's functionality and usage instructions. For detailed implementation steps, refer to the Outlook VBA documentation or consult with a qualified IT professional.







