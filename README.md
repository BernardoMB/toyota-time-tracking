# toyota-time-tracking
This application automates the timesheet reporting process for Toyota and Advent and it is part of a larger automation process. 

The first step, and the only manual step of the process, which is carried out externally, is to forward as an attachment the approval email from the supervisor using a forwarding rule in Outlook's server. That approval email is then stored in the folder 'Approvals' in GMail's server.

	
	- **Step 1**: Forward the approval email from the supervisor to the 'Approvals' folder in GMail's server. `Right Click on Approval -> Forward as Attachment -> Type in Receiver -> Send` (TODO: Automate this step using no-code PowerAutomate).

	- **Step 2**: Run this program

In this seconds step the application will read the approval email from the 'Approvals' folder. The approval email attachment will be downloaded into the file system.

The program will then read the approval email and send it as an attachment to the employer using the SMTP server configured in the application's mail service. The email will be sent to the employer's email address, which is also configured in the application settings.

The next step in the process is to generate the timesheet for the current week. This timehseet is then stored in the file system and sent to the employer and the supervisor as an attachment.

Sending the reporting email to the employer is straight forward, but sending the reporting email to the supervisor requires some additional steps. 

To send the reporting email to the supervisor, the application sends out the reporting email to the corporate address inbox and then a [https://make.powerautomate.com/](Prower Automate flow) will trigger and forward the email to the supervisor's email address. The Power Automate flow is configured to run every Friday at 5:00 PM and it will forward the email to the supervisor's email address, which is also configured in the application settings.

Note: This The application needs to be run every Friday at 5:00 PM. It will generate files and send the emails. 

# Intallation
- Clone the code repository.
- Install Packages
- Add the following user secrets
	```bash
	{
	  "Mailtrap:Username": "*****",
	  "Mailtrap:Password": "*****"
	  "GoogleAppPassword": "*****"
	}
	```
- Change existing configurations in `appsettings.json`

# Debug

Make sure you can connect to the SMTP sever by running the following command:
```bash
Test-NetConnection smtp.mailtrap.io -Port 587
// or
telnet smtp.mailtrap.io 587
```

Run the aplication from the comman line using:
```bash
 dotnet run -- <day> <hours> ...
 ```
 The arguments are optional.

# Integrations

## Power Automate
As a workaround to automate in corporate domain.

## Mailtrap
This application uses *Mailtrap* to send emails programatically.
DNS records are registered under domain `bernardomondragon.net`.

## Gmail
To access personal accounts.

## Office Open XML
This application uses OfficeOpenXML libraries for excel file manipulation.
