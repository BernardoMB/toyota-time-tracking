# toyota-time-tracking
This application automates the timesheet reporting process for Toyota and Advent.

# Intallation
- Clone the code repository.
- Install Packages
- Add the following user secrets
	```bash
	{
	  "Mailtrap:Username": "*****",
	  "Mailtrap:Password": "*****"
	}
	```

# Integrations

## Mailtrap
This application uses *Mailtrap* to send emails programatically.
DNS records are registered under domain `bernardomondragon.net`.

## Office Open XML
This application uses OfficeOpenXML libraries for excel file manipulation.
