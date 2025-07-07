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
- Change existing configurations in `appsettings.json`

# Debug

Run the aplication from the comman line using:
```bash
 dotnet run -- <day> <hours> ...
 ```
 The arguments are optional.

# Integrations

## Mailtrap
This application uses *Mailtrap* to send emails programatically.
DNS records are registered under domain `bernardomondragon.net`.

## Office Open XML
This application uses OfficeOpenXML libraries for excel file manipulation.
