# Auto-DMARC Auditor

I work at an MSP, and because of this, I get a lot of emails with DMARC aggregate reports. And honestly, trying to parse all those is a pain. So I automated it.

This consists of two components: An Outlook VBA Class that automatically calls the script for each new email in the Inbox/DMARC folder with an attachment, and a Powershell Script that unzips and analyzes the attachments.

This requires [BurntToast v1.0.0](https://www.powershellgallery.com/packages/BurntToast/1.0.0-Preview1) which you can install directly with the `import-module burnttoast -allowprerelease` cmdlet in Powershell.

This requires [7Zip4Powershell 2.2.0](https://www.powershellgallery.com/packages/7Zip4Powershell/2.2.0) or greater, provided a new update doesn't break anything. And also, as a consequence, [7Zip](https://www.7-zip.org/), which you can also install install via [Chocolatey](https://community.chocolatey.org/packages/7zip)