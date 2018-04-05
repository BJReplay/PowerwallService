# PowerwallService
This windows service can log data from a local Powerwall to a local database and / or an azure database, can control pre-charging during off-peak, and can upload data to PVOutput.

Important: This tool is designed for users who have access to the customise screen that allows them to set a backup percentage in self consumption mode.  If you do not have this option, the API calls that this service makes to request your powerwall to charge or to stop charging may not work.  You may end up with your powerwall stuck in a mode you do not want.  You may not be able to get it out of this mode.  You may be voiding your warranty.  You may end up with a very expensive, very heavy battery that is NOT saving you any money.  You run this program entirely at your own risk.
==================================================

Powerwall Service basic installation instructions.
--------------------------------------------------

Prerequisites - .Net Framework 4.6 or above
-------------

Unzip into a directory (e.g. C:\Tools\PowerwallService)

Open an administrator command prompt and navigate to that director

Run the following command to install the service
	
	installutil PowerwallService.exe

In the same directory, edit PowerwallService.exe.config in your favourite XML or text editor (an XML editor is good, because it will highlight if you corrupt the structure)

Review the comments on each parameter

Ignore the parameters that are NOT CURRENTLY USED

Make sure you find and replace (or delete, after reviewing) every TODO in the file

Note that there are TODO comments in comments, and TODO tags in values - you must replace or remove these

Don't attempt to set up a database for logging yet - let's just get the service running.

From the administrator command prompt you opened earlier, once you've edited the config file

	NET START PowerwallService

With any luck, the service will start successfully.

If so, hop into event viewer, windows logs, application log, and look for at least three entries from PowerwallService as it starts.

You'll then see a "running" event (Source=PowerwallService, EventID = 200) every minute.

Wait just over 10 minutes, and you should see EventID 1000 which will show you've got a solar forecast back from the API.

All entries (and errors) will be logged under source Powerwall Service.

Key log entries to look out for are Events 500 to 520 - this is where the service is logging it's thoughts about pre-charging, and 1000 - where it logs the next three day's forecast (including today) every hour.  You should see EventID 503 every 10 minutes where it reports your powerwall SOC and whether it is in the operation period (e.g. off peak) or not.

If you stay up until the start of off peak (or set the clock forward on your PC / server) you will see it report when it is in the operation period.

If you want to trigger an immediate forecast check after starting the service, PAUSE then RESUME the service - this triggers the 10 minute check immediately.

The two *Filter.xml files in the Zip allow you to set up custom views that show Key (charge / forecast) and All events.

If you want to set up SQL Server based logging (either azure or local) see below - otherwise skip this section.


SQL Server Setup Notes

Find TODO in the script and use the instructions to determine the official name of your time zone and replace AUS Eastern Standard Standard Time with your timezone

Check all TODOs in the script - you'll either need to change something, or make sure you run the script in sections.

The Azure script may need modification to be run - in all likelihood you will create the database through the azure portal, 
and then run the script in the target database from the point starting after USE [PWHistory]
