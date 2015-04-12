# EggSink

Script for synchronizing data between Microsoft Exchange Server and Google Calendar.  Uses https://github.com/jamesiarmes/php-ews to talk to Exchange Web Services and https://github.com/google/google-api-php-client to talk to the Google Calendar. Requires at least PHP 5.4.

**_NOTE: Currently only does one-way synchronization of events from an Exchange calendar to Google Calendar._**

Simply download the source for this and setup the configuration.  You will need to have an active Exchange account and enabled Google Calendar API for your existing Google account.

To configure, create a config directory in the root of the project, and place a config.php file under the config directory. The config.php file should have the following settings defined:

```php
const SYNC_DAYS_FROM_NOW = 1; // number of days in the future to sync

const EXCHANGE_SERVER = ''; // the hostname of the Exchange server
const EXCHANGE_USERNAME = ''; // the username for the Exchange server
const EXCHANGE_PASSWORD = ''; // the password for the Exchange server

const GOOGLE_CALENDAR_ID = ''; // the ID of the Google Calendar being synced
const GOOGLE_CLIENT_ID = ''; // Google Calendar API service account client ID
const GOOGLE_EMAIL = ''; // Google Calendar API service account email
const GOOGLE_KEY_FILE = '*.p12'; // Google Calendar API service account p12 file name
```

Make sure the P12 file is also placed into the config directory.

To run the script (assuming php is in the path and already in the EggSink project directory), use the command line:
```
php eggsink.php
```

This script is probably most useful when configured to run on a schedule. 