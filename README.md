# outlook-google-calendar-sync
Syncs an outlook calendar with your google calendar locally, without a service. This is useful when you're unable to authorize third party applications access to your office account due to organizational requirements. 

# Setup

1. Go to the google cloud console: https://console.cloud.google.com/welcome
2. Create a new project
3. Go to "Enable APIs and Services" and enable "Google Calendar API": https://console.cloud.google.com/apis/library/calendar-json.googleapis.com
4. Go to the Credentials tab: https://console.cloud.google.com/apis/credentials
5. Click "Create Credentials" -> "Service Account"
6. Configure as desired.
7. When created, open the service account and copy the email address for it: `{name}@{project}.iam.gserviceaccount.com`
8. Go to the Keys tab and click to create a new key, save this as google.json (or whatever you configure) where the application is set to run.
9. Go to your google calendar: https://calendar.google.com/calendar
10. Go to settings for your calendars, create a new calendar.
11. On the settings for the new calendar, go down to "Share with specific people or groups"
12. Click to "Add people and groups" and enter the email copied from step 7. Select "Manage changes to events".
13. Scroll down to the "Integrate calendar" section and copy the calendar ID: `{unique-id}@group.calendar.google.com`
14. Place this in your app settings for the app under `CalendarId`

# Build/Run

1. Make sure you have MSBuild in your path and open a terminal at the project root
2. Run `msbuild /t:publish /p:Configuration=Release`
3. It will create a build in the publish folder that you can put the following into:

## appsettings.json
```json
{
  "GoogleCalendarApi": {
    "CalendarId": "{secret-id-here}", // This comes from step 13 above
    "CredentialPath": "google.json",
    "Timezone": "America/New_York"
  }
}
```

## google.json
```json
{
  "type": "service_account",
  "project_id": "",
  "private_key_id": "",
  "private_key": "-----BEGIN PRIVATE KEY-----STUFFGOESHERE-----END PRIVATE KEY-----",
  "client_email": "{service-account-id}@{project-id}.iam.gserviceaccount.com",
  "client_id": "",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/something",
  "universe_domain": "googleapis.com"
}
```

4. Once the above files have been setup, you can run the program:

