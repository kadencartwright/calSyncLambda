# CalSync

### About

This is a system to synchronize complex meeting groups with irregular recurring meetings.

it use AWS lambda and Microsoft's graph api to keep an office 365 group or distribution list in sync with a calendar of the same name inside a specified calendar group belonging to a particular user.

### How to deploy

you will need:

- The AWS CLI installed
- The AWS SAM CLI installed
  - Sign in with an admin account, or an IAM role with the following permissions
    - AWSLambdaFullAccess
    - IAMFullAccess
    - AmazonS3FullAccess
    - AmazonAPIGatewayAdministrator
    - AWSCloudFormationFullAccess
- Generate an app registration in your Azure AD tenant
  - generate client secret for your registration and make note of it's value. you will use this and the Application (client) ID as environment variables for the lambda Functions to use in order to call MSGraph
  - grant the app the proper permissions
    - from your app registration page go to 'api permissions'
    - click 'add a permission'
    - select 'Microsoft Graph'
    - select 'Application Permissions'
    - select the following permissions
      - Calendars.ReadWrite
      - Group.ReadWrite.All
      - GroupMember.Read.All
      - User.Read.All
      - OrgContact.read.all
    - sign in with an admin account for your Azure tenant and grant the permissions
- Fill in environment variables (descriptions of $VAR's to follow)

deploy (first time):

    $cd ./subscription-refresh && tsc && cd ../ && sam build -m subscription-refresh/package.json -s ./subscription-refresh && sam deploy --guided

deploy (after first deploy):

    $cd ./subscription-refresh && tsc && cd ../ && sam build -m subscription-refresh/package.json -s ./subscription-refresh && sam deploy --guided

## Azure AD env setup

### The Anchor user

CalSync requires a user account to use as an anchor. This user is where we will set up the calendar group to hold our synced calendars

### The Synced Calendar Group

Cal Sync works by watching for changes in calendars in a given calendar group belonging to the anchor user. It receives a web hook any time an event Is added to one of the calendars within the synced Calendar Group

### Synced Calendars

CalSync Watches for new events in the synced calendars (anything contained in the Synced Calendar Group). When an event is added, Calsync checks to see If there is a group in the azure tenant whose email address matches the naming scheme for the calendar. There is no database of calendars and their corresponding groups, so the link only works when the naming is correct on both calendar and group side.

The relationship between groups and calendars is explained below in the section " Calendar and Group relationship"

### Synced Groups

CalSync also watches for new and removed members of a group. CalSync will retrieve all events in the calendar corresponding to the group that was updated and adjust their attendeee list accordingly.

- when a user is added, they are added to all events on the corresponding calendar
- when a user is removed, they are removed from all events on the corresponding calendar

any of these updates only affect meetings who's scheduled end time is in the future, not past meetings

The relationship between groups and calendars is explained below in the section " Calendar and Group relationship"

## Environment Variables

In order to set up the environment for CalSync to work properly, you must configure a few environment variables. These can be added most easily by editing the template.yaml.copy file in the project root. Simply add new entries under

- Globals:
  - Function:
    - Environment:
      - Variables:

You will need the variables listed below:

- CLIENT_ID: The client ID for your App Registration in Azure AD

- CLIENT_SECRET: The client secret for your App Registration in Azure AD

- WEBHOOK_URL: The endpoint where your api gateway listens to requests. something like: `https://yourApiAddress.execute-api.us-east-2.amazonaws.com/Prod/Webhooks`

- SUBSCRIPTION_SECRET: Any random string. this will be used to verify that web hooks are actually from Microsoft. this string is sent when a web hook subscription is made and any authentic requests will contain this string as a parameter

- TENANT_ID: Your Azure AD tenant ID

CALENDAR_OWNER_UPN: The Anchor user for the calendar group. looks like `yourcalendaranchoruser@yourdomain.com`

CALENDAR_GROUP_NAME: The name of the calendar group which will hold your calendar. This is so that calendar anchor account can belong to an actual user and not interfere with their personal calendar.

GROUP_EMAIL_PREPEND: The string that should be prepended to the email address of any sync group.

## Calendar and Group relationship

As stated above, the relationship between a calendar and a group is held intact by their names.

For example:

##### Given that:

- `GROUP_EMAIL_PREPEND` is set to `calsyncgroup`
- `CALENDAR_OWNER_UPN` is set to `user@contoso.com`
- `CALENDAR_GROUP_NAME` is set to `Company Meetings`

Then a hypothetical Group/Calendar relationship would look as follows.

- Group:
  - Group name : `C Level Officers meeting group`
  - Group email: `calsyncgroupclevelofficers@contoso.com`
- Calendar:
  - Calendar name: `C Level Officers`

The Group email must be a transformed version of the Calendar name as it appears in the Anchor Users calendar group. This transformation is as follows:

- all lowercase
- no spaces

this is so that the format of the group name can be transformed to a string that can be part of an email address.
