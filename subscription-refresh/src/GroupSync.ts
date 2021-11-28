import { Client } from "@microsoft/microsoft-graph-client";
import { getAllPagesFromGraph } from "./Helpers";
import "isomorphic-fetch";

const groupSync = async function (
  webhookData: any,
  client: Client
): Promise<string> {
  console.log("GROUP SYNC ACTIVITY RUNNING");
  //assign the input from the orchestrator(Groups) function to a local variable
  await handleGroupChange(webhookData, client);
  //this doesnt need to necessarily return this token, it is for testing purposes
  return "success";
};
export async function processGroupChange(resourceData: any, client: Client) {
  if (resourceData["members@delta"]) {
    try {
      await groupUpdatedHandler(resourceData, client);
    } catch (e) {
      console.log(e);
    }
  } else {
    console.log(
      "change data did not contain a `members@delta` parameter so no processing was done"
    );
  }
}
export async function handleGroupChange(
  changeData: { value: any[] },
  client: Client
) {
  let changes = changeData.value;
  for (const change of changes) {
    await processGroupChange(change.resourceData, client);
  }
}
let groupUpdatedHandler: (resourceData: any, client: Client) => void = async (
  resourceData: any,
  client: Client
) => {
  //examine the section of webhook data that will tell us what users are different as well as whether they were added or deleted
  let userChanges = resourceData["members@delta"];

  //iterate through the users that the webhook reported as affected.
  //from here we will discern whether a user was added or removed from a group and respond appropriately
  console.log(`received ${userChanges.length} user changes`);
  for (const userChange of userChanges) {
    try {
      //we need the user object to construct the attendee object,
      //the attendee object will be used to update each event in the calendar corresponding to the group they were added or removed from
      console.log("Getting User data");
      const isOrgContact = userChange["@odata.type"]?.includes("orgContact");

      let user = isOrgContact
        ? await client.api(`/contacts/${userChange.id}`).get()
        : await client.api(`/users/${userChange.id}`).get();
      let attendee = {
        emailAddress: { address: user.mail, name: user.displayName },
      };
      console.log({ user, isOrgContact });
      //wait for the group resource (selected only the displayName from graph, see 'odata query params' in MS Graph docs for details) and set the name to a var
      console.log("getting groups display name");
      //let groupName = (await client.api(`/groups/${change.resourceData.id}?$select=displayName`).get()).displayName
      let groupMail: String = (
        await client.api(`/groups/${resourceData.id}?$select=mail`).get()
      ).mail;

      //get the cal groups of the user that owns the calendars to sync to. set the name of the group as an app config setting or in local.settings.json for local development
      //call .value at the end of the api call because the useful part of the response to us is stored in .value
      // this api call is being filtered by startswith(). see odata query params in MS Graph docs for usage info
      console.log("getting calendar groups");
      let calendarGroups = await getAllPagesFromGraph(
        client.api(
          `/users/${process.env.CALENDAR_OWNER_UPN}/calendarGroups?$filter=startswith(name,'${process.env.CALENDAR_GROUP_NAME}')`
        )
      );
      console.log(`found ${calendarGroups.length} calendar groups`);
      //now that we have all the groups we want to sync, we can iterate through them and act on all their children (calendars)
      for (const group of calendarGroups) {
        //the filter at the end of this api call ensures we only update events in the calendar corresponding to the group that has been changed
        console.log("getting calendars in group");
        let calendars = (
          await getAllPagesFromGraph(
            client.api(
              `/users/${process.env.CALENDAR_OWNER_UPN}/calendarGroups/${group.id}/calendars`
            )
          )
        ).filter((calendar) => {
          return groupMail
            ?.toLowerCase()
            .startsWith(
              `${process.env.GROUP_EMAIL_PREPEND}${calendar.name
                ?.toLowerCase()
                .replace(/\s/g, "")}@`
            );
        });
        console.log(`found ${calendars.length} calendars`);
        //now that we have all the calendars contained in the current group we can iterate through them and act on all their children (individual events)
        for (const calendar of calendars) {
          //get all the individual events in a calendar
          console.log("getting events in calendar");
          let events = await getAllPagesFromGraph(
            client.api(
              `/users/${process.env.CALENDAR_OWNER_UPN}/calendarGroups/${group.id}/calendars/${calendar.id}/events`
            )
          );
          console.log({ userChange });
          // our filter helper function
          const eventIsFuture = (event: any) =>
            new Date(event.end.dateTime).getTime() > new Date().getTime();
          //this is the param in the webhook that tells us if the given user was added or removed.
          if (userChange["@removed"]) {
            for (const event of events.filter(eventIsFuture)) {
              console.log("updating event");
              console.log("event:");
              console.log({ eventEnd: event.end.dateTime });
              console.log(
                `removing attendee '${attendee.emailAddress.name}' with email address '${attendee.emailAddress.address}'`
              );
              //only affect future events, not past ones
              await client
                .api(
                  `/users/${process.env.CALENDAR_OWNER_UPN}/calendarGroups/${group.id}/calendars/${calendar.id}/events/${event.id}?$select=id`
                )
                .update({
                  attendees: [
                    ...event.attendees.filter(
                      (e: { emailAddress: { address: any } }) =>
                        e.emailAddress.address != attendee.emailAddress.address
                    ),
                  ],
                });
              console.log(
                `removed attendee with email '${attendee.emailAddress.address}' from event '${event}'`
              );
              //if the user was removed, we will destructure a list of all the current attendees minus the affected user and send back to graph to update the attendee list in the event
            }
          } else {
            for (const event of events.filter(eventIsFuture)) {
              console.log("updating event");
              console.log("event:");
              console.log(new Date(event.end.dateTime));
              //only affect future events, not past ones
              console.log(
                `adding attendee '${attendee.emailAddress.name}' with email address '${attendee.emailAddress.address}'`
              );
              await client
                .api(
                  `/users/${process.env.CALENDAR_OWNER_UPN}/calendarGroups/${group.id}/calendars/${calendar.id}/events/${event.id}?$select=id`
                )
                .update({
                  attendees: [
                    ...event.attendees.filter(
                      (e: { emailAddress: { address: any } }) =>
                        e.emailAddress.address != attendee.emailAddress.address
                    ),
                    attendee,
                  ],
                }); //if attendee is already in an event we do not want to add them again, so I filtered this array to avoid duplications
              console.log(
                `added attendee with email '${attendee.emailAddress.address}''`
              );
              //if the user was added, we will destructure a list of all the current attendees and append the affected user and send back to graph to update the attendee list in the event
              //we also subtract the attendee from the original list to avoid attendee duplication
            }
          }
        }
      }
    } catch (e) {
      console.log("error while getting users data", { error: e });
    }
  }
};

export default groupSync;
