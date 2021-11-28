import { Client } from "@microsoft/microsoft-graph-client";
import { ResourceInterface } from "./interfaces/ResourceInterface";
import createOrRefreshSubscription from "./CreateOrRefreshSubscription";
import { syncAllEventsInCalendar } from "./Helpers";
import "isomorphic-fetch";

const calendarSync = async function (
  webhookData: any,
  client: Client
): Promise<string> {
  console.log("CALENDAR SYNC ACTIVITY RUNNING");

  //here we assign the changes recieved in the webook to a var with a more intuitive name
  let changes = webhookData.value;

  //here we get the user id to use in graph request paths
  const userId = await getUserID(client);

  for (const change of changes) {
    try {
      //here we get the affected calendar Id so that we can access it later
      let calendarId = (
        await client
          .api(`/users/${userId}/calendars/${change.resourceData.id}`)
          .get()
      ).id;
      //here we get the group Id that the affected calendar is a member of
      let calendarGroupId = await getCalendarGroupID(client, userId);
      //here i request the calendars in the group and map the value returned down to just an array of the ids to iterate through next
      let calendarIds = (
        await client
          .api(`/users/${userId}/calendarGroups/${calendarGroupId}/calendars`)
          .get()
      ).value.map((calendar: { id: any }) => calendar.id);
      //here we test the client state param in the webhook to ensure it came from MS.
      //we also check to see if the calendar we recieved the update for is in the relevant calendar group (i.e., a member of process.env.CALENDAR_GROUP_NAME)
      if (
        change.clientState == process.env.SUBSCRIPTION_SECRET &&
        calendarIds.includes(calendarId)
      ) {
        //if true, we will create a resource and then createOrRefreshSubscription for it
        const url = process.env.WEBHOOK_URL.includes("http")
          ? `${process.env.WEBHOOK_URL}/Events`
          : `https://${process.env.WEBHOOK_URL}/Events`;
        let currentResource: ResourceInterface = {
          name: `/users/${userId}/calendars/${calendarId}/events`,
          url,
          secret: process.env.SUBSCRIPTION_SECRET,
          changeType: "created",
        };

        await createOrRefreshSubscription(currentResource, client, 2.9);
        await syncAllEventsInCalendar(calendarId, client);
      } else {
        console.log(
          "calendar is a member of another group or was recently deleted, not creating a subscription"
        );
        console.log({
          why: {
            subSecretMatches:
              change.clientState == process.env.SUBSCRIPTION_SECRET,
            calendarIdsInRelevantGroup: calendarIds,
            calendarId,
            calendarWasRelevant: calendarIds.includes(calendarId),
          },
        });
      }
    } catch (e) {
      console.log(
        "calendar is a member of another group or was recently deleted, not creating a subscription"
      );
      console.log("attempting to delete subscription to this resource");
      console.log(e);
      try {
        await client.api(`/subscriptions/${change.subscriptionId}`).delete();
        console.log("deleted subscription");
      } catch (e) {
        console.log("the subscription must not exist, so we couldnt delete it");
        console.log(e);
      }
    }
  }
  //this doesnt need to necessarily return this token, it is for testing purposes
  return "function ran";
};
export const processCalendarChange = async (
  resourceData: any,
  client: Client
) => {
  console.log({ resourceData });
};
/**
 * gets the ID of the anchor user for the meeting distributor
 * @param client the ms graph client
 * @returns the id of the user who owns the synced calendar group
 */
export const getUserID = async (client: Client) => {
  return (await client.api(`/users/${process.env.CALENDAR_OWNER_UPN}`).get())
    .id;
};
/**
 * this function will get the calendar group ID for synced calendars.
 * It will either use the given user ID or pull it from graph if not provided
 */
export const getCalendarGroupID = async (client: Client, userId?: string) => {
  const idOfUser = userId ? userId : await getUserID(client);
  return (
    await client
      .api(
        `/users/${idOfUser}/calendarGroups?$filter=startswith(name,'${process.env.CALENDAR_GROUP_NAME}')`
      )
      .get()
  ).value[0].id;
};
export default calendarSync;
