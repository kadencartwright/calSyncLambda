import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { MyAuthenticationProvider } from "./MyAuthenticationProvider";

import createOrRefreshSubscription from "./CreateOrRefreshSubscription";
import { syncAllEventsInCalendar } from "./Helpers";
// const axios = require('axios')
// const url = 'http://checkip.amazonaws.com/';

export const lambdaHandler = async function (event: any): Promise<String> {
  let subscriptionLength = 2.9;

  //see if subscription exists

  const options = {
    authProvider: new MyAuthenticationProvider(),
  };
  const client = Client.initWithMiddleware(options);

  let groupResource = {
    name: "/groups",
    url: `${process.env.WEBHOOK_URL}/Groups`,
    secret: process.env.SUBSCRIPTION_SECRET,
    changeType: "updated,deleted",
  };

  //event resources need to be an array.
  //array should contain all synced calendars, as well as the master group to listen for new calendars to subscribe to

  let eventResource = [];
  console.log("getting user id");

  let userId = (
    await client.api(`/users/${process.env.CALENDAR_OWNER_UPN}`).get()
  ).id;

  console.log("getting calendar group");
  let calendarGroup = (
    await client
      .api(
        `/users/${userId}/calendargroups?$filter=name eq '${process.env.CALENDAR_GROUP_NAME}'`
      )
      .get()
  ).value[0];
  console.log("getting calendars in group");
  let calendars = (
    await client
      .api(`/users/${userId}/calendargroups/${calendarGroup.id}/calendars`)
      .get()
  ).value;
  for (const calendar of calendars) {
    console.log("getting calendar id");
    let calendarId = (
      await client
        .api(
          `/users/${userId}/calendargroups/${calendarGroup.id}/calendars/${calendar.id}`
        )
        .get()
    ).id;
    eventResource.push({
      name: `/users/${userId}/calendars/${calendarId}/events`,
      url: `${process.env.WEBHOOK_URL}/Events`,
      secret: process.env.SUBSCRIPTION_SECRET,
      changeType: "created",
      calendarId: calendarId,
    });
  }
  let calendarGroupResource = {
    changeType: "updated,deleted",
    url: `${process.env.WEBHOOK_URL}/Calendars`,
    secret: process.env.SUBSCRIPTION_SECRET,
    name: `/users/${userId}/calendargroups/${calendarGroup.id}/calendars`,
  };

  let resources = [groupResource, ...eventResource, calendarGroupResource];

  for (const resource of resources) {
    await createOrRefreshSubscription(resource, client, subscriptionLength);
  }

  return "success";
};
