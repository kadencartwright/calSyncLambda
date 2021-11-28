import { Client, GraphRequest } from "@microsoft/microsoft-graph-client";
import { MyAuthenticationProvider } from "./MyAuthenticationProvider";
async function syncAllEventsInCalendar(calendarId: String, client: Client) {
  console.log("getting calendar info from graph");
  let calendarName = (
    await client
      .api(`/users/${process.env.CALENDAR_OWNER_UPN}/calendars/${calendarId}`)
      .get()
  ).name;
  //find the group
  try {
    console.log("getting corresponding group info");
    let group = (
      await client
        .api(`/groups`)
        .filter(
          `startswith(mail,'${process.env.GROUP_EMAIL_PREPEND}${calendarName
            .replace(/\s/g, "")
            .toLowerCase()}@' ) `
        )
        .get()
    ).value[0];
    if (group) {
      console.log("group found:", { group });
      let events = await getAllPagesFromGraph(
        client.api(
          `/users/${process.env.CALENDAR_OWNER_UPN}/calendars/${calendarId}/events`
        )
      );
      console.log("getting group members");
      let groupMembers = await getAllPagesFromGraph(
        client.api(`/groups/${group.id}/members`).select("mail,displayName,id")
      );
      for (const event of events) {
        let attendees = [];
        for (let groupMember of groupMembers) {
          if (groupMember["@odata.type"].includes("graph.orgContact")) {
            //this means member is contact not AD user. need to fetch contact info.\                console.log('member:')
            console.log(JSON.stringify(groupMember));
            groupMember = await client
              .api(`/contacts/${groupMember.id}`)
              .select("mail,displayName,id")
              .get();
          }
          let found = false;
          for (const attendee of event.attendees) {
            if (groupMember.mail == attendee.emailAddress.address) {
              found = true;
              break;
            }
          }
          if (!found) {
            attendees.push({
              emailAddress: {
                address: groupMember.mail,
                name: groupMember.displayName,
              },
            });
          }
        }
        console.log("updating attendees");
        await client
          .api(
            `/users/${process.env.CALENDAR_OWNER_UPN}/events/${event.id}?$select=id`
          )
          .update({ attendees: [...event.attendees, ...attendees] });
      }
    } else {
      console.log("no group found");
    }
  } catch (e) {
    console.log(e);
  }
}
async function getAllPagesFromGraph(graphRequest: GraphRequest) {
  const { allPages } = await getAllPagesAndDeltaTokenFromGraph(graphRequest);
  return allPages;
}
async function getAllPagesAndDeltaTokenFromGraph(graphRequest: GraphRequest) {
  let morePages = false;
  let allPages = [];
  let result = await graphRequest.get();
  let deltaLink =
    "@odata.deltaLink" in result ? result["@odata.deltaLink"] : null;
  allPages.push(...result.value);
  morePages = "@odata.nextLink" in result ? true : false;
  while (morePages) {
    await wait(400); //wait .4 seconds to avoid potential rate limiting
    //create an auth provider to pass to the GraphClient (see microsoft graph javascript SDK for information)
    const options = {
      authProvider: new MyAuthenticationProvider(),
    };
    //init the graphClient with our auth provider
    const client = await Client.initWithMiddleware(options);
    result = await client.api(result["@odata.nextLink"]).get();
    allPages.push(...result.value);
    morePages = "@odata.nextLink" in result ? true : false;
    /** if the result contains a nextLink, assign it to this variable, else, retain the last value */
    deltaLink =
      "@odata.deltaLink" in result ? result["@odata.deltaLink"] : deltaLink;
    console.log({ delta: result?.["@odata.deltaLink"] });
  }
  return { allPages, deltaLink };
}
const wait = async (ms: number) => {
  return new Promise<void>((resolve, reject) => {
    setTimeout(() => resolve(), ms);
  });
};
export {
  syncAllEventsInCalendar,
  getAllPagesFromGraph,
  getAllPagesAndDeltaTokenFromGraph,
};
