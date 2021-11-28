import EventSync from "./EventSync";
import GroupSync from "./GroupSync";
import CalendarSync from "./CalendarSync";
import { MyAuthenticationProvider } from "./MyAuthenticationProvider";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";

export const lambdaHandler = async (event: any): Promise<boolean> => {
  const options = {
    authProvider: new MyAuthenticationProvider(),
  };
  const client: Client = await Client.initWithMiddleware(options);
  console.log(JSON.stringify(event));
  if (!!event["type"]) {
    switch (event["type"].toLowerCase()) {
      case "events":
        console.log("events");
        console.log(await EventSync(event["data"], client));
        break;
      case "calendars":
        console.log("calendars");
        console.log(await CalendarSync(event["data"], client));
        break;
      default:
        break;
    }
  }
  return true;
};
