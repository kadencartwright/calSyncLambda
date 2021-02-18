import { Client } from "@microsoft/microsoft-graph-client";
import {ResourceInterface} from './interfaces/ResourceInterface'
import createOrRefreshSubscription from './CreateOrRefreshSubscription'
import {syncAllEventsInCalendar} from './Helpers'
import 'isomorphic-fetch'

const calendarSync = async function (webhookData:any, client:Client): Promise<string> {
    console.log('CALENDAR SYNC ACTIVITY RUNNING')

    //here we assign the changes recieved in the webook to a var with a more intuitive name
    let changes = webhookData.value;

    //here we get the user id to use in graph request paths
    let userId = (await client.api(`/users/${process.env.CALENDAR_OWNER_UPN}`).get()).id

    for (const change of changes) {
        try{
            //here we get the affected calendar Id so that we can access it later
            let calendarId = (await client.api(`/users/${userId}/calendars/${change.resourceData.id}`).get()).id;
            //here we get the group Id that the affected calendar is a member of
            let calendarGroupId =  (await client.api(`/users/${userId}/calendarGroups?$filter=startswith(name,'${process.env.CALENDAR_GROUP_NAME}')`).get()).value[0].id;
            //here i request the calendars in the group and map the value returned down to just an array of the ids to iterate through next
            let calendarIds = (await client.api(`/users/${userId}/calendarGroups/${calendarGroupId}/calendars`).get()).value.map((calendar: { id: any; })=>calendar.id)
            //here we test the client state param in the webhook to ensure it came from MS.
            //we also check to see if the calendar we recieved the update for is in the relevant calendar group (i.e., a member of process.env.CALENDAR_GROUP_NAME)
            if (change.clientState = process.env.SUBSCRIPTION_SECRET && calendarIds.includes(calendarId) ){
                //if true, we will create a resource and then createOrRefreshSubscription for it
                let currentResource:ResourceInterface = {
                    name:`/users/${userId}/calendars/${calendarId}/events`,
                    url: `${process.env.WEBHOOK_URL}/Events`,
                    secret: process.env.SUBSCRIPTION_SECRET,
                    changeType:"created"
                }

                await createOrRefreshSubscription(currentResource,client,2.9)
                await syncAllEventsInCalendar(calendarId,client);
                
            }else{
                console.log('calendar is a member of another group or was recently deleted, not creating a subscription')
            }
        }catch(e){
            console.log('calendar is a member of another group or was recently deleted, not creating a subscription')
            console.log('attempting to delete subscription to this resource')
            console.log(e)
            try{
                let subs = await client.api(`/subscriptions/`).get()
                let sub = subs.filter((x:{resource:String})=> x.resource == change["@odata.id"])
                await client.api(`/subscriptions/${change.subscriptionId}`).delete();
                console.log('deleted subscription')
            }catch(e){
                console.log('the subscription must not exist, so we couldnt delete it')
                console.log(e)
            }

        }
    }
    //this doesnt need to necessarily return this token, it is for testing purposes
    return 'function ran';
    
};

export default calendarSync;