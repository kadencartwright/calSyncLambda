import { Client } from "@microsoft/microsoft-graph-client";
import {MyAuthenticationProvider} from './MyAuthenticationProvider'
import 'isomorphic-fetch'

const eventSync = async function (webhookData:any, client:Client): Promise<string> {
    console.log('EVENT SYNC ACTIVITY RUNNING')
    console.log('getting a token from auth provider')
    let token = await new MyAuthenticationProvider().getAccessToken();

    //assign the input from the orchestrator(Events/index.ts) function to a local variable
    let changes = webhookData.value;

    for (const change of changes) {
        if (change.clientState = process.env.SUBSCRIPTION_SECRET){

            switch(change.changeType){
                case "created":
                    await createdHandler(change,client)
                    break;
                case "updated":
                    //do nothing
                    break;
                case "deleted":
                    //do nothing
                    break;
            }
        }
    };

    //this doesnt need to necessarily return this token, it is for testing purposes
    return token;
    
};


    let createdHandler:(change:any,client:Client)=>void = async function(change:any,client:Client){
        //get the id of the calendar the given event is a member of
        console.log('getting event info from graph')
        let event = (await client.api(change.resource).expand('calendar').get())
        //extract the name
        console.log('getting calendar info from graph')
        let calendarName = (await client.api(`/users/${process.env.CALENDAR_OWNER_UPN}/calendars/${event.calendar.id}`).get()).name
        //find the group
        console.log('getting corresponding group info')
        let group = (await client.api(`/groups`).filter(`startswith(mail,'${calendarName}@' ) `).get()).value[0]
        
        console.log('getting group members')
       try{     
            let membersResponse = (await client.api(`/groups/${group.id}/members`).select('mail,displayName').get())
            let attendees = event.attendees;
            for (const member of membersResponse.value){
                let found = false
                for (const attendee of attendees){
                    if (member.mail == attendee.emailAddress.address){
                        found = true
                        break
                    }
                }
                if (!found){
                    attendees.push({emailAddress:{address:member.mail,name:member.displayName}})
                }
            }
            console.log('updating attendees')
            await client.api(`/users/${process.env.CALENDAR_OWNER_UPN}/events/${event.id}?$select=id`)
                            .update({attendees:[...event.attendees,...attendees]})
        }catch(e){
                console.log(e)
            }
        
    }



export default eventSync;


