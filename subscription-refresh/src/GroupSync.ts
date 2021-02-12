import { Client } from "@microsoft/microsoft-graph-client";
import {getAllPagesFromGraph} from './Helpers'
import 'isomorphic-fetch'

const groupSync = async function (webhookData:any,client:Client): Promise<string> {
    console.log('GROUP SYNC ACTIVITY RUNNING')

    //assign the input from the orchestrator(Groups) function to a local variable
    let changes = webhookData.value;
    for (const change of changes) {
        if (change.clientState = process.env.SUBSCRIPTION_SECRET && change.resourceData['members@delta']){
            try{
                await groupUpdatedHandler(change,client)
            }catch(e){
                console.log(e)
            }
        }else{
            console.log('this group webhook did was not relevant so no processing was done')
            console.log('either subscription secret was incorrect or it did not contain a `members@delta` parameter')
        }
    }
    //this doesnt need to necessarily return this token, it is for testing purposes
    return 'success';
    
};


    let groupUpdatedHandler:(change:any,client:Client)=>void = async function(change:any,client:Client){


        //examine the section of webhook data that will tell us what users are different as well as whether they were added or deleted
        let userChanges = change.resourceData['members@delta']

        //iterate through the users that the webhook reported as affected. 
        //from here we will discern whether a user was added or removed from a group and respond appropriately
        for (const userChange of userChanges){

            //we need the user object to construct the attendee object,
            //the attendee object will be used to update each event in the calendar corresponding to the group they were added or removed from
            console.log('Getting User data')
            let user = await client.api(`/users/${userChange.id}`).get()
            let attendee = {emailAddress:{address:user.mail,name:user.displayName}}

            //wait for the group resource (selected only the displayName from graph, see 'odata query params' in MS Graph docs for details) and set the name to a var
            console.log('getting groups display name')
            let groupName = (await client.api(`/groups/${change.resourceData.id}?$select=displayName`).get()).displayName 


            //get the cal groups of the user that owns the calendars to sync to. set the name of the group as an app config setting or in local.settings.json for local development
            //call .value at the end of the api call because the useful part of the response to us is stored in .value
            // this api call is being filtered by startswith(). see odata query params in MS Graph docs for usage info
            console.log('getting calendar groups')
            let calendarGroups =  await getAllPagesFromGraph(client.api(`/users/${process.env.CALENDAR_OWNER_UPN}/calendarGroups?$filter=startswith(name,'${process.env.CALENDAR_GROUP_NAME}')`));
            //now that we have all the groups we want to sync, we can iterate through them and act on all their children (calendars)
            for (const group of  calendarGroups) {
                //the filter at the end of this api call ensures we only update events in the calendar corresponding to the group that has been changed
                console.log('getting calendars in group')
                let calendars = (await getAllPagesFromGraph(client.api(`/users/${process.env.CALENDAR_OWNER_UPN}/calendarGroups/${group.id}/calendars`)))
                .filter(calendar=> calendar.name.toLowerCase() == groupName.toLowerCase());
                //now that we have all the calendars contained in the current group we can iterate through them and act on all their children (individual events)
                for (const calendar of calendars){
                    //get all the individual events in a calendar
                    console.log('getting events in calendar')
                    let events = await getAllPagesFromGraph(client.api(`/users/${process.env.CALENDAR_OWNER_UPN}/calendarGroups/${group.id}/calendars/${calendar.id}/events`));
                    //this is the param in the webhook that tells us if the given user was added or removed.
                    if (userChange['@removed']=='deleted'){
                        for (const event of events){
                            console.log('updating event')
                            console.log('event:')
                            if( new Date(event.end.dateTime) > new Date()){
                                //only affect future events, not past ones
                                await client.api(`/users/${process.env.CALENDAR_OWNER_UPN}/calendarGroups/${group.id}/calendars/${calendar.id}/events/${event.id}?$select=id`)
                                .update({attendees:[...event.attendees.filter((e: { emailAddress: { address: any; }; })=>e.emailAddress.address!=attendee.emailAddress.address)]})
                            //if the user was removed, we will destructure a list of all the current attendees minus the affected user and send back to graph to update the attendee list in the event
                            }
                            
                        }
                    }else{
                        for (const event of events){
                            console.log('updating event')
                            console.log('event:')
                            console.log(new Date(event.end.dateTime))
                            if( new Date(event.end.dateTime) > new Date()){
                                //only affect future events, not past ones
                                await client.api(`/users/${process.env.CALENDAR_OWNER_UPN}/calendarGroups/${group.id}/calendars/${calendar.id}/events/${event.id}?$select=id`)
                                .update({attendees:[...event.attendees.filter((e: { emailAddress: { address: any; }; })=>e.emailAddress.address!=attendee.emailAddress.address),attendee]})//if attendee is already in an event we do not want to add them again, so I filtered this array to avoid duplications
                                //if the user was added, we will destructure a list of all the current attendees and append the affected user and send back to graph to update the attendee list in the event
                                //we also subtract the attendee from the original list to avoid attendee duplication
                            }
                            
                        }
                    }
                }
            }
        }
    }

export default groupSync;