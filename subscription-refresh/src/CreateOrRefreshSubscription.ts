import { ResponseType,Client } from '@microsoft/microsoft-graph-client';
import {getAllPagesFromGraph, syncAllEventsInCalendar} from './Helpers'
export default async function createOrRefreshSubscription(resource:any,client:Client,subscriptionLength:number){
    console.log('CREATING OR REFRESHING SUBCRIPTION')   
    console.log('getting subscriptions to check against')

    //if the sub does not exist, create it now
    let subsFromServer = (await getAllPagesFromGraph(client.api('/subscriptions')));

    let matchingSubscriptionFromServer = matchSubscription(resource,subsFromServer);

    //construct the new expiration date for the subscription. this is relevant/necessary for both new and updated subscriptions
    let newExpiration:Date = new Date()
    newExpiration.setDate(newExpiration.getDate()+subscriptionLength)


    if (!!matchingSubscriptionFromServer){//!!(non-null value) returns true, so this will run if we found a matching subscription from the server
        console.log(`subscription already exists`)
        let expiration = new Date(matchingSubscriptionFromServer.expirationDateTime)
        //dont refresh brand new subs
        let staleDate = new Date(expiration.setDate(expiration.getDate()-1))
        let staleSubscription:boolean = (new Date()>staleDate) 
        if (staleSubscription){    
            const updates = {
            expirationDateTime:newExpiration.toISOString()
            };
            console.log('updating subscription')
            await client.api(`/subscriptions/${matchingSubscriptionFromServer.id}`).update(updates);
        }
    }else{// if matchingSubscriptionFromServer is still null, a match was never found and we need to go and create the subscription
        //create subscription if it does not exist and resource does exist
        console.log(`subscription does not exist`)
        try {
            console.log('seeing if resource exists')
            //go and get the parent resource and see if it exists. (strip the /events off for calendar event subscriptions)
            await client.api(resource.name.split('/events')[0]).select('id').get();
            console.log('creating a subscription for the resource because it exists!')
            await client.api("/subscriptions").create({
                "changeType": resource.changeType,
                "notificationUrl": resource.url,
                "resource": resource.name,
                "expirationDateTime":newExpiration.toISOString(),
                "clientState": resource.secret,
                "latestSupportedTlsVersion": "v1_2"
            })
            if ('calendarId' in resource){
                await syncAllEventsInCalendar(resource.calendarId,client)
            }
        }catch(e){
            console.log(e)
        }
    }
    

}
//our sub matching function
var matchSubscription:(resourceToFind:any, subscriptions:Array<any>)=>any = function(resourcetoFind,subscriptions){
    let match = null;
    for (const sub of subscriptions){
        //if the current subscription name matches the resource name we are looking for, the subscription may already exist
        let nameMatches:boolean = sub.resource.toLowerCase() == resourcetoFind.name.toLowerCase()
        //if the current subscription has an app Id that matches the current app Id and the above is true, then the subscription does already exist
        let appIdMatches:boolean = sub.applicationId == process.env.CLIENT_ID
        let thisSubscriptionMatches = nameMatches && appIdMatches// returns true if both parameter equal each other

        if (thisSubscriptionMatches){
            match = sub;
            break
        }
    }
    return match
}
