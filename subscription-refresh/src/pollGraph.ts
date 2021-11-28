/**
 * about this function
 *  - It will be triggered at regular intervals via cloudbridge (probably every 1-2 minutes)
 *  - it will retrieve the latest delta token from DynamoDB
 *  - it will poll MS Graph for changes using the delta token to find new changes in group membership
 *  - if there are new changes, we will process them accordingly
 */
import * as AWS from "aws-sdk";
import { MyAuthenticationProvider } from "./MyAuthenticationProvider";
import { Client, GraphRequest } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { getAllPagesAndDeltaTokenFromGraph } from "./Helpers";
import { processGroupChange } from "./GroupSync";
import { getCalendarGroupID, getUserID } from "./CalendarSync";
export type ResourceToPoll = {
  resource: string;
  fields: string[];
  filterFunction: (...args: any) => boolean;
  processChange: (resourceData: any, client: Client) => Promise<void>;
};
export const lambdaHandler = async (event: any): Promise<boolean> => {
  const options = {
    authProvider: new MyAuthenticationProvider(),
  };
  const graphClient: Client = await Client.initWithMiddleware(options);
  /**
   * get the last delta token from dynamo
   */
  AWS.config.update({
    maxRetries: 15,
    retryDelayOptions: { base: 500 },
  });

  const resourcesToPoll: ResourceToPoll[] = [
    {
      resource: "groups",
      fields: ["displayName", "members", "mail"],
      filterFunction: (group) =>
        group?.mail
          ?.substring(0, process.env.GROUP_EMAIL_PREPEND.length)
          ?.includes(process.env.GROUP_EMAIL_PREPEND) &&
        "members@delta" in group,
      processChange: processGroupChange,
    },
  ];
  const docClient = new AWS.DynamoDB.DocumentClient({ region: "us-east-2" });
  try {
    for (const resource of resourcesToPoll) {
      await pollResource(graphClient, docClient, resource);
    }
  } catch (e) {
    console.log(e);
  }

  /**
   * retrieve new changes from graph using delta token
   */
  /**
   * perform group sync using new changes
   */

  return true;
};
const pollResource = async (
  graphClient: Client,
  documentClient: AWS.DynamoDB.DocumentClient,
  resource: ResourceToPoll
) => {
  console.log("polling resource:", { resource: resource.resource });
  const params: AWS.DynamoDB.DocumentClient.ScanInput = {
    TableName: process.env.TABLE_NAME,
    FilterExpression: "#res = :res",
    ExpressionAttributeNames: {
      "#res": "resource",
    },
    ExpressionAttributeValues: {
      ":res": resource.resource,
    },
  };
  const items = await getAllItemsFromDDB(params, documentClient);
  if (items.length === 0) {
    const requestLink = `/${resource.resource}/delta${
      resource.fields.length > 0 ? `?$select=${resource.fields.join()}` : ""
    }`;
    console.log({ requestLink });
    const initialRequest = graphClient.api(requestLink);

    await getAndProcessFromGraphAndCreateDynamoDBRecord(
      graphClient,
      documentClient,
      initialRequest,
      resource
    );
  } else {
    // use delta param to get changes
    //iterate through all items and process all of them in case there's more than one
    for (const deltaRecord of items) {
      const { id, deltaLink } = deltaRecord;
      //get change data and create the next delta record in dynamo
      const nextGraphRequest = graphClient.api(deltaLink);
      await getAndProcessFromGraphAndCreateDynamoDBRecord(
        graphClient,
        documentClient,
        nextGraphRequest,
        resource
      );
      //clean up -- delete record from dynamo
      await documentClient
        .delete({
          TableName: process.env.TABLE_NAME,
          Key: { id },
        })
        .promise();
    }
  }
};

const getAllItemsFromDDB = async (
  params: AWS.DynamoDB.DocumentClient.ScanInput,
  documentClient: AWS.DynamoDB.DocumentClient
) => {
  const scanResults: any[] = [];
  let items;
  do {
    items = await documentClient.scan(params).promise();
    items.Items.forEach((item) => scanResults.push(item));
    params.ExclusiveStartKey = items.LastEvaluatedKey;
  } while (typeof items.LastEvaluatedKey !== "undefined");
  return scanResults;
};
const getAndProcessFromGraphAndCreateDynamoDBRecord = async (
  graphClient: Client,
  documentClient: AWS.DynamoDB.DocumentClient,
  graphRequest: GraphRequest,
  resource: ResourceToPoll
) => {
  const { deltaLink, allPages } = await getAllPagesAndDeltaTokenFromGraph(
    graphRequest
  );
  console.log({ allPagesLength: allPages.length });

  //get initial token from graph
  //process group membership
  //if it contains the email prepend, and it has members,process the group
  const resourceChangesToProcess = allPages.filter(resource.filterFunction);
  if (resourceChangesToProcess?.length === 0)
    console.log(`no ${resource.resource} changes to process`);
  for (const resourceChange of resourceChangesToProcess) {
    console.log(
      `processing ${resource.resource} change`,
      JSON.stringify(resourceChange)
    );
    await resource.processChange(resourceChange, graphClient);
  }
  //store the token in dynamoDB
  await documentClient
    .put({
      TableName: process.env.TABLE_NAME,
      Item: {
        id: `${new Date().toISOString()}`,
        deltaLink,
        resource: resource.resource,
        retrievalDate: new Date().getTime(),
      },
    })
    .promise();
};
