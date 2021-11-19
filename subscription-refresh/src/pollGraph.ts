/**
 * about this function
 *  - It will be triggered at regular intervals via cloudbridge (probably every 1-2 minutes)
 *  - it will retrieve the latest delta token from DynamoDB
 *  - it will poll MS Graph for changes using the delta token to find new changes in group membership
 *  - if there are new changes, we will process them accordingly
 */
import * as AWS from "aws-sdk";
import { MyAuthenticationProvider } from "./MyAuthenticationProvider";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";

export const lambdaHandler = async (event: any): Promise<boolean> => {
  const options = {
    authProvider: new MyAuthenticationProvider(),
  };
  const client: Client = await Client.initWithMiddleware(options);
  /**
   * get the last delta token from dynamo
   */
  const docClient = new AWS.DynamoDB.DocumentClient({ region: "us-east-2" });
  const params = {
    TableName: process.env.TABLE_NAME,
    Key: {
      type: "deltaToken",
    },
    Limit: 1,
    ScanIndexForward: false, // true = ascending, false = descending
  };

  const result = await docClient.get(params);
  console.log(JSON.stringify(result));
  // if not exist, get the initial token from graph

  /**
   * retrieve new changes from graph using delta token
   */
  /**
   * perform group sync using new changes
   */

  return true;
};
