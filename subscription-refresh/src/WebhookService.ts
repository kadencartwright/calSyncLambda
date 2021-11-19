import {
  APIGatewayEvent,
  APIGatewayProxyEvent,
  APIGatewayProxyResult,
} from "aws-lambda";
import * as aws from "aws-sdk";

export const lambdaHandler = async (
  event: APIGatewayEvent
): Promise<APIGatewayProxyResult> => {
  console.log(`webhookService started`);
  if (!!event.queryStringParameters) {
    const token = event.queryStringParameters["validationToken"];
    return {
      body: token,
      statusCode: 200,
      headers: { "content-type": "text/plain" },
    };
  }
  if (!!event.body) {
    let body;
    try {
      body = JSON.parse(event.body);
    } catch (e) {
      body = event.body;
    }
    if (!body.value) {
      return { body: "", statusCode: 202 };
    }

    let type = event.pathParameters["webhookType"];
    let payload = {
      type: type,
      data: body,
    };
    console.log(process.env.WEBHOOK_RESPONDER_NAME);
    let iRequest: aws.Lambda.InvocationRequest = {
      FunctionName: process.env.WEBHOOK_RESPONDER_NAME,
      Payload: JSON.stringify(payload),
      InvocationType: "Event",
    };
    console.log({ payload: JSON.stringify(payload) });
    console.log("creating invoke request");
    let data = await invoke(iRequest);
    console.log(data);
  }

  return {
    body: "success",
    statusCode: 200,
    headers: { "content-type": "text/plain" },
  };
};

async function invoke(params: aws.Lambda.InvocationRequest) {
  const lambda = new aws.Lambda({ region: process.env.AWS_REGION });

  return new Promise(function (resolve, reject) {
    lambda.invoke(params, function (err, data) {
      if (err) {
        reject(err);
      } else {
        resolve(data);
      }
    });
  });
}
