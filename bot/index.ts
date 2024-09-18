// Import required packages
import * as restify from "restify";
import { CosmosClient } from "@azure/cosmos";
import { v4 as uuidv4 } from "uuid";
import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
  TurnContext,
} from "botbuilder";

// This bot's main dialog.
import { TeamsBot } from "./teamsBot";
import config from "./config";
import { authenticate, getAccessTokenOnBehalfOf } from './auth';
import { Scan } from "./models/scan";

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: config.botId,
  MicrosoftAppPassword: config.botPassword,
  MicrosoftAppType: "MultiTenant",
});

const botFrameworkAuthentication = createBotFrameworkAuthenticationFromConfiguration(
  null,
  credentialsFactory
);

const adapter = new CloudAdapter(botFrameworkAuthentication);

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );

  // Send a message to the user
  await context.sendActivity(`The bot encountered unhandled error:\n ${error.message}`);
  await context.sendActivity("To continue to run this bot, please fix the bot source code.");
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create the bot that will handle incoming messages.
const bot = new TeamsBot();

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});

// CORS Middleware function
function corsMiddleware(req: restify.Request, res: restify.Response, next: restify.Next) {
  res.header('Access-Control-Allow-Origin', '*'); // Allows all origins
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS'); // Allowed methods
  res.header('Access-Control-Allow-Headers', 'Authorization, Content-Type'); // Allowed headers

  // Handle preflight requests
  if (req.method === 'OPTIONS') {
    return res.send(204);
  }

  return next();
};

server.pre(corsMiddleware);

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
    await bot.run(context);
  });
});

// gets all previous scans
server.get("/api/scans", authenticate, async (req, res) => {
  // get scans for the current user
  const cosmosClient = new CosmosClient(config.cosmosConn);
  const { database } = await cosmosClient.databases.createIfNotExists({ id: config.databaseName });
  const { container } = await database.containers.createIfNotExists({ id: config.containerName });

  // perform query
  const { resources: items } = await container.items
    .query({ query: "SELECT * from c WHERE c.user=@id", parameters: [{ name: "@id", value: (req as any).user.oid }] })
    .fetchAll();

  // start a scan if the user doesn't have one
  if (items.length === 0) {
    const newScan = await startScan((req as any).user.oid, (req as any).user.tid, req);
    items.push(newScan);
  }

  // sort by scanDate
  items.sort((a, b) => new Date(b.scanDate).getTime() - new Date(a.scanDate).getTime());


  // return items
  res.send(items);
});

// NOTE: if you need to hack a token from graph explorer, put it here, otherwise a valid SSO will be performed
const HACK_TOKEN = "";

// starts a duplicate file scan
server.post("/api/scans", authenticate, async (req, res) => {
  // start the scan
  const newScan = await startScan((req as any).user.oid, (req as any).user.tid, req);

  // send the new scan record
  res.send(newScan);
});

const startScan = async (userId:string, tenantId:string, req:restify.Request) : Promise<Scan> => {
  // create a new scan
  const newScan:Scan = {
    id: uuidv4(),
    status: "active",
    user: userId,
    scanDate: new Date().toISOString(),
    duplicates: []
  };

  // create the new record in the database
  const cosmosClient = new CosmosClient(config.cosmosConn);
  const { database } = await cosmosClient.databases.createIfNotExists({ id: config.databaseName });
  const { container } = await database.containers.createIfNotExists({ id: config.containerName });
  const { resource: createdItem } = await container.items.create(newScan);

  // perform on behalf of flow for Graph access token
  let token:string = HACK_TOKEN;
  if (token === "") {
    const authHeader = req.headers["authorization"] as string;
    const t = authHeader.split(" ")[1];
    token = await getAccessTokenOnBehalfOf(t, "https://graph.microsoft.com/.default");
  }

  // prepare event grid payload
  const payload = [{
    id: uuidv4(),
    eventType: "duplicateScanInitiated",
    subject: "/desup/filescan",
    eventTime: new Date().toISOString(),
    data: {
      user: userId,
      tenant: tenantId,
      scanId: newScan.id,
      token: token
    },
    dataVersion: "1.0"
  }];

  // Submit the scan request to EventGrid
  await fetch(config.eventGridEndpoint, {
    method: "POST",
    headers: {
      "aeg-sas-key": config.eventGridKey,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  // return the new scan
  return newScan;
};

// deletes a duplicate file
server.del("/api/files/:scanid", authenticate, async (req, res) => {
  // perform on behalf of flow for Graph access token
  let token:string = HACK_TOKEN;
  if (token === "") {
    const authHeader = req.headers["authorization"] as string;
    const t = authHeader.split(" ")[1];
    token = await getAccessTokenOnBehalfOf(t, "https://graph.microsoft.com/.default");
  }

  // delete the files from OneDrive
  for (var i = 0; i < req.body.locations.length; i++) {
    if (req.body.locations[i].path !== req.body.fileToKeep) {
      // TODO: this should use error handling and keep track of success
      
      // perform the delete against graph...could also use Graph SDK
      const result = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${req.body.locations[i].id}`, {
        method: "DELETE",
        headers: {
            "Authorization": `Bearer ${token}`,
            "Content-Type": "application/json",
        }
      });
    }
  }

  // create the new record in the database
  const cosmosClient = new CosmosClient(config.cosmosConn);
  const { database } = await cosmosClient.databases.createIfNotExists({ id: config.databaseName });
  const { container } = await database.containers.createIfNotExists({ id: config.containerName });

  // fetch the scan record
  const { resources: [item] } = await container.items
    .query({ query: "SELECT * from c WHERE c.id=@id", parameters: [{ name: "@id", value: req.params.scanid }] })
    .fetchAll();
  
  // remove the duplicate that has been fixed and update the scan record
  for (var i = 0; i < (item as Scan).duplicates.length; i++) {
    if ((item as Scan).duplicates[i].fileName === req.body.fileName) {
      (item as Scan).duplicates.splice(i, 1);
    }
  }
  await container.items.upsert(item);

  // send confirmation to client
  res.send({ message: "success" });
});
