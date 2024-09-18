const config = {
  botId: process.env.BOT_ID,
  botPassword: process.env.BOT_PASSWORD,
  clientId: process.env.AAD_APP_CLIENT_ID,
  clientSecret: process.env.AAD_APP_CLIENT_SECRET,
  tenantId: process.env.TEAMS_APP_TENANT_ID,
  cosmosConn: process.env.COSMOS_CONN_STRING,
  databaseName: process.env.COSMOS_DATABASE_NAME,
  containerName: process.env.COSMOS_CONTAINER_NAME,
  eventGridEndpoint: process.env.EG_ENDPOINT,
  eventGridKey: process.env.EG_KEY,
};

export default config;
