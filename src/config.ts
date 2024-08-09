import { app } from "./utils/app.js";

const cliEntraAppId: string = '31359c7f-bd7e-475c-86db-fdb8c937548e';

export default {
  applicationName: `CLI for Microsoft 365 v${app.packageJson().version}`,
  delimiter: 'm365\$',
  cliEntraAppId: process.env.CLIMICROSOFT365_ENTRAAPPID || cliEntraAppId,
  tenant: process.env.CLIMICROSOFT365_TENANT || 'common',
  configstoreName: 'cli-m365-config'
};