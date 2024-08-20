import { app } from './utils/app.js';

export default {
  applicationName: `CLI for Microsoft 365 v${app.packageJson().version}`,
  delimiter: 'm365\$',
  configstoreName: 'cli-m365-config'
};