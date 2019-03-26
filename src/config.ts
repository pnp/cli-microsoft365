const packageJSON = require('../package.json');
const cliAadAppId: string = '31359c7f-bd7e-475c-86db-fdb8c937548e';

export default {
  applicationName: `SharePoint Framework CLI v${packageJSON.version}`,
  delimiter: 'o365\$',
  cliAadAppId: process.env.OFFICE365CLI_AADAPPID || cliAadAppId,
  tenant: process.env.OFFICE365CLI_TENANT || 'common'
};