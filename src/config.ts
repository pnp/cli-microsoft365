const packageJSON = require('../package.json');
const cliAadAppId: string = '31359c7f-bd7e-475c-86db-fdb8c937548e';

export default {
  applicationName: `SharePoint Framework CLI v${packageJSON.version}`,
  delimiter: 'o365\$',
  aadAadAppId: process.env.OFFICE365CLI_AADAADAPPID || process.env.OFFICE365CLI_AADAPPID || cliAadAppId,
  aadAzmgmtAppId: process.env.OFFICE365CLI_AADAZMGMTAPPID || process.env.OFFICE365CLI_AADAPPID || cliAadAppId,
  aadGraphAppId: process.env.OFFICE365CLI_AADGRAPHAPPID || process.env.OFFICE365CLI_AADAPPID || cliAadAppId,
  aadSpoAppId: process.env.OFFICE365CLI_AADSPOAPPID || process.env.OFFICE365CLI_AADAPPID || cliAadAppId
};