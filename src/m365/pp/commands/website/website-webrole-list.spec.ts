import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
import { pid } from '../../../../utils/pid.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './website-webrole-list.js';

const environment = 'Default-727dc1e9-3cd1-4d1f-8102-ab5c936e52f0';
const powerPageResponse = {
  "@odata.metadata": "https://api.powerplatform.com/powerpages/environments/Default-727dc1e9-3cd1-4d1f-8102-ab5c936e52f0/websites/$metadata#Websites",
  "id": "4916bb2c-91e1-4716-91d5-b6171928fac9",
  "name": "Site 1",
  "createdOn": "2024-10-27T12:00:03",
  "templateName": "DefaultPortalTemplate",
  "websiteUrl": "https://site-0uaq9.powerappsportals.com",
  "tenantId": "727dc1e9-3cd1-4d1f-8102-ab5c936e52f0",
  "dataverseInstanceUrl": "https://org0cd4b2b9.crm4.dynamics.com/",
  "environmentName": "Contoso (default)",
  "environmentId": "Default-727dc1e9-3cd1-4d1f-8102-ab5c936e52f0",
  "dataverseOrganizationId": "2d58aeac-74d4-4939-98d1-e05a70a655ba",
  "selectedBaseLanguage": 1033,
  "customHostNames": [],
  "websiteRecordId": "5eb107a6-5ac2-4e1c-a3b9-d5c21bbc10ce",
  "subdomain": "site-0uaq9",
  "packageInstallStatus": "Installed",
  "type": "Trial",
  "trialExpiringInDays": 86,
  "suspendedWebsiteDeletingInDays": 93,
  "packageVersion": "9.6.9.39",
  "isEarlyUpgradeEnabled": false,
  "isCustomErrorEnabled": true,
  "applicationUserAadAppId": "3f57aca7-5051-41b2-989d-26da8af7a53e",
  "ownerId": "33469a62-c3af-4cfe-b893-854eceab96da",
  "status": "OperationComplete",
  "siteVisibility": "private",
  "dataModel": "Enhanced"
};

const webroleResponse = [
  {
    "mspp_webroleid": "a242a363-6077-4cb7-b2d1-1714502d129a",
    "mspp_name": "Anonymous Users",
    "mspp_description": null,
    "mspp_key": null,
    "mspp_authenticatedusersrole": false,
    "mspp_anonymoususersrole": true,
    "mspp_createdon": "2026-01-21T22:10:56Z",
    "mspp_modifiedon": "2026-01-21T22:10:56Z",
    "statecode": 0,
    "statuscode": 1,
    "_mspp_websiteid_value": "5eb107a6-5ac2-4e1c-a3b9-d5c21bbc10ce",
    "_mspp_createdby_value": "b7aa2026-a8c1-f011-bbd2-000d3a66196e",
    "_mspp_modifiedby_value": "b7aa2026-a8c1-f011-bbd2-000d3a66196e"
  }
];

describe(commands.WEBSITE_WEBROLE_LIST, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertAccessTokenType').returns();
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      powerPlatform.getWebsiteById,
      powerPlatform.getWebsiteIdByUniqueName,
      powerPlatform.getDynamicsInstanceApiUrl,
      odata.getAllItems
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.WEBSITE_WEBROLE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if websiteName and websiteId are used at the same time', () => {
    const actual = commandOptionsSchema.safeParse({
      environmentName: environment,
      websiteId: '4916bb2c-91e1-4716-91d5-b6171928fac9',
      websiteName: 'Site 1'
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation with only websiteId', () => {
    const actual = commandOptionsSchema.safeParse({
      environmentName: environment,
      websiteId: '4916bb2c-91e1-4716-91d5-b6171928fac9'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with only websiteName', () => {
    const actual = commandOptionsSchema.safeParse({
      environmentName: environment,
      websiteName: 'Site 1'
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if neither websiteId, websiteName are provided', () => {
    const actual = commandOptionsSchema.safeParse({
      environmentName: environment
    });
    assert.strictEqual(actual.success, false);
  });

  it('retrieves webroles for Power Pages site by id', async () => {
    sinon.stub(powerPlatform, 'getWebsiteById').resolves(powerPageResponse);
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves('https://org0cd4b2b9.crm4.dynamics.com');
    sinon.stub(odata, 'getAllItems').resolves(webroleResponse);

    await command.action(logger, { options: { verbose: true, environmentName: environment, websiteId: '4916bb2c-91e1-4716-91d5-b6171928fac9' } });
    assert(loggerLogSpy.calledWith(webroleResponse));
  });

  it('retrieves webroles for Power Pages site by name', async () => {
    sinon.stub(powerPlatform, 'getWebsiteIdByUniqueName').resolves('5eb107a6-5ac2-4e1c-a3b9-d5c21bbc10ce');
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves('https://org0cd4b2b9.crm4.dynamics.com');
    sinon.stub(odata, 'getAllItems').resolves(webroleResponse);

    await command.action(logger, { options: { verbose: true, environmentName: environment, websiteName: 'Site 1' } });
    assert(loggerLogSpy.calledWith(webroleResponse));
  });

  it('outputs text friendly output when output is text', async () => {
    sinon.stub(powerPlatform, 'getWebsiteById').resolves(powerPageResponse);
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves('https://org0cd4b2b9.crm4.dynamics.com');
    sinon.stub(odata, 'getAllItems').resolves(webroleResponse);

    await command.action(logger, { options: { environmentName: environment, websiteId: '4916bb2c-91e1-4716-91d5-b6171928fac9', output: 'text' } });
    assert(loggerLogSpy.calledWith([
      {
        webroleid: 'a242a363-6077-4cb7-b2d1-1714502d129a',
        name: 'Anonymous Users',
        statuscode: 1
      }
    ]));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(powerPlatform, 'getWebsiteById').resolves(powerPageResponse);
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves('https://org0cd4b2b9.crm4.dynamics.com');
    sinon.stub(odata, 'getAllItems').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { environmentName: environment, websiteId: '4916bb2c-91e1-4716-91d5-b6171928fac9' } }),
      new CommandError('An error has occurred'));
  });
});