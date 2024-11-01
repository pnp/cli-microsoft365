import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './website-get.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { accessToken } from '../../../../utils/accessToken.js';

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

describe(commands.WEBSITE_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertDelegatedAccessToken').returns();
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
      powerPlatform.getWebsiteByName,
      powerPlatform.getWebsiteByUrl
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.WEBSITE_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name', 'websiteUrl', 'tenantId', 'subdomain', 'type', 'status', 'siteVisibility']);
  });

  it('retrieves the information for the Power Page website by url', async () => {
    sinon.stub(powerPlatform, 'getWebsiteByUrl').resolves(powerPageResponse);

    await command.action(logger, { options: { environmentName: environment, url: 'https://site-0uaq9.powerappsportals.com' } });
    assert(loggerLogSpy.calledWith(powerPageResponse));
  });

  it('retrieves the information for the Power Page website by name', async () => {
    sinon.stub(powerPlatform, 'getWebsiteByName').resolves(powerPageResponse);

    await command.action(logger, { options: { environmentName: environment, name: 'Site 1' } });
    assert(loggerLogSpy.calledWith(powerPageResponse));
  });

  it('retrieves the information for the Power Page website by id', async () => {
    sinon.stub(powerPlatform, 'getWebsiteById').resolves(powerPageResponse);

    await command.action(logger, { options: { environmentName: environment, id: '4916bb2c-91e1-4716-91d5-b6171928fac9' } });
    assert(loggerLogSpy.calledWith(powerPageResponse));
  });

  it('correctly handles error when getting information for a site that doesn\'t exist', async () => {
    sinon.stub(powerPlatform, 'getWebsiteByName').callsFake(() => { throw new Error('The specified Power Page website \'Site 1\' does not exist.'); });

    await assert.rejects(command.action(logger, { options: { verbose: true, environmentName: environment, name: 'Site 1' } } as any), new CommandError('The specified Power Page website \'Site 1\' does not exist.'));
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: environment, url: 'https://site-0uaq9.contoso.com' });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: environment, url: 'https://site-0uaq9.powerappsportals.com' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if url and name are used at the same time', () => {
    const actual = commandOptionsSchema.safeParse({
      environmentName: environment,
      url: 'https://site-0uaq9.powerappsportals.com',
      name: 'Site 1'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if url and id are used at the same time', () => {
    const actual = commandOptionsSchema.safeParse({
      environmentName: environment,
      url: 'https://site-0uaq9.powerappsportals.com',
      id: '4916bb2c-91e1-4716-91d5-b6171928fac9'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if name and id are used at the same time', () => {
    const actual = commandOptionsSchema.safeParse({
      environmentName: environment,
      id: '4916bb2c-91e1-4716-91d5-b6171928fac9',
      name: 'Site 1'
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation with only url', () => {
    const actual = commandOptionsSchema.safeParse({
      environmentName: environment,
      url: 'https://site-0uaq9.powerappsportals.com'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with only id', () => {
    const actual = commandOptionsSchema.safeParse({
      environmentName: environment,
      id: '4916bb2c-91e1-4716-91d5-b6171928fac9'
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with only name', () => {
    const actual = commandOptionsSchema.safeParse({
      environmentName: environment,
      name: 'Site 1'
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if url, id, and name are all used at the same time', () => {
    const actual = commandOptionsSchema.safeParse({
      environmentName: environment,
      url: 'https://site-0uaq9.powerappsportals.com',
      id: '4916bb2c-91e1-4716-91d5-b6171928fac9',
      name: 'Site 1'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither url, id, nor name are provided', () => {
    const actual = commandOptionsSchema.safeParse({
      environmentName: environment
    });
    assert.strictEqual(actual.success, false);
  });
});
