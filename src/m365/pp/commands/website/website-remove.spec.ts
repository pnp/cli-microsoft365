import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { pid } from '../../../../utils/pid.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './website-remove.js';
import request from '../../../../request.js';

const environment = 'Default-727dc1e9-3cd1-4d1f-8102-ab5c936e52f0';
const websiteId = '4916bb2c-91e1-4716-91d5-b6171928fac9';
const websiteName = 'Site 1';
const websiteUrl = 'https://site-0uaq9.powerappsportals.com';

const powerPageResponse = {
  id: websiteId,
  name: websiteName,
  createdOn: "2024-10-27T12:00:03",
  templateName: "DefaultPortalTemplate",
  websiteUrl: websiteUrl,
  tenantId: "727dc1e9-3cd1-4d1f-8102-ab5c936e52f0",
  dataverseInstanceUrl: "https://org0cd4b2b9.crm4.dynamics.com/",
  environmentName: "Contoso (default)",
  environmentId: environment,
  dataverseOrganizationId: "2d58aeac-74d4-4939-98d1-e05a70a655ba",
  selectedBaseLanguage: 1033,
  customHostNames: [],
  websiteRecordId: "5eb107a6-5ac2-4e1c-a3b9-d5c21bbc10ce",
  subdomain: "site-0uaq9",
  packageInstallStatus: "Installed",
  type: "Trial",
  trialExpiringInDays: 86,
  suspendedWebsiteDeletingInDays: 93,
  packageVersion: "9.6.9.39",
  isEarlyUpgradeEnabled: false,
  isCustomErrorEnabled: true,
  applicationUserAadAppId: "3f57aca7-5051-41b2-989d-26da8af7a53e",
  ownerId: "33469a62-c3af-4cfe-b893-854eceab96da",
  status: "OperationComplete",
  siteVisibility: "private",
  dataModel: "Enhanced"
};

describe(commands.WEBSITE_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let promptIssued: boolean = false;

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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      powerPlatform.getWebsiteById,
      powerPlatform.getWebsiteByName,
      powerPlatform.getWebsiteByUrl,
      request.delete,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.WEBSITE_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid Power Pages site URL', async () => {
    const actual = commandOptionsSchema.safeParse({ environmentName: environment, url: 'https://site-0uaq9.contoso.com' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the url option is a valid Power Pages site URL', async () => {
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

  it('removes the Power Pages website by id with force option', async () => {
    const deleteStub = sinon.stub(request, 'delete').resolves();

    await command.action(logger, {
      options: {
        environmentName: environment,
        id: websiteId,
        force: true
      }
    });

    assert(deleteStub.calledOnce);
  });

  it('removes the Power Pages website by name when prompt confirmed', async () => {
    sinon.stub(powerPlatform, 'getWebsiteByName').resolves(powerPageResponse);
    sinon.stub(request, 'delete').resolves();

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        verbose: true,
        environmentName: environment,
        name: websiteName
      }
    });

    assert(loggerLogToStderrSpy.called);
  });

  it('removes the Power Pages website by id when prompt confirmed with verbose', async () => {
    sinon.stub(request, 'delete').resolves();

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        verbose: true,
        environmentName: environment,
        id: websiteId
      }
    });

    assert(loggerLogToStderrSpy.called);
  });

  it('removes the Power Pages website by url when prompt confirmed with verbose', async () => {
    sinon.stub(powerPlatform, 'getWebsiteByUrl').resolves(powerPageResponse);
    sinon.stub(request, 'delete').resolves();

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        verbose: true,
        environmentName: environment,
        url: websiteUrl
      }
    });

    assert(loggerLogToStderrSpy.called);
  });

  it('removes the Power Pages website by url with force option', async () => {
    sinon.stub(powerPlatform, 'getWebsiteByUrl').resolves(powerPageResponse);
    const deleteStub = sinon.stub(request, 'delete').resolves();

    await command.action(logger, {
      options: {
        environmentName: environment,
        url: websiteUrl,
        force: true
      }
    });

    assert(deleteStub.calledOnce);
  });

  it('does not remove website when prompt is not confirmed', async () => {
    const deleteStub = sinon.stub(request, 'delete').resolves();

    await command.action(logger, {
      options: {
        environmentName: environment,
        id: websiteId
      }
    });

    assert(deleteStub.notCalled);
  });

  it('prompts before removing website by url', async () => {
    sinon.stub(powerPlatform, 'getWebsiteByUrl').resolves(powerPageResponse);

    await command.action(logger, {
      options: {
        environmentName: environment,
        url: websiteUrl
      }
    });

    assert(promptIssued);
  });

  it('correctly handles error when removing website fails', async () => {
    const errorMessage = 'An error has occurred';
    sinon.stub(request, 'delete').rejects(new Error(errorMessage));

    await assert.rejects(
      command.action(logger, {
        options: {
          environmentName: environment,
          id: websiteId,
          force: true
        }
      }),
      new CommandError(errorMessage)
    );
  });
});