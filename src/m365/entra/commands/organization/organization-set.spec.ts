import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './organization-set.js';

describe(commands.ORGANIZATION_SET, () => {
  const organizationId = '84841066-274d-4ec0-a5c1-276be684bdd3';
  const organizationName = 'Contoso';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch,
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ORGANIZATION_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'foo',
      marketingNotificationEmails: 'marketing@contoso.com'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both id and displayName are provided', () => {
    const actual = commandOptionsSchema.safeParse({
      id: organizationId,
      displayName: organizationName,
      marketingNotificationEmails: 'marketing@contoso.com'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither id nor displayName is provided', () => {
    const actual = commandOptionsSchema.safeParse({
      marketingNotificationEmails: 'marketing@contoso.com'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if contactEmail is not a valid email', () => {
    const actual = commandOptionsSchema.safeParse({
      id: organizationId,
      contactEmail: 'contactcontosocom'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if marketingNotificationEmails contains invalid email', () => {
    const actual = commandOptionsSchema.safeParse({
      id: organizationId,
      marketingNotificationEmails: 'marketing@contoso.com,foocontoso.com'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if securityComplianceNotificationMails contains invalid email', () => {
    const actual = commandOptionsSchema.safeParse({
      id: organizationId,
      securityComplianceNotificationMails: 'security@contoso.com,foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if technicalNotificationMails contains invalid email', () => {
    const actual = commandOptionsSchema.safeParse({
      id: organizationId,
      technicalNotificationMails: 'support@contoso.com,@contoso.com'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if neither contactEmail, marketingNotificationEmails, securityComplianceNotificationMails, securityComplianceNotificationPhones, statementUrl, nor technicalNotificationMails is provided', () => {
    const actual = commandOptionsSchema.safeParse({ id: organizationId });
    assert.notStrictEqual(actual.success, true);
  });

  it('updates an organization specified by id', async () => {
    const patchRequestStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/organization/${organizationId}`) {
        return;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      id: organizationId,
      marketingNotificationEmails: 'marketing@contoso.com',
      securityComplianceNotificationMails: 'security@contoso.com',
      securityComplianceNotificationPhones: '(123) 456-7890, (987) 654-3210',
      technicalNotificationMails: 'it@contoso.com,support@contoso.com',
      contactEmail: 'contact@contoso.com',
      statementUrl: 'https://contoso.com/privacyStatement',
      verbose: true
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert(patchRequestStub.called);
  });

  it('updates an organization specified by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/organization?$select=id,displayName`) {
        return {
          value: [{
            id: organizationId,
            displayName: organizationName
          }]
        };
      }

      throw 'Invalid request';
    });

    const patchRequestStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/organization/${organizationId}`) {
        return;
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      displayName: organizationName,
      marketingNotificationEmails: 'marketing@contoso.com',
      securityComplianceNotificationMails: 'security@contoso.com',
      securityComplianceNotificationPhones: '(123) 456-7890, (987) 654-3210',
      technicalNotificationMails: 'it@contoso.com,support@contoso.com',
      contactEmail: 'contact@contoso.com',
      statementUrl: 'https://contoso.com/privacyStatement'
    });
    await command.action(logger, { options: parsedSchema.data! });
    assert(patchRequestStub.called);
  });

  it('throws error when no organization specified by name was found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/organization?$select=id,displayName`) {
        return {
          value: [{
            id: organizationId,
            displayName: 'foo'
          }]
        };
      }

      throw 'Invalid request';
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      displayName: organizationName,
      marketingNotificationEmails: 'marketing@contoso.com',
      securityComplianceNotificationMails: 'security@contoso.com',
      securityComplianceNotificationPhones: '(123) 456-7890, (987) 654-3210',
      technicalNotificationMails: 'it@contoso.com,support@contoso.com',
      contactEmail: 'contact@contoso.com',
      statementUrl: 'https://contoso.com/privacyStatement'
    });

    await assert.rejects(command.action(logger, { options: parsedSchema.data! }), new CommandError(`The specified organization '${organizationName}' does not exist.`));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'patch').rejects({
      error: {
        code: "Request_BadRequest",
        message: "Invalid tenant identifier; it must match that of the requested tenant.",
        innerError: {
          date: "2025-05-23T11:36:44",
          'request-id': "fa792713-8c17-48ae-aaeb-2d9653954815",
          'client-request-id': "101755c1-8c5a-140c-97b6-975938bc6b5d"
        }
      }
    });

    const parsedSchema = commandOptionsSchema.safeParse({
      id: organizationId,
      marketingNotificationEmails: 'marketing@contoso.com'
    });
    await assert.rejects(command.action(logger, {
      options: parsedSchema.data!
    }), new CommandError('Invalid tenant identifier; it must match that of the requested tenant.'));
  });
});