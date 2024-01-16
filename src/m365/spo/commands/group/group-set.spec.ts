import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './group-set.js';

const validId = 1;
const validName = "Project leaders";
const validWebUrl = 'https://contoso.sharepoint.com/sites/project-x';
const validOwnerEmail = 'john.doe@contoso.com';
const validOwnerUserName = 'john.doe@contoso.com';

const userInfoResponse = {
  userPrincipalName: validOwnerUserName
};

const ensureUserResponse = {
  Id: 3
};

describe(commands.GROUP_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.patch,
      cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when group id is not a number', async () => {
    const actual = await command.validate({
      options: {
        webUrl: validWebUrl,
        id: 'invalid id'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both ownerEmail and ownerUserName are specified', async () => {
    const actual = await command.validate({
      options: {
        webUrl: validWebUrl,
        id: validId,
        ownerEmail: validOwnerEmail,
        ownerUserName: validOwnerUserName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid web URL is passed', async () => {
    const actual = await command.validate({
      options: {
        webUrl: 'invalid',
        id: validId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid options specified', async () => {
    const actual = await command.validate({
      options: {
        webUrl: validWebUrl,
        id: validId,
        allowRequestToJoinLeave: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('successfully updates group settings by id', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetById(${validId})`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        id: validId,
        allowRequestToJoinLeave: true,
        verbose: true
      }
    });
  });

  it('successfully updates group settings by name', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetByName('${validName}')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        name: validName,
        allowRequestToJoinLeave: true,
        verbose: true
      }
    });
  });

  it('successfully updates group owner by ownerEmail, retrieves group by id', async () => {
    sinon.stub(cli, 'executeCommandWithOutput').resolves({
      stdout: JSON.stringify(userInfoResponse),
      stderr: ''
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetById(${validId})`) {
        return;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureUser('${userInfoResponse.userPrincipalName}')?$select=Id`) {
        return ensureUserResponse;
      }

      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetById(${validId})/SetUserAsOwner(${ensureUserResponse.Id})`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        id: validId,
        ownerEmail: validOwnerEmail,
        verbose: true
      }
    });
  });

  it('successfully updates group owner by ownerUserName, retrieves group by name', async () => {
    sinon.stub(cli, 'executeCommandWithOutput').resolves({
      stdout: JSON.stringify(userInfoResponse),
      stderr: ''
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetByName('${validName}')`) {
        return;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureUser('${userInfoResponse.userPrincipalName}')?$select=Id`) {
        return ensureUserResponse;
      }

      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetByName('${validName}')/SetUserAsOwner(${ensureUserResponse.Id})`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        name: validName,
        ownerUserName: validOwnerUserName,
        verbose: true
      }
    });
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetByName('${validName}')`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: validWebUrl,
        name: validName,
        autoAcceptRequestToJoinLeave: true,
        verbose: true
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });
});
