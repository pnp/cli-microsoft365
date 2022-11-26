import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./group-set');

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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.patch,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [{ options: ['id', 'name'] }]);
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
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetById(${validId})`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        id: validId,
        allowRequestToJoinLeave: true
      }
    });
  });

  it('successfully updates group settings by name', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetByName('${validName}')`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        name: validName,
        allowRequestToJoinLeave: true
      }
    });
  });

  it('successfully updates group owner by ownerEmail, retrieves group by id', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify(userInfoResponse),
      stderr: ''
    }));
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetById(${validId})`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureUser('${userInfoResponse.userPrincipalName}')?$select=Id`) {
        return Promise.resolve(ensureUserResponse);
      }

      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetById(${validId})/SetUserAsOwner(${ensureUserResponse.Id})`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        id: validId,
        ownerEmail: validOwnerEmail
      }
    });
  });

  it('successfully updates group owner by ownerUserName, retrieves group by name', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(() => Promise.resolve({
      stdout: JSON.stringify(userInfoResponse),
      stderr: ''
    }));
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetByName('${validName}')`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/ensureUser('${userInfoResponse.userPrincipalName}')?$select=Id`) {
        return Promise.resolve(ensureUserResponse);
      }

      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetByName('${validName}')/SetUserAsOwner(${ensureUserResponse.Id})`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    await command.action(logger, {
      options: {
        webUrl: validWebUrl,
        name: validName,
        ownerUserName: validOwnerUserName
      }
    });
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `${validWebUrl}/_api/web/sitegroups/GetByName('${validName}')`) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid Request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: validWebUrl,
        name: validName,
        autoAcceptRequestToJoinLeave: true
      }
    }), new CommandError('An error has occurred'));
  });
});
