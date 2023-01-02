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
import * as SpoUserGetCommand from '../user/user-get';
import * as SpoGroupGetCommand from '../group/group-get';
const command: Command = require('./web-roleassignment-remove');

describe(commands.WEB_ROLEASSIGNMENT_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];
  let promptOptions: any;

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
    requests = [];
    sinon.stub(Cli, 'prompt').callsFake(async (options) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.executeCommandWithOutput,
      Cli.prompt
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
    assert.strictEqual(command.name.startsWith(commands.WEB_ROLEASSIGNMENT_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', principalId: 11 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', principalId: 11 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the principalId option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', principalId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the principalId option is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', principalId: 11 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('remove role assignment from web', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('_api/web/roleassignments/removeroleassignment(principalid=\'11\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        principalId: 11,
        confirm: true
      }
    });
  });

  it('remove role assignment from web get principal id by upn', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/roleassignments/removeroleassignment(principalid=\'11\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === SpoUserGetCommand) {
        return Promise.resolve({
          stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "i:0#.f|membership|someaccount@tenant.onmicrosoft.com","Title": "Some Account","PrincipalType": 1,"Email": "someaccount@tenant.onmicrosoft.com","Expiration": "","IsEmailAuthenticationGuestUser": false,"IsShareByEmailGuestUser": false,"IsSiteAdmin": true,"UserId": {"NameId": "1003200097d06dd6","NameIdIssuer": "urn:federation:microsoftonline"},"UserPrincipalName": "someaccount@tenant.onmicrosoft.com"}'
        });
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        upn: 'someaccount@tenant.onmicrosoft.com',
        confirm: true
      }
    });
  });

  it('correctly handles error when upn does not exist', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/roleassignments/removeroleassignment(principalid=\'11\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    const error = 'no user found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === SpoUserGetCommand) {
        return Promise.reject(error);
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        upn: 'someaccount@tenant.onmicrosoft.com',
        confirm: true
      }
    } as any), new CommandError(error));
  });

  it('remove role assignment from web get principal id by group name', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/roleassignments/removeroleassignment(principalid=\'11\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === SpoGroupGetCommand) {
        return Promise.resolve({
          stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "otherGroup","Title": "otherGroup","PrincipalType": 8,"AllowMembersEditMembership": false,"AllowRequestToJoinLeave": false,"AutoAcceptRequestToJoinLeave": false,"Description": "","OnlyAllowMembersViewMembership": true,"OwnerTitle": "Some Account","RequestToJoinLeaveEmailSetting": null}'
        });
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        groupName: 'someGroup',
        confirm: true
      }
    });
  });

  it('correctly handles error when group does not exist', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/roleassignments/removeroleassignment(principalid=\'11\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    const error = 'no group found';
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === SpoGroupGetCommand) {
        return Promise.reject(error);
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        groupName: 'someGroup',
        confirm: true
      }
    } as any), new CommandError(error));
  });

  it('aborts removing role assignment when prompt not confirmed', async () => {
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        groupName: 'someGroup'
      }
    });
    assert(requests.length === 0);
  });

  it('prompts before removing role assignment when confirmation argument not passed', async () => {
    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        groupName: 'someGroup'
      }
    });

    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }
    assert(promptIssued);
  });

  it('removes role assignment when prompt confirmed', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/roleassignments/removeroleassignment(principalid=\'11\')') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === SpoGroupGetCommand) {
        return Promise.resolve({
          stdout: '{"Id": 11,"IsHiddenInUI": false,"LoginName": "otherGroup","Title": "otherGroup","PrincipalType": 8,"AllowMembersEditMembership": false,"AllowRequestToJoinLeave": false,"AutoAcceptRequestToJoinLeave": false,"Description": "","OnlyAllowMembersViewMembership": true,"OwnerTitle": "Some Account","RequestToJoinLeaveEmailSetting": null}'
        });
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        groupName: 'someGroup'
      }
    });
  });
});
