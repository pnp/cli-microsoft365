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
import { urlUtil } from '../../../../utils/urlUtil';
import { formatting } from '../../../../utils/formatting';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as SpoFileGetCommand from './file-get';
import * as SpoGroupGetCommand from '../group/group-get';
import * as SpoUserGetCommand from '../user/user-get';
const command: Command = require('./file-roleassignment-remove');

describe(commands.FILE_ROLEASSIGNMENT_REMOVE, () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/contoso-sales';
  const fileUrl = '/sites/contoso-sales/documents/Test1.docx';
  const fileId = 'b2307a39-e878-458b-bc90-03bc578531d6';
  const principalId = 2;
  const upn = 'user1@contoso.onmicrosoft.com';
  const groupName = 'saleGroup';

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
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
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      Cli.prompt,
      Cli.executeCommandWithOutput,
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.FILE_ROLEASSIGNMENT_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [
      { options: ['fileUrl', 'fileId'] },
      { options: ['upn', 'groupName', 'principalId'] }
    ]);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileId: fileId, principalId: principalId, confirm: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: 'foo', principalId: principalId, confirm: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the principalId option is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, principalId: 'Hi', confirm: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if webUrl and fileId are valid', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: '0cd891ef-afce-4e55-b836-fce03286cccf', principalId: principalId, confirm: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing role assignment from the file when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        webUrl: webUrl,
        fileId: fileId,
        principalId: principalId
      }
    });

    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing role assignment from the file when confirm option is not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');

    await command.action(logger, {
      options: {
        webUrl: webUrl,
        fileId: fileId,
        principalId: principalId
      }
    });

    assert(postSpy.notCalled);
  });

  it('remove role assignment from the file by relative URL and principal Id when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, fileUrl);
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields/roleassignments/removeroleassignment(principalid='${principalId}')`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        fileUrl: fileUrl,
        principalId: principalId
      }
    });
  });

  it('remove role assignment from the file by Id and upn', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoFileGetCommand) {
        return {
          stdout: `{"LinkingUri": "https://contoso.sharepoint.com/sites/contoso-sales/documents/Test1.docx?d=wc39926a80d2c4067afa6cff9902eb866","Name": "Test1.docx","ServerRelativeUrl": "/sites/contoso-sales/documents/Test1.docx","UniqueId": "b2307a39-e878-458b-bc90-03bc578531d6"}`
        };
      }

      if (command === SpoUserGetCommand) {
        return {
          stdout: '{"Id": 2,"IsHiddenInUI": false,"LoginName": "i:0#.f|membership|user1@contoso.onmicrosoft.com","Title": "User1","PrincipalType": 1,"Email": "user1@contoso.onmicrosoft.com","Expiration": "","IsEmailAuthenticationGuestUser": false,"IsShareByEmailGuestUser": false,"IsSiteAdmin": true,"UserId": {"NameId": "1003200097d06dd6","NameIdIssuer": "urn:federation:microsoftonline"},"UserPrincipalName": "user1@contoso.onmicrosoft.com"}'
        };
      }

      throw new CommandError('Unknown case');
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, fileUrl);
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields/roleassignments/removeroleassignment(principalid='${principalId}')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        fileId: fileId,
        upn: upn,
        confirm: true
      }
    });
  });

  it('remove role assignment from the file by Id and group name', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoFileGetCommand) {
        return {
          stdout: `{"LinkingUri": "https://contoso.sharepoint.com/sites/contoso-sales/documents/Test1.docx?d=wc39926a80d2c4067afa6cff9902eb866","Name": "Test1.docx","ServerRelativeUrl": "/sites/contoso-sales/documents/Test1.docx","UniqueId": "b2307a39-e878-458b-bc90-03bc578531d6"}`
        };
      }

      if (command === SpoGroupGetCommand) {
        return {
          stdout: '{"Id": 2,"IsHiddenInUI": false,"LoginName": "saleGroup","Title": "saleGroup","PrincipalType": 8,"AllowMembersEditMembership": false,"AllowRequestToJoinLeave": false,"AutoAcceptRequestToJoinLeave": false,"Description": "","OnlyAllowMembersViewMembership": true,"OwnerTitle": "Some Account","RequestToJoinLeaveEmailSetting": null}'
        };
      }

      throw new CommandError('Unknown case');
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, fileUrl);
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields/roleassignments/removeroleassignment(principalid='${principalId}')`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        fileId: fileId,
        groupName: groupName,
        confirm: true
      }
    });
  });

  it('correctly handles error when removing file role assignment', async () => {
    const errorMessage = 'request rejected';
    sinon.stub(request, 'post').callsFake(async () => { throw errorMessage; });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        webUrl: webUrl,
        fileUrl: fileUrl,
        principalId: principalId,
        confirm: true
      }
    }), new CommandError(errorMessage));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });
});
