import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as SpoWebRetentionLabelListCommand from '../web/web-retentionlabel-list';
const command: Command = require('./listitem-retentionlabel-ensure');

describe(commands.LISTITEM_RETENTIONLABEL_ENSURE, () => {

  //#region Mock Responses
  const mockSpoWebRetentionLabelListResponseArray = [
    {
      "AcceptMessagesOnlyFromSendersOrMembers": false,
      "AccessType": null,
      "AllowAccessFromUnmanagedDevice": null,
      "AutoDelete": true,
      "BlockDelete": true,
      "BlockEdit": false,
      "ComplianceFlags": 1,
      "ContainsSiteLabel": false,
      "DisplayName": "",
      "EncryptionRMSTemplateId": null,
      "HasRetentionAction": true,
      "IsEventTag": false,
      "MultiStageReviewerEmail": null,
      "NextStageComplianceTag": null,
      "Notes": null,
      "RequireSenderAuthenticationEnabled": false,
      "ReviewerEmail": null,
      "SharingCapabilities": null,
      "SuperLock": false,
      "TagDuration": 2555,
      "TagId": "def61080-111c-4aea-b72f-5b60e516e36c",
      "TagName": "Some label",
      "TagRetentionBasedOn": "CreationAgeInDays",
      "UnlockedAsDefault": false
    }
  ];
  //#endregion

  const webUrl = 'https://contoso.sharepoint.com';
  const listUrl = '/sites/project-x/list';
  const listTitle = 'test';
  const listId = 'b2307a39-e878-458b-bc90-03bc578531d6';
  const labelName = 'Some label';
  const labelId = 'def61080-111c-4aea-b72f-5b60e516e36c';

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerLogStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    loggerLogStderrSpy = sinon.spy(logger, 'logToStderr');
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoWebRetentionLabelListCommand) {
        return { stdout: JSON.stringify(mockSpoWebRetentionLabelListResponseArray) };
      }

      throw 'Unknown case';
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_RETENTIONLABEL_ENSURE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if listId, listTitle or listUrl option is not passed', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listItemId: 1, name: labelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listItemId: 1, listTitle: listTitle, name: labelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: '12345', listItemId: 1, name: labelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both listId and listTitle options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, listTitle: listTitle, listItemId: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listItemId is not passed', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listItemId is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listItemId: 'abc', listTitle: listTitle, name: labelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if name or id option is not passed', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listItemId: 1, listId: listId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: '12345', listItemId: 1, listId: listId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: labelId, listItemId: 1, listId: listId } }, commandInfo);
    assert(actual);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, listItemId: 1, name: labelName } }, commandInfo);
    assert(actual);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, listItemId: 1, name: labelName } }, commandInfo);
    assert(actual);
  });

  it('applies a retentionlabel based on listId and name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/items(1)/SetComplianceTag()`
        && JSON.stringify(opts.data) === '{"complianceTag":"Some label","isTagPolicyHold":true,"isTagPolicyRecord":false,"isEventBasedTag":false,"isTagSuperLock":false,"isUnlockedAsDefault":false}') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: false,
        listId: listId,
        webUrl: webUrl,
        listItemId: 1,
        name: labelName
      }
    });
    assert(loggerLogStderrSpy.notCalled);
  });

  it('applies a retentionlabel based on listId and id (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/items(1)/SetComplianceTag()`
        && JSON.stringify(opts.data) === '{"complianceTag":"Some label","isTagPolicyHold":true,"isTagPolicyRecord":false,"isEventBasedTag":false,"isTagSuperLock":false,"isUnlockedAsDefault":false}') {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        listId: listId,
        webUrl: webUrl,
        listItemId: 1,
        id: labelId
      }
    }));
  });

  it('applies a retentionlabel based on listTitle and id (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')/items(1)/SetComplianceTag()`
        && JSON.stringify(opts.data) === '{"complianceTag":"Some label","isTagPolicyHold":true,"isTagPolicyRecord":false,"isEventBasedTag":false,"isTagSuperLock":false,"isUnlockedAsDefault":false}') {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        listTitle: listTitle,
        webUrl: webUrl,
        listItemId: 1,
        id: labelId
      }
    }));
  });

  it('applies a retentionlabel based on listUrl and id (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetList(@a1)/items(1)/SetComplianceTag()?@a1='${formatting.encodeQueryParameter(listUrl)}'`
        && JSON.stringify(opts.data) === '{"complianceTag":"Some label","isTagPolicyHold":true,"isTagPolicyRecord":false,"isEventBasedTag":false,"isTagSuperLock":false,"isUnlockedAsDefault":false}') {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        listUrl: listUrl,
        webUrl: webUrl,
        listItemId: 1,
        id: labelId
      }
    }));
  });

  it('throws an error when a retentionlabel cannot be found on the site', async () => {
    await assert.rejects(command.action(logger, {
      options: {
        debug: false,
        listId: listId,
        webUrl: webUrl,
        listItemId: 1,
        name: 'Some non-existing label'
      }
    }), new CommandError('The specified retention label does not exist'));
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';

    sinon.stub(request, 'post').callsFake(async () => { throw { error: { error: { message: errorMessage } } }; });

    await assert.rejects(command.action(logger, {
      options: {
        debug: false,
        listUrl: listUrl,
        webUrl: 'https://contoso.sharepoint.com',
        listItemId: 1,
        name: labelName
      }
    }), new CommandError(errorMessage));
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });
});