import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './listitem-retentionlabel-ensure.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.LISTITEM_RETENTIONLABEL_ENSURE, () => {

  //#region Mock Responses
  const retentionLabelListMock = [
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
      "IsEventTag": true,
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

  const retentionLabelListMockResponse = {
    "odata.metadata": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Collection(SP.CompliancePolicy.ComplianceTag)",
    "value": retentionLabelListMock
  };

  const listDetailsMock = {
    "RootFolder": {
      "ServerRelativeUrl": "/sites/project-x/list"
    }
  };
  //#endregion

  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listUrl = '/sites/project-x/list';
  const listTitle = 'test';
  const listId = 'b2307a39-e878-458b-bc90-03bc578531d6';
  const labelName = 'Some label';
  const labelId = 'def61080-111c-4aea-b72f-5b60e516e36c';

  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerLogStderrSpy: sinon.SinonSpy;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
    loggerLogStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_RETENTIONLABEL_ENSURE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if listId, listTitle or listUrl option is not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: webUrl, listId: listId, listTitle: listTitle, listItemId: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listItemId is not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: webUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listItemId is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listItemId: 'abc', listTitle: listTitle, name: labelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if name or id option is not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

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

  it('applies a retentionlabel based on listId and name without assetId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'${listId}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`) {
        return listDetailsMock;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`
        && JSON.stringify(opts.data) === '{"listUrl":"https://contoso.sharepoint.com/sites/project-x/list","complianceTagValue":"Some label","itemIds":[1]}') {
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

  it('applies a retentionlabel based on listId, id and with assetId (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'${listId}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`) {
        return listDetailsMock;
      }
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP.CompliancePolicy.SPPolicyStoreProxy.GetAvailableTagsForSite(siteUrl=@a1)?@a1='${formatting.encodeQueryParameter(webUrl)}'`) {
        return retentionLabelListMockResponse;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`
        && JSON.stringify(opts.data) === '{"listUrl":"https://contoso.sharepoint.com/sites/project-x/list","complianceTagValue":"Some label","itemIds":[1]}') {
        return;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists(guid'${listId}')/items(1)/ValidateUpdateListItem()`) {
        return {
          "value": [
            {
              "ErrorCode": 0,
              "ErrorMessage": null,
              "FieldName": "ComplianceAssetId",
              "FieldValue": "XYZ",
              "HasException": false,
              "ItemId": 1
            }
          ]
        };
      }

      throw 'Invalid request';
    });


    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        listId: listId,
        webUrl: webUrl,
        listItemId: 1,
        id: labelId,
        assetId: 'XYZ'
      }
    }));
  });

  it('applies a retentionlabel based on listTitle, id and assetId (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`) {
        return listDetailsMock;
      }
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP.CompliancePolicy.SPPolicyStoreProxy.GetAvailableTagsForSite(siteUrl=@a1)?@a1='${formatting.encodeQueryParameter(webUrl)}'`) {
        return retentionLabelListMockResponse;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`
        && JSON.stringify(opts.data) === '{"listUrl":"https://contoso.sharepoint.com/sites/project-x/list","complianceTagValue":"Some label","itemIds":[1]}') {
        return;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')/items(${1})/ValidateUpdateListItem()`) {
        return {
          "value": [
            {
              "ErrorCode": 0,
              "ErrorMessage": null,
              "FieldName": "ComplianceAssetId",
              "FieldValue": "XYZ",
              "HasException": false,
              "ItemId": 1
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        listTitle: listTitle,
        webUrl: webUrl,
        listItemId: 1,
        id: labelId,
        assetId: 'XYZ'
      }
    }));
  });

  it('applies a retentionlabel based on listUrl, id and assetId (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP.CompliancePolicy.SPPolicyStoreProxy.GetAvailableTagsForSite(siteUrl=@a1)?@a1='${formatting.encodeQueryParameter(webUrl)}'`) {
        return retentionLabelListMockResponse;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`
        && JSON.stringify(opts.data) === '{"listUrl":"https://contoso.sharepoint.com/sites/project-x/list","complianceTagValue":"Some label","itemIds":[1]}') {
        return;
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList(@listUrl)/items(${1})/ValidateUpdateListItem()?@listUrl='${formatting.encodeQueryParameter(listUrl)}'`) {
        return {
          "value": [
            {
              "ErrorCode": 0,
              "ErrorMessage": null,
              "FieldName": "ComplianceAssetId",
              "FieldValue": "XYZ",
              "HasException": false,
              "ItemId": 1
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        listUrl: listUrl,
        webUrl: webUrl,
        listItemId: 1,
        id: labelId,
        assetId: 'XYZ'
      }
    }));
  });

  it('throws an error when a retentionlabel cannot be found on the site', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`
        && JSON.stringify(opts.data) === '{"listUrl":"https://contoso.sharepoint.com/sites/project-x/list","complianceTagValue":"Some non-existing label","itemIds":[1]}') {
        throw {
          error: {
            "odata.error": {
              "code": "-1, System.NotSupportedException",
              "message": {
                "lang": "nl-NL",
                "value": "Cannot find retention label with name: Some non-existing label"
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: false,
        listUrl: listUrl,
        webUrl: webUrl,
        listItemId: 1,
        name: 'Some non-existing label'
      }
    }), new CommandError('Cannot find retention label with name: Some non-existing label'));
  });

  it('throws an error when a retentionlabel cannot be found by id on the site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP.CompliancePolicy.SPPolicyStoreProxy.GetAvailableTagsForSite(siteUrl=@a1)?@a1='${formatting.encodeQueryParameter(webUrl)}'`) {
        return {
          "odata.metadata": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Collection(SP.CompliancePolicy.ComplianceTag)",
          "value": []
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: false,
        listUrl: listUrl,
        webUrl: webUrl,
        listItemId: 1,
        id: 'invalid'
      }
    }), new CommandError('The specified retention label does not exist or is not published to this SharePoint site. Use the name of the label if you want to apply an unpublished label.'));
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