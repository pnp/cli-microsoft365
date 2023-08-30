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
import spoListItemRetentionLabelEnsureCommand from '../listitem/listitem-retentionlabel-ensure.js';
import command from './file-retentionlabel-ensure.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.FILE_RETENTIONLABEL_ENSURE, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const fileUrl = `/Shared Documents/Fo'lde'r/Document.docx`;
  const fileId = 'b2307a39-e878-458b-bc90-03bc578531d6';
  const listId = 1;
  const retentionlabelName = "retentionlabel";
  const SpoListItemRetentionLabelEnsureCommandOutput = `{ "stdout": "", "stderr": "" }`;
  const fileResponse = {
    ListItemAllFields: {
      Id: listId,
      ParentList: {
        Id: '75c4d697-bbff-40b8-a740-bf9b9294e5aa'
      }
    }
  };

  const retentionLabelResponse = {
    value: [{
      AcceptMessagesOnlyFromSendersOrMembers: false,
      AccessType: null,
      AllowAccessFromUnmanagedDevice: null,
      AutoDelete: true,
      BlockDelete: true,
      BlockEdit: false,
      ComplianceFlags: 1,
      ContainsSiteLabel: false,
      DisplayName: '',
      EncryptionRMSTemplateId: null,
      HasRetentionAction: true,
      IsEventTag: true,
      MultiStageReviewerEmail: null,
      NextStageComplianceTag: null,
      Notes: null,
      RequireSenderAuthenticationEnabled: false,
      ReviewerEmail: null,
      SharingCapabilities: null,
      SuperLock: false,
      TagDuration: 2555,
      TagId: 'f6e20c71-7d56-414d-bb98-8ee927a308bd',
      TagName: retentionlabelName,
      TagRetentionBasedOn: 'Retention Label Parent',
      UnlockedAsDefault: false
    }]
  };

  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      Cli.executeCommandWithOutput,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_RETENTIONLABEL_ENSURE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds the retentionlabel from a file based on fileUrl', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')?$expand=ListItemAllFields,ListItemAllFields/ParentList&$select=ServerRelativeUrl,ListItemAllFields/ParentList/Id,ListItemAllFields/Id`) {
        return fileResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoListItemRetentionLabelEnsureCommand) {
        return ({
          stdout: SpoListItemRetentionLabelEnsureCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        fileUrl: fileUrl,
        webUrl: webUrl,
        name: retentionlabelName
      }
    }));
  });

  it('adds the retentionlabel from a file based on fileId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFileById('${fileId}')?$expand=ListItemAllFields,ListItemAllFields/ParentList&$select=ServerRelativeUrl,ListItemAllFields/ParentList/Id,ListItemAllFields/Id`) {
        return fileResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoListItemRetentionLabelEnsureCommand) {
        return ({
          stdout: SpoListItemRetentionLabelEnsureCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        debug: true,
        fileId: fileId,
        webUrl: webUrl,
        name: retentionlabelName
      }
    }));
  });

  it('adds the event based retentionlabel from a file based on fileUrl with assetId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')?$expand=ListItemAllFields,ListItemAllFields/ParentList&$select=ServerRelativeUrl,ListItemAllFields/ParentList/Id,ListItemAllFields/Id`) {
        return fileResponse;
      }

      if (opts.url === `https://contoso.sharepoint.com/_api/SP.CompliancePolicy.SPPolicyStoreProxy.GetAvailableTagsForSite(siteUrl=@a1)?@a1='${formatting.encodeQueryParameter(webUrl)}'`) {
        return retentionLabelResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoListItemRetentionLabelEnsureCommand) {
        return ({
          stdout: SpoListItemRetentionLabelEnsureCommandOutput
        });
      }

      throw new CommandError('Unknown case');
    });


    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/lists(guid'${fileResponse.ListItemAllFields.ParentList.Id}')/items(${listId})/ValidateUpdateListItem()`) {
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
        verbose: true,
        fileUrl: fileUrl,
        webUrl: webUrl,
        name: retentionlabelName,
        assetId: 'XYZ'
      }
    }));
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Something went wrong';

    sinon.stub(request, 'get').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: retentionlabelName,
        fileUrl: fileUrl,
        webUrl: webUrl
      }
    }), new CommandError(errorMessage));
  });

  it('fails validation if both fileUrl or fileId options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: webUrl, name: retentionlabelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileUrl: fileUrl, name: retentionlabelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileUrl: fileUrl, name: retentionlabelName } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: webUrl, fileId: '12345', name: retentionlabelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the fileId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, name: retentionlabelName } }, commandInfo);
    assert(actual);
  });

  it('fails validation if both fileId and fileUrl options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: webUrl, fileId: fileId, fileUrl: fileUrl, name: retentionlabelName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});