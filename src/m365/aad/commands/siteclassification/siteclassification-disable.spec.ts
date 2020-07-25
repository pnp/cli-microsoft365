import commands from '../../commands';
import Command, { CommandError, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./siteclassification-disable');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';

describe(commands.SITECLASSIFICATION_DISABLE, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      },
      prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
        promptOptions = options;
        cb({ continue: false });
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    promptOptions = undefined;
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.delete
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });


  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITECLASSIFICATION_DISABLE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('prompts before disabling siteclassification when confirm option not passed', (done) => {
    cmdInstance.action({ options: { debug: false } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles Microsoft 365 Tenant siteclassification is not enabled', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings`) {
        return Promise.resolve({
          value: [
          ]
        })
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({ options: { debug: true, confirm: true } }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Site classification is not enabled.')));
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('handles Microsoft 365 Tenant siteclassification missing DirectorySettingTemplate', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings`) {
        return Promise.resolve({
          value: [
            {
              "id": "d20c475c-6f96-449a-aee8-08146be187d3",
              "displayName": "Group.Unified_not_exist",
              "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
              "values": [
                {
                  "name": "CustomBlockedWordsList",
                  "value": ""
                },
                {
                  "name": "EnableMSStandardBlockedWords",
                  "value": "false"
                },
                {
                  "name": "ClassificationDescriptions",
                  "value": ""
                },
                {
                  "name": "DefaultClassification",
                  "value": "TopSecret"
                },
                {
                  "name": "PrefixSuffixNamingRequirement",
                  "value": ""
                },
                {
                  "name": "AllowGuestsToBeGroupOwner",
                  "value": "false"
                },
                {
                  "name": "AllowGuestsToAccessGroups",
                  "value": "true"
                },
                {
                  "name": "GuestUsageGuidelinesUrl",
                  "value": ""
                },
                {
                  "name": "GroupCreationAllowedGroupId",
                  "value": ""
                },
                {
                  "name": "AllowToAddGuests",
                  "value": "true"
                },
                {
                  "name": "UsageGuidelinesUrl",
                  "value": "https://test"
                },
                {
                  "name": "ClassificationList",
                  "value": "TopSecret"
                },
                {
                  "name": "EnableGroupCreation",
                  "value": "true"
                }
              ]
            }
          ]
        });
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({ options: { debug: true, confirm: true } }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Missing DirectorySettingTemplate for \"Group.Unified\"")));
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('handles Microsoft 365 Tenant siteclassification missing UnifiedGroupSetting ID', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings`) {
        return Promise.resolve({
          value: [
            {
              "id_doesnotexists": "d20c475c-6f96-449a-aee8-08146be187d3",
              "displayName": "Group.Unified",
              "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
              "values": [
                {
                  "name": "CustomBlockedWordsList",
                  "value": ""
                },
                {
                  "name": "EnableMSStandardBlockedWords",
                  "value": "false"
                },
                {
                  "name": "ClassificationDescriptions",
                  "value": ""
                },
                {
                  "name": "DefaultClassification",
                  "value": "TopSecret"
                },
                {
                  "name": "PrefixSuffixNamingRequirement",
                  "value": ""
                },
                {
                  "name": "AllowGuestsToBeGroupOwner",
                  "value": "false"
                },
                {
                  "name": "AllowGuestsToAccessGroups",
                  "value": "true"
                },
                {
                  "name": "GuestUsageGuidelinesUrl",
                  "value": ""
                },
                {
                  "name": "GroupCreationAllowedGroupId",
                  "value": ""
                },
                {
                  "name": "AllowToAddGuests",
                  "value": "true"
                },
                {
                  "name": "UsageGuidelinesUrl",
                  "value": "https://test"
                },
                {
                  "name": "ClassificationList",
                  "value": "TopSecret"
                },
                {
                  "name": "EnableGroupCreation",
                  "value": "true"
                }
              ]
            }
          ]
        });
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({ options: { debug: true, confirm: true } }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Missing UnifiedGroupSettting id")));
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('handles Microsoft 365 Tenant siteclassification empty UnifiedGroupSetting ID', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings`) {
        return Promise.resolve({
          value: [
            {
              "id": "",
              "displayName": "Group.Unified",
              "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
              "values": [
                {
                  "name": "CustomBlockedWordsList",
                  "value": ""
                },
                {
                  "name": "EnableMSStandardBlockedWords",
                  "value": "false"
                },
                {
                  "name": "ClassificationDescriptions",
                  "value": ""
                },
                {
                  "name": "DefaultClassification",
                  "value": "TopSecret"
                },
                {
                  "name": "PrefixSuffixNamingRequirement",
                  "value": ""
                },
                {
                  "name": "AllowGuestsToBeGroupOwner",
                  "value": "false"
                },
                {
                  "name": "AllowGuestsToAccessGroups",
                  "value": "true"
                },
                {
                  "name": "GuestUsageGuidelinesUrl",
                  "value": ""
                },
                {
                  "name": "GroupCreationAllowedGroupId",
                  "value": ""
                },
                {
                  "name": "AllowToAddGuests",
                  "value": "true"
                },
                {
                  "name": "UsageGuidelinesUrl",
                  "value": "https://test"
                },
                {
                  "name": "ClassificationList",
                  "value": "TopSecret"
                },
                {
                  "name": "EnableGroupCreation",
                  "value": "true"
                }
              ]
            }
          ]
        });
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({ options: { debug: true, confirm: true } }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Missing UnifiedGroupSettting id")));
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('handles disabling site classification without prompting', (done) => {
    let deleteRequestIssued = false;
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings`) {
        return Promise.resolve({
          value: [
            {
              "id": "d20c475c-6f96-449a-aee8-08146be187d3",
              "displayName": "Group.Unified",
              "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
              "values": [
                {
                  "name": "CustomBlockedWordsList",
                  "value": ""
                },
                {
                  "name": "EnableMSStandardBlockedWords",
                  "value": "false"
                },
                {
                  "name": "ClassificationDescriptions",
                  "value": ""
                },
                {
                  "name": "DefaultClassification",
                  "value": "TopSecret"
                },
                {
                  "name": "PrefixSuffixNamingRequirement",
                  "value": ""
                },
                {
                  "name": "AllowGuestsToBeGroupOwner",
                  "value": "false"
                },
                {
                  "name": "AllowGuestsToAccessGroups",
                  "value": "true"
                },
                {
                  "name": "GuestUsageGuidelinesUrl",
                  "value": ""
                },
                {
                  "name": "GroupCreationAllowedGroupId",
                  "value": ""
                },
                {
                  "name": "AllowToAddGuests",
                  "value": "true"
                },
                {
                  "name": "UsageGuidelinesUrl",
                  "value": "https://test"
                },
                {
                  "name": "ClassificationList",
                  "value": "TopSecret"
                },
                {
                  "name": "EnableGroupCreation",
                  "value": "true"
                }
              ]
            }
          ]
        });
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings/d20c475c-6f96-449a-aee8-08146be187d3`) {
        deleteRequestIssued = true;

        return Promise.resolve({
          value: [
          ]

        });
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({ options: { confirm: true } }, (err: any) => {
      try {
        assert(deleteRequestIssued);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('handles disabling site classification without prompting (debug)', (done) => {
    let deleteRequestIssued = false;
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings`) {
        return Promise.resolve({
          value: [
            {
              "id": "d20c475c-6f96-449a-aee8-08146be187d3",
              "displayName": "Group.Unified",
              "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
              "values": [
                {
                  "name": "CustomBlockedWordsList",
                  "value": ""
                },
                {
                  "name": "EnableMSStandardBlockedWords",
                  "value": "false"
                },
                {
                  "name": "ClassificationDescriptions",
                  "value": ""
                },
                {
                  "name": "DefaultClassification",
                  "value": "TopSecret"
                },
                {
                  "name": "PrefixSuffixNamingRequirement",
                  "value": ""
                },
                {
                  "name": "AllowGuestsToBeGroupOwner",
                  "value": "false"
                },
                {
                  "name": "AllowGuestsToAccessGroups",
                  "value": "true"
                },
                {
                  "name": "GuestUsageGuidelinesUrl",
                  "value": ""
                },
                {
                  "name": "GroupCreationAllowedGroupId",
                  "value": ""
                },
                {
                  "name": "AllowToAddGuests",
                  "value": "true"
                },
                {
                  "name": "UsageGuidelinesUrl",
                  "value": "https://test"
                },
                {
                  "name": "ClassificationList",
                  "value": "TopSecret"
                },
                {
                  "name": "EnableGroupCreation",
                  "value": "true"
                }
              ]
            }
          ]
        });
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings/d20c475c-6f96-449a-aee8-08146be187d3`) {
        deleteRequestIssued = true;

        return Promise.resolve({
          value: [
          ]
        });
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({ options: { debug: true, confirm: true } }, (err: any) => {
      try {
        assert(deleteRequestIssued && cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('aborts removing the group when prompt not confirmed', (done) => {
    const postSpy = sinon.spy(request, 'delete');
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(postSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles disabling site classification when prompt confirmed', (done) => {
    let deleteRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings`) {
        return Promise.resolve({
          value: [
            {
              "id": "d20c475c-6f96-449a-aee8-08146be187d3",
              "displayName": "Group.Unified",
              "templateId": "62375ab9-6b52-47ed-826b-58e47e0e304b",
              "values": [
                {
                  "name": "CustomBlockedWordsList",
                  "value": ""
                },
                {
                  "name": "EnableMSStandardBlockedWords",
                  "value": "false"
                },
                {
                  "name": "ClassificationDescriptions",
                  "value": ""
                },
                {
                  "name": "DefaultClassification",
                  "value": "TopSecret"
                },
                {
                  "name": "PrefixSuffixNamingRequirement",
                  "value": ""
                },
                {
                  "name": "AllowGuestsToBeGroupOwner",
                  "value": "false"
                },
                {
                  "name": "AllowGuestsToAccessGroups",
                  "value": "true"
                },
                {
                  "name": "GuestUsageGuidelinesUrl",
                  "value": "https://test"
                },
                {
                  "name": "GroupCreationAllowedGroupId",
                  "value": ""
                },
                {
                  "name": "AllowToAddGuests",
                  "value": "true"
                },
                {
                  "name": "UsageGuidelinesUrl",
                  "value": "https://test"
                },
                {
                  "name": "ClassificationList",
                  "value": "TopSecret"
                },
                {
                  "name": "EnableGroupCreation",
                  "value": "true"
                }
              ]
            }
          ]
        });
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/settings/d20c475c-6f96-449a-aee8-08146be187d3`) {
        deleteRequestIssued = true;

        return Promise.resolve({
          value: [
          ]
        });
      }

      return Promise.reject('Invalid Request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(deleteRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});