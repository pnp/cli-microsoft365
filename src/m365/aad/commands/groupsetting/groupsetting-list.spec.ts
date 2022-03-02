import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./groupsetting-list');

describe(commands.GROUPSETTING_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.GROUPSETTING_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName']);
  });

  it('lists group setting templates', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
        return Promise.resolve({
          "value": [
            {
              "id": "68498d53-e3e8-47fd-bf19-eff723d5707e",
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
                  "value": ""
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
                  "value": ""
                },
                {
                  "name": "ClassificationList",
                  "value": ""
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

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith([{
          "id": "68498d53-e3e8-47fd-bf19-eff723d5707e",
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
              "value": ""
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
              "value": ""
            },
            {
              "name": "ClassificationList",
              "value": ""
            },
            {
              "name": "EnableGroupCreation",
              "value": "true"
            }
          ]
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists group setting templates (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
        return Promise.resolve({
          "value": [
            {
              "id": "68498d53-e3e8-47fd-bf19-eff723d5707e",
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
                  "value": ""
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
                  "value": ""
                },
                {
                  "name": "ClassificationList",
                  "value": ""
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

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith([{
          "id": "68498d53-e3e8-47fd-bf19-eff723d5707e",
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
              "value": ""
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
              "value": ""
            },
            {
              "name": "ClassificationList",
              "value": ""
            },
            {
              "name": "EnableGroupCreation",
              "value": "true"
            }
          ]
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('includes all properties in output type json', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
        return Promise.resolve({
          "value": [
            {
              "id": "68498d53-e3e8-47fd-bf19-eff723d5707e",
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
                  "value": ""
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
                  "value": ""
                },
                {
                  "name": "ClassificationList",
                  "value": ""
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

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, output: 'json' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "68498d53-e3e8-47fd-bf19-eff723d5707e",
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
                "value": ""
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
                "value": ""
              },
              {
                "name": "ClassificationList",
                "value": ""
              },
              {
                "name": "EnableGroupCreation",
                "value": "true"
              }
            ]
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groupSettings`) {
        return Promise.reject({
          error: {
            "error": {
              "code": "Request_ResourceNotFound",
              "message": "An error has occurred",
              "innerError": {
                "request-id": "7e192558-7438-46db-a4c9-5dca83d2ec96",
                "date": "2018-02-21T20:38:50"
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});