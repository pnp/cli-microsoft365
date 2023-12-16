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
import command from './people-profilecardproperty-add.js';

describe(commands.PEOPLE_PROFILECARDPROPERTY_ADD, () => {

  //#region Mocked Responses
  const propertyResponse = {
    "directoryPropertyName": "userPrincipalName",
    "annotations": []
  };

  const customAttributePropertyResponse = {
    "directoryPropertyName": "customAttribute1",
    "annotations": [
      {
        "displayName": "Cost center",
        "localizations": [
          {
            "languageTag": "nl-NL",
            "displayName": "Kostenplaats"
          }
        ]
      }
    ]
  };

  const customAttributePropertyTextResponse = {
    "directoryPropertyName": "customAttribute1",
    "displayName": "Cost center",
    "displayName nl-NL": "Kostenplaats"
  };
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PEOPLE_PROFILECARDPROPERTY_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the name is not a valid value.', async () => {
    const actual = await command.validate({ options: { name: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the name is customAttribute1 and the displayName option is not used.', async () => {
    const actual = await command.validate({ options: { name: 'customAttribute1' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if a localization property has an invalid name.', async () => {
    const actual = await command.validate({ options: { name: 'customAttribute1', displayName: 'Cost center', 'invalid-nl-NL': 'Kostenplaats' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if a the localization option is used for a non-extension attribute.', async () => {
    const actual = await command.validate({ options: { name: 'userPrincipalName', 'displayName-nl-NL': 'Kostenplaats' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the displayName option is used for a non-extension attribute.', async () => {
    const actual = await command.validate({ options: { name: 'userPrincipalName', displayName: 'Cost center' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the name is set to userPrincipalName.', async () => {
    const actual = await command.validate({ options: { name: 'userPrincipalName' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if the name is customAttribute1 and the displayName option is used.', async () => {
    const actual = await command.validate({ options: { name: 'customAttribute1', displayName: 'Cost center' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if a correct localization option is used.', async () => {
    const actual = await command.validate({ options: { name: 'customAttribute1', displayName: 'Cost center', 'displayName-nl-NL': 'Kostenplaats' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly adds profile card property for userPrincipalName', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties`) {
        return propertyResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { name: 'userPrincipalName' } });
    assert(loggerLogSpy.calledOnceWithExactly(propertyResponse));
  });

  it('correctly adds profile card property for userPrincipalName (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties`) {
        return propertyResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { name: 'userPrincipalName', debug: true } });
    assert(loggerLogSpy.calledOnceWithExactly(propertyResponse));
  });

  it('correctly adds profile card property for fax', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties`) {
        return propertyResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { name: 'fax' } });
    assert(loggerLogSpy.calledOnceWithExactly(propertyResponse));
  });

  it('correctly adds profile card property for stateOrProvince', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties`) {
        return propertyResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { name: 'stateOrProvince' } });
    assert(loggerLogSpy.calledOnceWithExactly(propertyResponse));
  });

  it('correctly adds profile card property for alias (json output)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties`) {
        return propertyResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { name: 'alias', output: 'json' } });
    assert(loggerLogSpy.calledOnceWithExactly(propertyResponse));
  });

  it('correctly adds profile card property for alias (text output)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties`) {
        return propertyResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { name: 'alias', output: 'text' } });
    assert(loggerLogSpy.calledOnceWithExactly(propertyResponse));
  });

  it('correctly adds profile card property for an customAttribute', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties`) {
        return customAttributePropertyResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { name: 'customAttribute1', displayName: 'Cost center' } });
    assert(loggerLogSpy.calledOnceWithExactly(customAttributePropertyResponse));
  });

  it('correctly adds profile card property for an customAttribute with a localization', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties`) {
        return customAttributePropertyResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { name: 'customAttribute1', displayName: 'Cost center', 'displayName-nl-NL': 'Kostenplaats' } });
    assert(loggerLogSpy.calledOnceWithExactly(customAttributePropertyResponse));
  });

  it('correctly adds profile card property for an customAttribute with a localization (text output)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties`) {
        return customAttributePropertyResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, { options: { name: 'customAttribute1', displayName: 'Cost center', 'displayName-nl-NL': 'Kostenplaats', output: 'text' } });
    assert(loggerLogSpy.calledOnceWithExactly(customAttributePropertyTextResponse));
  });

  it('uses correct casing for name when incorrect casing is used', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties`) {
        return customAttributePropertyResponse;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { name: 'ALIAS', output: 'json' } });
    assert.strictEqual(postStub.lastCall.args[0].data.directoryPropertyName, 'Alias');
  });

  it('fails when the addition conflicts with an existing property', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties`) {
        throw {
          "error": {
            "code": "409",
            "message": "Conflicts with existing entry",
            "innerError": {
              "peopleAdminErrorCode": "PeopleAdminItemConflict",
              "peopleAdminRequestId": "36d1ea9e-83f8-49c9-7ebc-6f6c24ca03cc",
              "peopleAdminClientRequestId": "174cf6d3-6cde-46a8-b4f3-5d4d07354ac2",
              "date": "2023-11-02T15:22:36",
              "request-id": "174cf6d3-6cde-46a8-b4f3-5d4d07354ac2",
              "client-request-id": "174cf6d3-6cde-46a8-b4f3-5d4d07354ac2"
            }
          }
        };
      }

      throw `Invalid request ${opts.url}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: 'userPrincipalName'
      }
    }), new CommandError(`Conflicts with existing entry`));
  });
});