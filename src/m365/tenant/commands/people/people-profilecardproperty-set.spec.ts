import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './people-profilecardproperty-set.js';

describe(commands.PEOPLE_PROFILECARDPROPERTY_SET, () => {
  const profileCardPropertyName = 'customAttribute1';
  const displayName = 'Cost center';
  const dutchTranslation = 'Kostencentrum';
  const germanTranslation = 'Kostenstelle';

  //#region Mocked responses
  const response = {
    directoryPropertyName: profileCardPropertyName,
    annotations: [
      {
        displayName: displayName,
        localizations: []
      }
    ]
  };

  const responseWithTranslations = {
    directoryPropertyName: profileCardPropertyName,
    annotations: [
      {
        displayName: displayName,
        localizations: [
          {
            languageTag: 'nl-NL',
            displayName: dutchTranslation
          },
          {
            languageTag: 'de',
            displayName: germanTranslation
          }
        ]
      }
    ]
  };
  //#endregion

  let log: any[];
  let loggerLogSpy: sinon.SinonSpy;
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PEOPLE_PROFILECARDPROPERTY_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when name is invalid', async () => {
    const actual = await command.validate({ options: { name: 'invalid', displayName: displayName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid unknown option is passed', async () => {
    const actual = await command.validate({ options: { name: profileCardPropertyName, displayName: displayName, 'nl-NL': dutchTranslation } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name and displayName is valid', async () => {
    const actual = await command.validate({ options: { name: profileCardPropertyName, displayName: displayName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when name is valid with different capitalization', async () => {
    const actual = await command.validate({ options: { name: 'cUstoMATTriBUtE1', displayName: displayName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid unknown option is passed', async () => {
    const actual = await command.validate({ options: { name: profileCardPropertyName, displayName: displayName, 'displayName-nl-NL': dutchTranslation } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('logs an output when command runs successfully', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/${profileCardPropertyName}`) {
        return response;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { name: profileCardPropertyName, displayName: displayName, verbose: true } });
    assert(loggerLogSpy.calledOnceWith(response));
  });

  it('updates profile card property information without translations correctly', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/${profileCardPropertyName}`) {
        return response;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { name: profileCardPropertyName, displayName: displayName } });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      annotations: [
        {
          displayName: displayName,
          localizations: []
        }
      ]
    });
  });

  it('updates profile card property information with translations correctly', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/${profileCardPropertyName}`) {
        return responseWithTranslations;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { name: profileCardPropertyName, displayName: displayName, 'displayName-nl-NL': dutchTranslation, 'displayName-de': germanTranslation } });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      annotations: [
        {
          displayName: displayName,
          localizations: [
            {
              languageTag: 'nl-NL',
              displayName: dutchTranslation
            },
            {
              languageTag: 'de',
              displayName: germanTranslation
            }
          ]
        }
      ]
    });
  });

  it('logs an output correctly when updating profile card property information with translations with text output', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/${profileCardPropertyName}`) {
        return responseWithTranslations;
      }

      throw 'Invalid Request';
    });

    const textOutput = {
      directoryPropertyName: profileCardPropertyName,
      displayName: displayName,
      'displayName nl-NL': dutchTranslation,
      'displayName de': germanTranslation
    };

    await command.action(logger, { options: { name: profileCardPropertyName, displayName: displayName, 'displayName-nl-NL': dutchTranslation, 'displayName-de': germanTranslation, output: 'text' } });
    assert(loggerLogSpy.calledOnceWith(textOutput));
  });

  it('uses correct casing for name when incorrect casing is used', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties/${profileCardPropertyName}`) {
        return response;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { name: profileCardPropertyName.toUpperCase(), displayName: displayName } });
    assert(patchStub.called);
  });

  it('handles unexpected API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'patch').rejects({
      error: {
        message: errorMessage
      }
    });

    await assert.rejects(command.action(logger, { options: { name: profileCardPropertyName } } as any),
      new CommandError(errorMessage));
  });
});