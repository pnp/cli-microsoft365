import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './people-profilecardproperty-list.js';

describe(commands.PEOPLE_PROFILECARDPROPERTY_LIST, () => {
  const profileCardPropertyName1 = 'customAttribute1';
  const profileCardPropertyName2 = 'customAttribute2';

  //#region Mocked responses
  const response =
  {
    value: [
      {
        directoryPropertyName: profileCardPropertyName1,
        annotations: [
          {
            displayName: 'Department',
            localizations: [
              {
                languageTag: 'de',
                displayName: 'Abteilung'
              },
              {
                languageTag: 'pl',
                displayName: 'Departament'
              }
            ]
          }
        ]
      },
      {
        directoryPropertyName: "Alias",
        annotations: []
      },
      {
        directoryPropertyName: profileCardPropertyName2,
        annotations: [
          {
            displayName: 'Cost center',
            localizations: [
              {
                languageTag: 'de',
                displayName: 'Kostenstelle'
              }
            ]
          }
        ]
      },
      {
        directoryPropertyName: profileCardPropertyName2,
        annotations: [
          {
            displayName: 'Cost center',
            localizations: [
              {
                languageTag: 'de',
                displayName: 'Kostenstelle'
              }
            ]
          }
        ]
      }
    ]
  };
  //#endregion

  let log: any[];
  let loggerLogSpy: sinon.SinonSpy;
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PEOPLE_PROFILECARDPROPERTY_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists profile card properties', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties`) {
        return response;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledOnceWith(response.value));
  });

  it('lists profile card properties information for other than json output', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/admin/people/profileCardProperties`) {
        return response;
      }

      throw 'Invalid Request';
    });

    const textOutput = [
      {
        directoryPropertyName: profileCardPropertyName1,
        displayName: response.value[0].annotations[0].displayName,
        ['displayName ' + response.value[0].annotations[0].localizations[0].languageTag]: response.value[0].annotations[0].localizations[0].displayName,
        ['displayName ' + response.value[0].annotations[0].localizations[1].languageTag]: response.value[0].annotations[0].localizations[1].displayName
      },
      {
        directoryPropertyName: profileCardPropertyName2,
        displayName: response.value[2].annotations[0].displayName,
        ['displayName ' + response.value[2].annotations[0].localizations[0].languageTag]: response.value[2].annotations[0].localizations[0].displayName
      },
      {
        directoryPropertyName: profileCardPropertyName2,
        displayName: response.value[3].annotations[0].displayName,
        ['displayName ' + response.value[3].annotations[0].localizations[0].languageTag]: response.value[3].annotations[0].localizations[0].displayName
      },
      {
        directoryPropertyName: "Alias"
      }
    ];

    await command.action(logger, { options: { output: 'text' } });
    assert(loggerLogSpy.calledOnceWith(textOutput));
  });

  it('handles unexpected API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects({
      error: {
        message: errorMessage
      }
    });

    await assert.rejects(command.action(logger, { options: { debug: true } }),
      new CommandError(errorMessage));
  });
});