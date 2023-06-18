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
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./app-list');

describe(commands.APP_LIST, () => {
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
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
    assert.strictEqual(command.name, commands.APP_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'distributionMethod']);
  });

  it('fails validation if invalid distribution method specified', async () => {
    const actual = await command.validate({ options: { distributionMethod: 'invalid distribution method' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if valid distribution method specified', async () => {
    const actual = await command.validate({ options: { distributionMethod: 'store' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('lists Microsoft Teams apps in the organization app catalog', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=distributionMethod eq 'organization'`) {
        return {
          "value": [
            {
              "id": "7131a36d-bb5f-46b8-bb40-0b199a3fad74",
              "externalId": "4f0cd7c8-995e-4868-812d-d1d402a81eca",
              "displayName": "WsInfo",
              "distributionMethod": "organization"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { distributionMethod: 'organization' } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "7131a36d-bb5f-46b8-bb40-0b199a3fad74",
        "externalId": "4f0cd7c8-995e-4868-812d-d1d402a81eca",
        "displayName": "WsInfo",
        "distributionMethod": "organization"
      }
    ]));
  });

  it('lists Microsoft Teams apps in the organization app catalog and Microsoft Teams store', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps`) {
        return {
          "value": [
            {
              "id": "012be6ac-6f34-4ffa-9344-b857f7bc74e1",
              "externalId": null,
              "displayName": "Pickit Images",
              "distributionMethod": "store"
            },
            {
              "id": "01b22ab6-c657-491c-97a0-d745bea11269",
              "externalId": null,
              "displayName": "Hootsuite",
              "distributionMethod": "store"
            },
            {
              "id": "02d14659-a28b-4007-8544-b279c0d3628b",
              "externalId": null,
              "displayName": "Pivotal Tracker",
              "distributionMethod": "store"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { all: true, debug: true } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "012be6ac-6f34-4ffa-9344-b857f7bc74e1",
        "externalId": null,
        "displayName": "Pickit Images",
        "distributionMethod": "store"
      },
      {
        "id": "01b22ab6-c657-491c-97a0-d745bea11269",
        "externalId": null,
        "displayName": "Hootsuite",
        "distributionMethod": "store"
      },
      {
        "id": "02d14659-a28b-4007-8544-b279c0d3628b",
        "externalId": null,
        "displayName": "Pivotal Tracker",
        "distributionMethod": "store"
      }
    ]));
  });

  it('correctly handles error when retrieving apps', async () => {
    sinon.stub(request, 'get').rejects({
      "error": {
        "code": "ErrorOccured",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { output: 'json' } } as any), new CommandError('An error has occurred'));
  });
});
