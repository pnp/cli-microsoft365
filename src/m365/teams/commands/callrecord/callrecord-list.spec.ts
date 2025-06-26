import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { z } from 'zod';
import commands from '../../commands.js';
import command from './callrecord-list.js';
import { odata } from '../../../../utils/odata.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { CommandError } from '../../../../Command.js';

describe(commands.CALLRECORD_LIST, () => {
  const validStartDateTime = new Date(Date.now() - 5 * 24 * 60 * 60 * 1000).toISOString();
  const validEndDateTime = new Date().toISOString();
  const validUserId = '4c3cd651-9c89-4d16-b578-28d425ea5eed';
  const validUserName = 'john.doe@contoso.com';

  const response = [
    {
      id: '145ae53b-7781-47c4-a5f7-b0e043012624',
      version: 1,
      type: 'peerToPeer',
      modalities: [
        'audio'
      ],
      lastModifiedDateTime: '2025-05-29T17:41:33.5066667Z',
      startDateTime: '2025-05-29T17:27:20.428943Z',
      endDateTime: '2025-05-29T17:27:20.428943Z',
      joinWebUrl: '',
      organizer: {
        acsUser: null,
        spoolUser: null,
        phone: null,
        guest: null,
        encrypted: null,
        onPremises: null,
        acsApplicationInstance: null,
        spoolApplicationInstance: null,
        applicationInstance: null,
        application: null,
        device: null,
        user: {
          id: '42559007-03c6-42c8-971f-cb79fd381a5a',
          displayName: 'John Doe',
          tenantId: '9d66187e-13f0-4666-9bac-be67ddd4b676'
        }
      },
      participants: [],
      'organizer_v2': {
        id: '42559007-03c6-42c8-971f-cb79fd381a5a',
        identity: {
          endpointType: null,
          acsUser: null,
          spoolUser: null,
          phone: null,
          guest: null,
          encrypted: null,
          onPremises: null,
          acsApplicationInstance: null,
          spoolApplicationInstance: null,
          applicationInstance: null,
          application: null,
          device: null,
          azureCommunicationServicesUser: null,
          assertedIdentity: null,
          user: {
            id: '42559007-03c6-42c8-971f-cb79fd381a5a',
            displayName: 'John Doe',
            tenantId: '9d66187e-13f0-4666-9bac-be67ddd4b676',
            userPrincipalName: 'john.doe@contoso.com'
          }
        },
        administrativeUnitInfos: []
      }
    },
    {
      id: '7fb28012-9b8a-4c06-b9ba-298a6a5aad89',
      version: 2,
      type: 'groupCall',
      modalities: [
        'audio'
      ],
      lastModifiedDateTime: '2025-05-29T18:02:52.84Z',
      startDateTime: '2025-05-29T17:27:51.6528954Z',
      endDateTime: '2025-05-29T17:30:34.2244066Z',
      joinWebUrl: '',
      organizer: {
        acsUser: null,
        spoolUser: null,
        phone: null,
        guest: null,
        encrypted: null,
        onPremises: null,
        acsApplicationInstance: null,
        spoolApplicationInstance: null,
        applicationInstance: null,
        application: null,
        device: null,
        user: {
          id: '42559007-03c6-42c8-971f-cb79fd381a5a',
          displayName: 'John Doe',
          tenantId: '9d66187e-13f0-4666-9bac-be67ddd4b676'
        }
      },
      participants: [],
      'organizer_v2': {
        id: '42559007-03c6-42c8-971f-cb79fd381a5a',
        identity: {
          endpointType: null,
          acsUser: null,
          spoolUser: null,
          phone: null,
          guest: null,
          encrypted: null,
          onPremises: null,
          acsApplicationInstance: null,
          spoolApplicationInstance: null,
          applicationInstance: null,
          application: null,
          device: null,
          azureCommunicationServicesUser: null,
          assertedIdentity: null,
          user: {
            id: '42559007-03c6-42c8-971f-cb79fd381a5a',
            displayName: 'John Doe',
            tenantId: '9d66187e-13f0-4666-9bac-be67ddd4b676',
            userPrincipalName: 'john.doe@contoso.com'
          }
        },
        administrativeUnitInfos: []
      }
    },
    {
      id: 'd8cd0a5d-81b0-47ac-9fe1-eaa7df9af57d',
      version: 2,
      type: 'groupCall',
      modalities: [
        'audio',
        'video'
      ],
      lastModifiedDateTime: '2025-05-29T18:16:52.5333333Z',
      startDateTime: '2025-05-29T17:37:38.3486338Z',
      endDateTime: '2025-05-29T17:45:18.3793478Z',
      joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/',
      organizer: {
        acsUser: null,
        spoolUser: null,
        phone: null,
        guest: null,
        encrypted: null,
        onPremises: null,
        acsApplicationInstance: null,
        spoolApplicationInstance: null,
        applicationInstance: null,
        application: null,
        device: null,
        user: {
          id: '42559007-03c6-42c8-971f-cb79fd381a5a',
          displayName: 'John Doe',
          tenantId: '9d66187e-13f0-4666-9bac-be67ddd4b676'
        }
      },
      participants: [],
      'organizer_v2': {
        id: '42559007-03c6-42c8-971f-cb79fd381a5a',
        identity: {
          endpointType: null,
          acsUser: null,
          spoolUser: null,
          phone: null,
          guest: null,
          encrypted: null,
          onPremises: null,
          acsApplicationInstance: null,
          spoolApplicationInstance: null,
          applicationInstance: null,
          application: null,
          device: null,
          azureCommunicationServicesUser: null,
          assertedIdentity: null,
          user: {
            id: '42559007-03c6-42c8-971f-cb79fd381a5a',
            displayName: 'John Doe',
            tenantId: '9d66187e-13f0-4666-9bac-be67ddd4b676',
            userPrincipalName: 'john.doe@contoso.com'
          }
        },
        administrativeUnitInfos: []
      }
    }
  ];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  let assertAccessTokenTypeStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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

    assertAccessTokenTypeStub = sinon.stub(accessToken, 'assertAccessTokenType').resolves();
  });

  afterEach(() => {
    sinonUtil.restore([
      odata.getAllItems,
      accessToken.assertAccessTokenType
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CALLRECORD_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct default properties', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'type', 'startDateTime', 'endDateTime']);
  });

  it('fails validation when userId is not a valid guid', async () => {
    const actual = commandOptionsSchema.safeParse({
      userId: 'invalid'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when userName is not a valid UPN', async () => {
    const actual = commandOptionsSchema.safeParse({
      userName: 'invalid'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when startDateTime is not a date', async () => {
    const actual = commandOptionsSchema.safeParse({
      startDateTime: 'invalid'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when startDateTime is more than 30 days ago', async () => {
    const invalidStartDateTime = new Date();
    invalidStartDateTime.setDate(invalidStartDateTime.getDate() - 31);

    const actual = commandOptionsSchema.safeParse({
      startDateTime: invalidStartDateTime.toISOString()
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when startDateTime is in the future', async () => {
    const invalidStartDateTime = new Date();
    invalidStartDateTime.setHours(invalidStartDateTime.getHours() + 1);

    const actual = commandOptionsSchema.safeParse({
      startDateTime: invalidStartDateTime.toISOString()
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when endDateTime is not a date', async () => {
    const actual = commandOptionsSchema.safeParse({
      endDateTime: 'invalid'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when endDateTime is more than 30 days ago', async () => {
    const invalidEndDateTime = new Date();
    invalidEndDateTime.setDate(invalidEndDateTime.getDate() - 31);

    const actual = commandOptionsSchema.safeParse({
      endDateTime: invalidEndDateTime.toISOString()
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when endDateTime is in the future', async () => {
    const invalidEndDateTime = new Date();
    invalidEndDateTime.setHours(invalidEndDateTime.getHours() + 1);

    const actual = commandOptionsSchema.safeParse({
      endDateTime: invalidEndDateTime.toISOString()
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when both userId and userName are specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      userId: validUserId,
      userName: validUserName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation when startDateTime is after endDateTime', async () => {
    const actual = commandOptionsSchema.safeParse({
      startDateTime: validStartDateTime,
      endDateTime: new Date(new Date(validStartDateTime).getTime() - 1000).toISOString()
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('passes validation when all parameters are valid', async () => {
    const actual = commandOptionsSchema.safeParse({
      userId: validUserId,
      startDateTime: validStartDateTime,
      endDateTime: validEndDateTime
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when only userName is specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      userName: validUserName
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when only startDateTime is specified', async () => {
    const actual = commandOptionsSchema.safeParse({
      startDateTime: validStartDateTime
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when no parameters are specified', async () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, true);
  });

  it('successfully outputs result', async () => {
    sinon.stub(odata, 'getAllItems').resolves(response);

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(response));
  });

  it('asserts that the command only runs with application permissions', async () => {
    sinon.stub(odata, 'getAllItems').resolves(response);

    await command.action(logger, { options: {} });
    assert(assertAccessTokenTypeStub.calledOnceWith('application'));
  });

  it('successfully gets all call records', async () => {
    const odataStub = sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === 'https://graph.microsoft.com/v1.0/communications/callRecords') {
        return response;
      }

      throw 'Invalid GET request: ' + url;
    });

    await command.action(logger, { options: { verbose: true } });
    assert(odataStub.calledOnce);
  });

  it('successfully gets all call records for a user by ID', async () => {
    const odataStub = sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === 'https://graph.microsoft.com/v1.0/communications/callRecords?$filter=participants_v2/any(p:p/id eq \'4c3cd651-9c89-4d16-b578-28d425ea5eed\')') {
        return response;
      }

      throw 'Invalid GET request: ' + url;
    });

    await command.action(logger, { options: { userId: validUserId } });
    assert(odataStub.calledOnce);
  });

  it('successfully gets all call records for a user by UPN', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').callsFake(async (upn) => {
      if (upn === validUserName) {
        return validUserId;
      }

      throw 'Invalid UPN: ' + upn;
    });

    const odataStub = sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === 'https://graph.microsoft.com/v1.0/communications/callRecords?$filter=participants_v2/any(p:p/id eq \'4c3cd651-9c89-4d16-b578-28d425ea5eed\')') {
        return response;
      }

      throw 'Invalid GET request: ' + url;
    });

    await command.action(logger, { options: { userName: validUserName } });
    assert(odataStub.calledOnce);
  });

  it('successfully get call records with start and end date', async () => {
    const odataStub = sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `https://graph.microsoft.com/v1.0/communications/callRecords?$filter=startDateTime ge ${validStartDateTime} and startDateTime lt ${validEndDateTime}`) {
        return response;
      }

      throw 'Invalid GET request: ' + url;
    });

    await command.action(logger, { options: { startDateTime: validStartDateTime, endDateTime: validEndDateTime } });
    assert(odataStub.calledOnce);
  });

  it('successfully get call records with start date', async () => {
    const odataStub = sinon.stub(odata, 'getAllItems').callsFake(async url => {
      if (url === `https://graph.microsoft.com/v1.0/communications/callRecords?$filter=startDateTime ge ${validStartDateTime}`) {
        return response;
      }

      throw 'Invalid GET request: ' + url;
    });

    await command.action(logger, { options: { startDateTime: validStartDateTime } });
    assert(odataStub.calledOnce);
  });

  it('correctly handles error when retrieving call records', async () => {
    const error = {
      error: {
        code: 'UnknownError',
        message: 'An unknown error has occurred.'
      }
    };
    sinon.stub(odata, 'getAllItems').rejects(error);

    await assert.rejects(command.action(logger, { options: {} }),
      new CommandError(`An unknown error has occurred.`));
  });
});