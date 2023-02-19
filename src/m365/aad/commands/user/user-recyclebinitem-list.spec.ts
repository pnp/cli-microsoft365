import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { odata } from '../../../../utils/odata';
const command: Command = require('./user-recyclebinitem-list');

describe(commands.USER_RECYCLEBINITEM_LIST, () => {
  const deletedUsersResponse = [{ "businessPhones": [], "displayName": "John Doe", "givenName": "John Doe", "jobTitle": "Developer", "mail": "john@contoso.com", "mobilePhone": "0476345130", "officeLocation": "Washington", "preferredLanguage": "nl-BE", "surname": "John", "userPrincipalName": "7e06b56615f340138bf879874d52e68ajohn@contoso.com", "id": "7e06b566-15f3-4013-8bf8-79874d52e68a" }];
  const graphGetUrl = 'https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user';

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      odata.getAllItems
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_RECYCLEBINITEM_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'userPrincipalName']);
  });

  it('retrieves deleted users', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === graphGetUrl) {
        return deletedUsersResponse;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(deletedUsersResponse));
  });

  it('correctly handles API error', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async () => {
      throw {
        "error": {
          "code": "Invalid_Request",
          "message": "An error has occured while processing this request.",
          "innerError": {
            "request-id": "9b0df954-93b5-4de9-8b99-43c204a8aaf8",
            "date": "2018-04-24T18:56:48"
          }
        }
      };
    });

    await assert.rejects(command.action(logger, { options: { confirm: true } } as any),
      new CommandError(`An error has occured while processing this request.`));
  });
});