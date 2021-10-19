import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./externalconnection-add');

describe(commands.EXTERNALCONNECTION_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const externalConnectionAddResponse: any = {
    "@odata.context": "https://graph.microsoft.com/beta/$metadata#connections/$entity",
    "id": "TestConnectionForCLI",
    "name": "Twitter Connector",
    "description": "Connector for showing key tweets",
    "connectorId": null,
    "state": null,
    "ingestedItemsCount": null,
    "searchSettings": null,
    "activitySettings": null,
    "complianceSettings": null,
    "configuration": {
      "authorizedApps": [
        "00000000-0000-0000-0000-000000000000"
      ],
      "authorizedAppIds": [
        "00000000-0000-0000-0000-000000000000"
      ]
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    (command as any).items = [];
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.EXTERNALCONNECTION_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds an external connection', (done) => {
      const options: any = {
        debug: false,
        externalConnectionId: 'TestConnectionForCLI',
        externalConnectionName: 'Test Connection for CLI',
        externalConnectionDescription: 'Test connection that will not do anything',
      };
  
      command.action(logger, { options: options } as any, () => {
        try {
          assert(loggerLogSpy.calledWith(externalConnectionAddResponse));
          done();
        }
        catch (e) {
          done(e);
        }
      });
    });


  it('correctly handles error', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "code": "Error",
          "message": "An error has occurred",
          "innerError": {
            "request-id": "9b0df954-93b5-4de9-8b99-43c204a8aaf8",
            "date": "2018-04-24T18:56:48"
          }
        }
      });
    });

    command.action(logger, { options: { debug: false, subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`An error has occurred`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if externalConnectionId is not set', () => {
    const actual = command.validate({ options: { subject: 'Lorem ipsum', to: 'mail@domain.com', bodyContents: 'Lorem ipsum', bodyContentType: 'Invalid' } });
    assert.notStrictEqual(actual, true);
  });

});