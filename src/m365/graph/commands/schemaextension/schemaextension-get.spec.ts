import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./schemaextension-get');

describe(commands.SCHEMAEXTENSION_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.SCHEMAEXTENSION_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });
  it('gets schema extension', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions/$entity",
          "id": "adatumisv_exo2",
          "description": "sample description",
          "targetTypes": [
            "Message"
          ],
          "status": "Available",
          "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
          "properties": [
            {
              "name": "p1",
              "type": "String"
            },
            {
              "name": "p2",
              "type": "String"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });
    command.action(logger, {
      options: {
        debug: false,
        id: 'adatumisv_exo2'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions/$entity",
          "id": "adatumisv_exo2",
          "description": "sample description",
          "targetTypes": [
            "Message"
          ],
          "status": "Available",
          "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
          "properties": [
            {
              "name": "p1",
              "type": "String"
            },
            {
              "name": "p2",
              "type": "String"
            }
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
      }
    });
  });
  it('gets schema extension(debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions/$entity",
          "id": "adatumisv_exo2",
          "description": "sample description",
          "targetTypes": [
            "Message"
          ],
          "status": "Available",
          "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
          "properties": [
            {
              "name": "p1",
              "type": "String"
            },
            {
              "name": "p2",
              "type": "String"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });
    command.action(logger, {
      options: {
        debug: true,
        id: 'adatumisv_exo2'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions/$entity",
          "id": "adatumisv_exo2",
          "description": "sample description",
          "targetTypes": [
            "Message"
          ],
          "status": "Available",
          "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
          "properties": [
            {
              "name": "p1",
              "type": "String"
            },
            {
              "name": "p2",
              "type": "String"
            }
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('handles error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`) > -1) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });
    command.action(logger, {
      options: {
        debug: true,
        id: 'adatumisv_exo2'
      }
    } as any, (err?: any) => {
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
