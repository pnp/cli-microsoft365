import commands from '../../commands';
import Command, { CommandOption, CommandError} from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./schemaextension-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.SCHEMAEXTENSION_GET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

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
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
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
      if ((opts.url as string).indexOf(`schemaExtensions`)> -1) {
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
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        id: 'adatumisv_exo2',
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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
      if ((opts.url as string).indexOf(`schemaExtensions`)> -1) {
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
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        id: 'adatumisv_exo2',
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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
      if ((opts.url as string).indexOf(`schemaExtensions`)> -1) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        id: 'adatumisv_exo2',
      }
    }, (err?: any) => {
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
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
