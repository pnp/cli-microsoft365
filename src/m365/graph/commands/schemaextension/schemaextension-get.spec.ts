import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError} from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./schemaextension-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.SCHEMAEXTENSION_GET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
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
      vorpal.find,
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
    assert.equal(command.name.startsWith(commands.SCHEMAEXTENSION_GET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
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
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

it('fails validation if the id is not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: null
      }
    });
    assert.notEqual(actual, true);
  });

  it('passes validation if the id is specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        id: 'adatumisv_exo2'
      }
    });
    assert.equal(actual, true);
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.SCHEMAEXTENSION_GET));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
});
