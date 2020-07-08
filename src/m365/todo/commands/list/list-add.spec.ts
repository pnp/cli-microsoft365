import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./list-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LIST_ADD, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post,
      Date.now
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
    assert.equal(command.name.startsWith(commands.LIST_ADD), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('adds To Do task list', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/me/todo/lists`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#lists/$entity",
          "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/ZGlTQ==\"",
          "displayName": "FooList",
          "isOwner": true,
          "isShared": false,
          "wellknownListName": "none",
          "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIgAAA="
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        name: "FooList"
      }
    }, () => {
      try {
        assert.equal(JSON.stringify(log[0]), JSON.stringify({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#lists/$entity",
          "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/ZGlTQ==\"",
          "displayName": "FooList",
          "isOwner": true,
          "isShared": false,
          "wellknownListName": "none",
          "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIgAAA="
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds To Do task list (verbose)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/me/todo/lists`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#lists/$entity",
          "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/ZGlTQ==\"",
          "displayName": "FooList",
          "isOwner": true,
          "isShared": false,
          "wellknownListName": "none",
          "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIgAAA="
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        verbose: true,
        name: "FooList"
      }
    }, () => {
      try {
        const expected = vorpal.chalk.green('DONE');
        assert.equal(log.filter(l => l == expected).length, 1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error correctly', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        name: "FooList"
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

  it('fails validation if name is not set', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        name: null
      }
    });
    assert.notEqual(actual, true);
  });

  it('passes validation when all parameters are valid', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        name: 'Foo'
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
    assert(find.calledWith(commands.LIST_ADD));
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