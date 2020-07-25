import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./page-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';

describe(commands.PAGE_REMOVE, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let promptOptions: any;

  const fakeRestCalls: (pageName?: string) => sinon.SinonStub = (pageName: string = 'page.aspx') => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/${pageName}')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon
      .stub(command as any, 'getRequestDigest')
      .callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
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
      },
      prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
        promptOptions = options;
        cb({ continue: false });
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([request.post]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PAGE_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes a modern page without confirm prompt', (done) => {
    fakeRestCalls();
    cmdInstance.action(
      {
        options: {
          debug: false,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          confirm: true
        }
      },
      () => {
        try {
          assert(cmdInstanceLogSpy.notCalled);
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('removes a modern page (debug) without confirm prompt', (done) => {
    fakeRestCalls();
    cmdInstance.action(
      {
        options: {
          debug: true,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          confirm: true
        }
      },
      () => {
        try {
          assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('removes a modern page (debug) without confirm prompt on root of tenant', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/getfilebyserverrelativeurl('/sitepages/page.aspx')`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: true,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com',
          confirm: true
        }
      },
      () => {
        try {
          assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('removes a modern page with confirm prompt', (done) => {
    fakeRestCalls();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: true });
    };
    cmdInstance.action(
      {
        options: {
          debug: false,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a'
        }
      },
      () => {
        try {
          assert(cmdInstanceLogSpy.notCalled);
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('removes a modern page (debug) with confirm prompt', (done) => {
    fakeRestCalls();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: true });
    };
    cmdInstance.action(
      {
        options: {
          debug: true,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a'
        }
      },
      () => {
        try {
          assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('should prompt before removing page when confirmation argument not passed', (done) => {
    fakeRestCalls();
    cmdInstance.action(
      {
        options: {
          debug: true,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a'
        }
      },
      () => {
        let promptIssued = false;

        if (promptOptions && promptOptions.type === 'confirm') {
          promptIssued = true;
        }

        try {
          assert(promptIssued);
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('should abort page removal when prompt not confirmed', (done) => {
    let postCallSpy = fakeRestCalls();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action(
      {
        options: {
          debug: true,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a'
        }
      },
      () => {
        try {
          assert(postCallSpy.notCalled === true);
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('automatically appends the .aspx extension', (done) => {
    fakeRestCalls();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action(
      {
        options: {
          debug: false,
          name: 'page',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          confirm: true
        }
      },
      () => {
        try {
          assert(cmdInstanceLogSpy.notCalled);
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('correctly handles OData error when removing modern page', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action(
      {
        options: {
          debug: false,
          name: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          confirm: true
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('supports debug mode', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying name', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--name') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webUrl', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying confirm', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--confirm') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if webUrl is not an absolute URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'page.aspx', webUrl: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { name: 'page.aspx', webUrl: 'http://foo' }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name and webURL specified and webUrl is a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { name: 'page.aspx', webUrl: 'https://contoso.sharepoint.com' }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when name has no extension', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { name: 'page', webUrl: 'https://contoso.sharepoint.com' }
    });
    assert.strictEqual(actual, true);
  });
});
