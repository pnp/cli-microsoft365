import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./file-rename');
const fileRemoveCommand: Command = require('./file-remove');

describe(commands.FILE_RENAME, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const renameResponseJson = [
    {
      'ErrorCode': 0,
      'ErrorMessage': null,
      'FieldName': 'FileLeafRef',
      'FieldValue': 'test 2.docx',
      'HasException': false,
      'ItemId': 642
    }
  ];

  const renameValue = {
    value: renameResponseJson
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get, 
      request.post,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FILE_RENAME), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'foo', sourceUrl: 'abc', targetUrl: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', sourceUrl: 'abc', targetUrl: 'abc' } });
    assert.strictEqual(actual, true);
  });

  it('should command complete successfully', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args) : Promise<any> => {
      if (command === fileRemoveCommand) {
        if (args.options.webUrl === 'https://contoso.sharepoint.com/sites/portal') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('ValidateUpdateListItem()') > -1) {
        return Promise.resolve(renameValue);
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`GetFileByServerRelativeUrl`) > -1) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: 
      { 
        webUrl: 'https://contoso.sharepoint.com/sites/portal', 
        sourceUrl: 'Shared Documents/abc.pdf',
        force: true,
        targetFileName: 'def.pdf'
      } }, () => {
      try {
        assert(loggerLogSpy.calledWith(renameResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should command complete successfully with all options', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args) : Promise<any> => {
      if (command === fileRemoveCommand) {
        if (args.options.webUrl === 'https://contoso.sharepoint.com/sites/portal') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('ValidateUpdateListItem()') > -1) {
        return Promise.resolve(renameValue);
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`GetFileByServerRelativeUrl`) > -1) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        sourceUrl: 'Shared Documents/abc.pdf',
        force: true,
        targetFileName: 'def.pdf'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(renameResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should command complete successfully and not check for recycle', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('ValidateUpdateListItem()') > -1) {
        return Promise.resolve(renameValue);
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`GetFileByServerRelativeUrl`) > -1) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        sourceUrl: 'Shared Documents/abc.pdf',
        targetFileName: 'def.pdf'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(renameResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should continue if file cannot be recycled because it does not exist', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === fileRemoveCommand) {
        if (args.options.webUrl === 'https://contoso.sharepoint.com/sites/portal') {
          return Promise.reject({
            error: {
              message: 'File does not exist'
            }
          });
        }
        return Promise.reject(new CommandError('Invalid URL'));
      }
      return Promise.reject(new CommandError('Unknown case'));
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('ValidateUpdateListItem()') > -1) {
        return Promise.resolve(renameValue);
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`GetFileByServerRelativeUrl`) > -1) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        sourceUrl: 'Shared Documents/abc.pdf',
        force: true,
        targetFileName: 'def.pdf'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(renameResponseJson));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should throw error if file cannot be recycled', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === fileRemoveCommand) {
        if (args.options.webUrl === 'https://contoso.sharepoint.com/sites/portal') {
          return Promise.reject({
            error: {
              message: 'Locked for use'
            }
          });
        }
        return Promise.reject(new CommandError('Invalid URL'));
      }
      return Promise.reject(new CommandError('Unknown case'));
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`GetFileByServerRelativeUrl`) > -1) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/sites/portal',
        sourceUrl: 'Shared Documents/abc.pdf',
        force: true,
        targetFileName: 'def.pdf'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Locked for use')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });
});