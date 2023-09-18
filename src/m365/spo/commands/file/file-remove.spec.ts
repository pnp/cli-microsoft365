import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './file-remove.js';

describe(commands.FILE_REMOVE, () => {
  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];
  let promptIssued: boolean = false;

  before(() => {
    cli = Cli.getInstance();
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
    requests = [];
    sinon.stub(Cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.promptForConfirmation,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('excludes options from URL processing', () => {
    assert.deepStrictEqual((command as any).getExcludedOptionsWithUrls(), ['url']);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.strictEqual((alias && alias.indexOf(commands.PAGE_TEMPLATE_REMOVE) > -1), true);
  });

  it('prompts before removing file when confirmation argument not passed (id)', async () => {
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com', id: '0cd891ef-afce-4e55-b836-fce03286cccf' } });

    assert(promptIssued);
  });

  it('prompts before removing file when confirmation argument not passed (title)', async () => {
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com', id: '0cd891ef-afce-4e55-b836-fce03286cccf' } });

    assert(promptIssued);
  });

  it('aborts removing file when prompt not confirmed', async () => {
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com', id: '0cd891ef-afce-4e55-b836-fce03286cccf' } });
    assert(requests.length === 0);
  });

  it('removes the file when prompt confirmed (id)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/GetFileById(guid'`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com', id: '0cd891ef-afce-4e55-b836-fce03286cccf' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/GetFileById(guid'`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the file when webUrl does not includes a trailing /', async () => {
    const siteUrl: string = 'https://contoso.sharepoint.com';
    const fileUrl: string = 'SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/' + fileUrl)}')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, { options: { webUrl: siteUrl, url: fileUrl } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/' + fileUrl)}')`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the file when webUrl includes a trailing /', async () => {
    const siteUrl: string = 'https://contoso.sharepoint.com/';
    const fileUrl: string = 'SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/' + fileUrl)}')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, { options: { webUrl: siteUrl, url: fileUrl } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/' + fileUrl)}')`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the file when webUrl does not includes a trailing / and fileUrl is server relative', async () => {
    const siteUrl: string = 'https://contoso.sharepoint.com';
    const fileUrl: string = '/SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, { options: { webUrl: siteUrl, url: fileUrl } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the file when webUrl includes a trailing / and fileUrl is server relative', async () => {
    const siteUrl: string = 'https://contoso.sharepoint.com/';
    const fileUrl: string = '/SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, { options: { webUrl: siteUrl, url: fileUrl } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the file when webUrl (subsite) does not includes a trailing / ', async () => {
    const siteUrl: string = 'https://contoso.sharepoint.com/sites/subsite';
    const fileUrl: string = 'SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/subsite/' + fileUrl)}')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: siteUrl, url: fileUrl } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/subsite/' + fileUrl)}')`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the file when webUrl (subsite) includes a trailing /', async () => {
    const siteUrl: string = 'https://contoso.sharepoint.com/sites/subsite/';
    const fileUrl: string = 'SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/subsite/' + fileUrl)}')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, { options: { webUrl: siteUrl, url: fileUrl } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/subsite/' + fileUrl)}')`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the file when webUrl (subsite) does not includes a trailing / and fileUrl is server relative', async () => {
    const siteUrl: string = 'https://contoso.sharepoint.com/sites/subsite';
    const fileUrl: string = '/sites/subsite/SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: siteUrl, url: fileUrl } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the file when webUrl (subsite) includes a trailing / and fileUrl is server relative', async () => {
    const siteUrl: string = 'https://contoso.sharepoint.com/sites/subsite/';
    const fileUrl: string = '/sites/subsite/SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: siteUrl, url: fileUrl } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the file when webUrl (subsite) does not includes a trailing / and fileUrl is site relative', async () => {
    const siteUrl: string = 'https://contoso.sharepoint.com/sites/subsite';
    const fileUrl: string = 'SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/subsite/' + fileUrl)}')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: siteUrl, url: fileUrl } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/subsite/' + fileUrl)}')`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the file when webUrl (subsite) includes a trailing / and fileUrl is site relative', async () => {
    const siteUrl: string = 'https://contoso.sharepoint.com/sites/subsite/';
    const fileUrl: string = 'SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/subsite/' + fileUrl)}')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: siteUrl, url: fileUrl } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter('/sites/subsite/' + fileUrl)}')`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('recycles the file when prompt confirmed (id)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/recycle()`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com', id: '0cd891ef-afce-4e55-b836-fce03286cccf', recycle: true } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/recycle()`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the file when prompt confirmed (url)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com', url: '0cd891ef-afce-4e55-b836-fce03286cccf' } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/GetFileByServerRelativePath(DecodedUrl='`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('recycles the file when prompt confirmed (url)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/recycle()`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com', url: '0cd891ef-afce-4e55-b836-fce03286cccf', recycle: true } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/recycle()`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('command correctly handles file remove reject request', async () => {
    const err = 'An error has occurred';
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: err
          }
        }
      }
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById(') > -1) {
        throw error;
      }

      throw 'Invalid request';
    });

    const actionId: string = '0cd891ef-afce-4e55-b836-fce03286cccf';

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com',
        force: true
      }
    }), new CommandError(err));
  });

  it('uses correct API url when id option is passed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById(guid') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    await command.action(logger, {
      options: {
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com',
        force: true
      }
    });
  });

  it('uses correct API url when url option is passed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileByServerRelativePath(DecodedUrl=') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    const actionUrl: string = 'SharedDocuments/Test.docx';

    await command.action(logger, {
      options: {
        url: actionUrl,
        webUrl: 'https://contoso.sharepoint.com',
        force: true
      }
    });
  });

  it('uses correct API url when recycle option is passed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/recycle()') > -1) {
        return;
      }

      throw 'Invalid request';
    });

    const actionId: string = '0cd891ef-afce-4e55-b836-fce03286cccf';

    await command.action(logger, {
      options: {
        id: actionId,
        recycle: true,
        webUrl: 'https://contoso.sharepoint.com',
        force: true
      }
    });
  });

  it('fails validation if both id and title options are not passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', id: '0cd891ef-afce-4e55-b836-fce03286cccf' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if both id and url options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', url: 'Documents' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
