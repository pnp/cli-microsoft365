import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-add.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.APP_ADD, () => {
  const fsStats: fs.Stats = {
    isDirectory: () => false,
    isFile: () => false,
    isBlockDevice: () => false,
    isCharacterDevice: () => false,
    isSymbolicLink: () => false,
    isFIFO: () => false,
    isSocket: () => false,
    dev: 0,
    ino: 0,
    mode: 0,
    nlink: 0,
    uid: 0,
    gid: 0,
    rdev: 0,
    size: 0,
    blksize: 0,
    blocks: 0,
    atimeMs: 0,
    mtimeMs: 0,
    ctimeMs: 0,
    birthtimeMs: 0,
    atime: new Date(),
    mtime: new Date(),
    ctime: new Date(),
    birthtime: new Date()
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let requests: any[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
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
    requests = [];
    sinon.stub(request, 'get').resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      fs.readFileSync,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds new app to the tenant app catalog', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return '{"CheckInComment":"","CheckOutType":2,"ContentTag":"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4,3","CustomizedPageStatus":0,"ETag":"\\"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4\\"","Exists":true,"IrmEnabled":false,"Length":"3752","Level":1,"LinkingUri":null,"LinkingUrl":"","MajorVersion":3,"MinorVersion":0,"Name":"spfx-01.sppkg","ServerRelativeUrl":"/sites/apps/AppCatalog/spfx.sppkg","TimeCreated":"2018-05-25T06:59:20Z","TimeLastModified":"2018-05-25T08:23:18Z","Title":"spfx-01-client-side-solution","UIVersion":1536,"UIVersionLabel":"3.0","UniqueId":"bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"}';
        }
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'readFileSync').returns('123');

    await command.action(logger, { options: { filePath: 'spfx.sppkg', output: 'text' } });
    assert(loggerLogSpy.calledWith("bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"));
  });

  it('adds new app to the tenant app catalog (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return '{"CheckInComment":"","CheckOutType":2,"ContentTag":"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4,3","CustomizedPageStatus":0,"ETag":"\\"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4\\"","Exists":true,"IrmEnabled":false,"Length":"3752","Level":1,"LinkingUri":null,"LinkingUrl":"","MajorVersion":3,"MinorVersion":0,"Name":"spfx-01.sppkg","ServerRelativeUrl":"/sites/apps/AppCatalog/spfx.sppkg","TimeCreated":"2018-05-25T06:59:20Z","TimeLastModified":"2018-05-25T08:23:18Z","Title":"spfx-01-client-side-solution","UIVersion":1536,"UIVersionLabel":"3.0","UniqueId":"bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"}';
        }
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'readFileSync').returns('123');
    try {
      await command.action(logger, { options: { debug: true, filePath: 'spfx.sppkg' } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/tenantappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0 &&
          r.headers.binaryStringRequestBody &&
          r.data) {
          correctRequestIssued = true;
        }
      });

      assert(correctRequestIssued);
    }
    finally {
      sinonUtil.restore([
        request.post,
        fs.readFileSync
      ]);
    }
  });

  it('adds new app to a site app catalog (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return '{"CheckInComment":"","CheckOutType":2,"ContentTag":"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4,3","CustomizedPageStatus":0,"ETag":"\\"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4\\"","Exists":true,"IrmEnabled":false,"Length":"3752","Level":1,"LinkingUri":null,"LinkingUrl":"","MajorVersion":3,"MinorVersion":0,"Name":"spfx-01.sppkg","ServerRelativeUrl":"/sites/apps/AppCatalog/spfx.sppkg","TimeCreated":"2018-05-25T06:59:20Z","TimeLastModified":"2018-05-25T08:23:18Z","Title":"spfx-01-client-side-solution","UIVersion":1536,"UIVersionLabel":"3.0","UniqueId":"bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"}';
        }
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'readFileSync').returns('123');

    try {
      await command.action(logger, { options: { debug: true, filePath: 'spfx.sppkg', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } });
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/sitecollectionappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0 &&
          r.headers.binaryStringRequestBody &&
          r.data) {
          correctRequestIssued = true;
        }
      });

      assert(correctRequestIssued);
    }
    finally {
      sinonUtil.restore([
        request.post,
        fs.readFileSync
      ]);
    }
  });

  it('returns all info about the added app in the JSON output mode', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return '{"CheckInComment":"","CheckOutType":2,"ContentTag":"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4,3","CustomizedPageStatus":0,"ETag":"\\"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4\\"","Exists":true,"IrmEnabled":false,"Length":"3752","Level":1,"LinkingUri":null,"LinkingUrl":"","MajorVersion":3,"MinorVersion":0,"Name":"spfx-01.sppkg","ServerRelativeUrl":"/sites/apps/AppCatalog/spfx.sppkg","TimeCreated":"2018-05-25T06:59:20Z","TimeLastModified":"2018-05-25T08:23:18Z","Title":"spfx-01-client-side-solution","UIVersion":1536,"UIVersionLabel":"3.0","UniqueId":"bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"}';
        }
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'readFileSync').returns('123');

    try {
      await command.action(logger, { options: { filePath: 'spfx.sppkg', output: 'json' } });
      assert(loggerLogSpy.calledWith(JSON.parse('{"CheckInComment":"","CheckOutType":2,"ContentTag":"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4,3","CustomizedPageStatus":0,"ETag":"\\"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4\\"","Exists":true,"IrmEnabled":false,"Length":"3752","Level":1,"LinkingUri":null,"LinkingUrl":"","MajorVersion":3,"MinorVersion":0,"Name":"spfx-01.sppkg","ServerRelativeUrl":"/sites/apps/AppCatalog/spfx.sppkg","TimeCreated":"2018-05-25T06:59:20Z","TimeLastModified":"2018-05-25T08:23:18Z","Title":"spfx-01-client-side-solution","UIVersion":1536,"UIVersionLabel":"3.0","UniqueId":"bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"}')));
    }
    finally {
      sinonUtil.restore([
        request.post,
        fs.readFileSync
      ]);
    }
  });

  it('correctly handles failure when the app already exists in the tenant app catalog', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          throw { error: JSON.stringify({ "odata.error": { "code": "-2130575257, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "A file with the name AppCatalog/spfx.sppkg already exists. It was last modified by i:0#.f|membership|admin@contoso.onmi on 24 Nov 2017 12:50:43 -0800." } } }) };
        }
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'readFileSync').returns('123');

    try {
      await assert.rejects(command.action(logger, { options: { debug: true, filePath: 'spfx.sppkg' } } as any),
        new CommandError('A file with the name AppCatalog/spfx.sppkg already exists. It was last modified by i:0#.f|membership|admin@contoso.onmi on 24 Nov 2017 12:50:43 -0800.'));
    }
    finally {
      sinonUtil.restore([
        request.post,
        fs.readFileSync
      ]);
    }
  });

  it('correctly handles failure when the app already exists in the site app catalog', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.reject({
            error: JSON.stringify({ "odata.error": { "code": "-2130575257, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "A file with the name AppCatalog/spfx.sppkg already exists. It was last modified by i:0#.f|membership|admin@contoso.onmi on 24 Nov 2017 12:50:43 -0800." } } })
          });
        }
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'readFileSync').returns('123');

    try {
      await assert.rejects(command.action(logger, { options: { debug: true, filePath: 'spfx.sppkg', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any),
        new CommandError('A file with the name AppCatalog/spfx.sppkg already exists. It was last modified by i:0#.f|membership|admin@contoso.onmi on 24 Nov 2017 12:50:43 -0800.'));
    }
    finally {
      sinonUtil.restore([
        request.post,
        fs.readFileSync
      ]);
    }
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.reject({ error: 'An error has occurred' });
        }
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'readFileSync').returns('123');

    try {
      await assert.rejects(command.action(logger, { options: { debug: true, filePath: 'spfx.sppkg' } } as any), new CommandError('An error has occurred'));
    }
    finally {
      sinonUtil.restore([
        request.post,
        fs.readFileSync
      ]);
    }
  });

  it('correctly handles random API error when sitecollection', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.reject({ error: 'An error has occurred' });
        }
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'readFileSync').returns('123');

    try {
      await assert.rejects(command.action(logger, { options: { debug: true, filePath: 'spfx.sppkg', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any),
        new CommandError('An error has occurred'));
    }
    finally {
      sinonUtil.restore([
        request.post,
        fs.readFileSync
      ]);
    }
  });

  it('correctly handles random API error (string error)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/tenantappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.reject('An error has occurred');
        }
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'readFileSync').returns('123');

    try {
      await assert.rejects(command.action(logger, { options: { debug: true, filePath: 'spfx.sppkg' } } as any),
        new CommandError('An error has occurred'));
    }
    finally {
      sinonUtil.restore([
        request.post,
        fs.readFileSync
      ]);
    }
  });

  it('correctly handles random API error when sitecollection (string error)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/sitecollectionappcatalog/Add(overwrite=false, url='spfx.sppkg')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return Promise.reject('An error has occurred');
        }
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'readFileSync').returns('123');

    try {
      await assert.rejects(command.action(logger, { options: { debug: true, filePath: 'spfx.sppkg', appCatalogScope: 'sitecollection', appCatalogUrl: 'https://contoso.sharepoint.com' } } as any),
        new CommandError('An error has occurred'));
    }
    finally {
      sinonUtil.restore([
        request.post,
        fs.readFileSync
      ]);
    }
  });

  it('handles promise error while getting tenant appcatalog', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true, filePath: 'spfx.sppkg'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('handles error while getting tenant appcatalog', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
              "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "18091989-62a6-4cad-9717-29892ee711bc", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.ServerException"
            }, "TraceCorrelationId": "18091989-62a6-4cad-9717-29892ee711bc"
          }
        ]);
      }
      if ((opts.url as string).indexOf('contextinfo') > -1) {
        return 'abc';
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true, filePath: 'spfx.sppkg'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('fails validation on invalid scope', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appCatalogScope: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation on valid \'tenant\' scope', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    const actual = await command.validate({ options: { appCatalogScope: 'tenant', filePath: 'abc' } }, commandInfo);
    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });

  it('passes validation on valid \'Tenant\' scope', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    const actual = await command.validate({ options: { appCatalogScope: 'Tenant', filePath: 'abc' } }, commandInfo);
    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });

  it('passes validation on valid \'SiteCollection\' scope', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    const actual = await command.validate({ options: { appCatalogScope: 'SiteCollection', appCatalogUrl: 'https://contoso.sharepoint.com', filePath: 'abc' } }, commandInfo);
    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });

  it('submits to tenant app catalog when scope not specified', async () => {
    // setup call to fake requests...
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return '{"CheckInComment":"","CheckOutType":2,"ContentTag":"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4,3","CustomizedPageStatus":0,"ETag":"\\"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4\\"","Exists":true,"IrmEnabled":false,"Length":"3752","Level":1,"LinkingUri":null,"LinkingUrl":"","MajorVersion":3,"MinorVersion":0,"Name":"spfx-01.sppkg","ServerRelativeUrl":"/sites/apps/AppCatalog/spfx.sppkg","TimeCreated":"2018-05-25T06:59:20Z","TimeLastModified":"2018-05-25T08:23:18Z","Title":"spfx-01-client-side-solution","UIVersion":1536,"UIVersionLabel":"3.0","UniqueId":"bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"}';
        }
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'readFileSync').returns('123');

    try {
      await command.action(logger, { options: { filePath: 'spfx.sppkg' } });
      let correctAppCatalogUsed = false;
      requests.forEach(r => {
        if (r.url.indexOf('/tenantappcatalog/') > -1) {
          correctAppCatalogUsed = true;
        }
      });

      assert(correctAppCatalogUsed);
    }
    finally {
      sinonUtil.restore([
        request.post,
        fs.readFileSync
      ]);
    }
  });

  it('submits to tenant app catalog when scope \'tenant\' specified ', async () => {
    // setup call to fake requests...
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return '{"CheckInComment":"","CheckOutType":2,"ContentTag":"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4,3","CustomizedPageStatus":0,"ETag":"\\"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4\\"","Exists":true,"IrmEnabled":false,"Length":"3752","Level":1,"LinkingUri":null,"LinkingUrl":"","MajorVersion":3,"MinorVersion":0,"Name":"spfx-01.sppkg","ServerRelativeUrl":"/sites/apps/AppCatalog/spfx.sppkg","TimeCreated":"2018-05-25T06:59:20Z","TimeLastModified":"2018-05-25T08:23:18Z","Title":"spfx-01-client-side-solution","UIVersion":1536,"UIVersionLabel":"3.0","UniqueId":"bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"}';
        }
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    try {
      await command.action(logger, { options: { appCatalogScope: 'tenant', filePath: 'spfx.sppkg' } });
      let correctAppCatalogUsed = false;
      requests.forEach(r => {
        if (r.url.indexOf('/tenantappcatalog/') > -1) {
          correctAppCatalogUsed = true;
        }
      });
      assert(correctAppCatalogUsed);
    }
    finally {
      sinonUtil.restore([
        request.post,
        fs.readFileSync
      ]);
    }
  });

  it('submits to sitecollection app catalog when scope \'sitecollection\' specified ', async () => {
    // setup call to fake requests...
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_api/web/`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0 &&
          opts.headers.binaryStringRequestBody &&
          opts.data) {
          return '{"CheckInComment":"","CheckOutType":2,"ContentTag":"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4,3","CustomizedPageStatus":0,"ETag":"\\"{BDA5CE2F-9AC7-4A6F-A98B-7AE1C168519E},4\\"","Exists":true,"IrmEnabled":false,"Length":"3752","Level":1,"LinkingUri":null,"LinkingUrl":"","MajorVersion":3,"MinorVersion":0,"Name":"spfx-01.sppkg","ServerRelativeUrl":"/sites/apps/AppCatalog/spfx.sppkg","TimeCreated":"2018-05-25T06:59:20Z","TimeLastModified":"2018-05-25T08:23:18Z","Title":"spfx-01-client-side-solution","UIVersion":1536,"UIVersionLabel":"3.0","UniqueId":"bda5ce2f-9ac7-4a6f-a98b-7ae1c168519e"}';
        }
      }

      throw 'Invalid request';
    });
    sinon.stub(fs, 'readFileSync').returns('123');

    try {
      await command.action(logger, { options: { appCatalogScope: 'sitecollection', filePath: 'spfx.sppkg', appCatalogUrl: 'https://contoso.sharepoint.com' } });
      let correctAppCatalogUsed = false;
      requests.forEach(r => {
        if (r.url.indexOf('/sitecollectionappcatalog/') > -1) {
          correctAppCatalogUsed = true;
        }
      });
      assert(correctAppCatalogUsed);
    }
    finally {
      sinonUtil.restore([
        request.post,
        fs.readFileSync
      ]);
    }
  });

  it('fails validation if file path doesn\'t exist', async () => {
    sinon.stub(fs, 'existsSync').returns(false);
    const actual = await command.validate({ options: { filePath: 'abc' } }, commandInfo);
    sinonUtil.restore(fs.existsSync);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if file path points to a directory', async () => {
    const stats = { ...fsStats, isDirectory: () => true };
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(stats);
    const actual = await command.validate({ options: { filePath: 'abc' } }, commandInfo);
    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid scope is specified', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    const actual = await command.validate({ options: { filePath: 'abc', appCatalogScope: 'foo' } }, commandInfo);

    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when path points to a valid file', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    const actual = await command.validate({ options: { filePath: 'abc' } }, commandInfo);

    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });

  it('passes validation when no scope is specified', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    const actual = await command.validate({ options: { filePath: 'abc' } }, commandInfo);

    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the scope is specified with \'tenant\'', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    const actual = await command.validate({ options: { filePath: 'abc', appCatalogScope: 'tenant' } }, commandInfo);

    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });


  it('should fail when \'sitecollection\' scope, but no appCatalogUrl specified', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    const actual = await command.validate({ options: { filePath: 'abc', appCatalogScope: 'sitecollection' } }, commandInfo);

    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.notStrictEqual(actual, true);
  });

  it('should not fail when \'tenant\' scope, but also appCatalogUrl specified', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    const actual = await command.validate({ options: { filePath: 'abc', appCatalogScope: 'tenant', appCatalogUrl: 'https://contoso.sharepoint.com' } }, commandInfo);

    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });

  it('should fail when \'sitecollection\' scope, but bad appCatalogUrl format specified', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'lstatSync').returns(fsStats);

    const actual = await command.validate({ options: { filePath: 'abc', appCatalogScope: 'sitecollection', appCatalogUrl: 'contoso.sharepoint.com' } }, commandInfo);

    sinonUtil.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.notStrictEqual(actual, true);
  });
});
