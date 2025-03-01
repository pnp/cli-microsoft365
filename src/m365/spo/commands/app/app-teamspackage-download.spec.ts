import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { PassThrough } from 'stream';
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
import command from './app-teamspackage-download.js';

describe(commands.APP_TEAMSPACKAGE_DOWNLOAD, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      fs.existsSync,
      fs.createWriteStream
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_TEAMSPACKAGE_DOWNLOAD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('downloads Teams app package when appItemUniqueId specified', async () => {
    const mockResponse = `{"data": 123}`;
    const responseStream = new PassThrough();
    responseStream.write(mockResponse);
    responseStream.end(); //Mark that we pushed all the data.

    const writeStream = new PassThrough();
    const fsStub = sinon.stub(fs, 'createWriteStream').returns(writeStream as any);

    setTimeout(() => {
      writeStream.emit('close');
    }, 5);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/appcatalog" };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetList('/sites/appcatalog/AppCatalog')/GetItemByUniqueId('335a5612-3e85-462d-9d5b-c014b5abeac4')?$expand=File&$select=Id,File/Name`) {
        return {
          "File": {
            "Name": "m365-spfx-wellbeing.sppkg"
          },
          "Id": 2,
          "ID": 2
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/tenantappcatalog/downloadteamssolution(2)/$value`) {
        return {
          data: responseStream
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { appItemUniqueId: '335a5612-3e85-462d-9d5b-c014b5abeac4' } });
    assert(fsStub.calledOnce);
  });

  it('downloads Teams app package when appItemId specified', async () => {
    const mockResponse = `{"data": 123}`;
    const responseStream = new PassThrough();
    responseStream.write(mockResponse);
    responseStream.end(); //Mark that we pushed all the data.

    const writeStream = new PassThrough();
    const fsStub = sinon.stub(fs, 'createWriteStream').returns(writeStream as any);

    setTimeout(() => {
      writeStream.emit('close');
    }, 5);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/appcatalog" };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetList('/sites/appcatalog/AppCatalog')/GetItemById(2)?$expand=File&$select=File/Name`) {
        return {
          "File": {
            "Name": "m365-spfx-wellbeing.sppkg"
          }
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/tenantappcatalog/downloadteamssolution(2)/$value`) {
        return { data: responseStream };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { appItemId: 2 } });
    assert(fsStub.calledOnce);
  });

  it('downloads Teams app package when appItemId specified (debug)', async () => {
    const mockResponse = `{"data": 123}`;
    const responseStream = new PassThrough();
    responseStream.write(mockResponse);
    responseStream.end(); //Mark that we pushed all the data.

    const writeStream = new PassThrough();
    const fsStub = sinon.stub(fs, 'createWriteStream').returns(writeStream as any);

    setTimeout(() => {
      writeStream.emit('close');
    }, 5);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/appcatalog" };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetList('/sites/appcatalog/AppCatalog')/GetItemById(2)?$expand=File&$select=File/Name`) {
        return {
          "File": {
            "Name": "m365-spfx-wellbeing.sppkg"
          }
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/tenantappcatalog/downloadteamssolution(2)/$value`) {
        return { data: responseStream };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { appItemId: 2, debug: true } });
    assert(fsStub.calledOnce);
  });

  it('downloads Teams app package when appName specified', async () => {
    const mockResponse = `{"data": 123}`;
    const responseStream = new PassThrough();
    responseStream.write(mockResponse);
    responseStream.end(); //Mark that we pushed all the data.

    const writeStream = new PassThrough();
    const fsStub = sinon.stub(fs, 'createWriteStream').returns(writeStream as any);

    setTimeout(() => {
      writeStream.emit('close');
    }, 5);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/appcatalog" };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('m365-spfx-wellbeing.sppkg')/ListItemAllFields?$select=Id`) {
        return {
          "Id": 2,
          "ID": 2
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/tenantappcatalog/downloadteamssolution(2)/$value`) {
        return {
          data: responseStream
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { appName: 'm365-spfx-wellbeing.sppkg' } });
    assert(fsStub.calledOnce);
  });

  it('saves the downloaded Teams package to file with name following the .sppkg file', async () => {
    const mockResponse = `{"data": 123}`;
    const responseStream = new PassThrough();
    responseStream.write(mockResponse);
    responseStream.end(); //Mark that we pushed all the data.

    const writeStream = new PassThrough();
    const fsStub = sinon.stub(fs, 'createWriteStream').returns(writeStream as any);

    setTimeout(() => {
      writeStream.emit('close');
    }, 5);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/appcatalog" };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('m365-spfx-wellbeing.sppkg')/ListItemAllFields?$select=Id`) {
        return {
          "Id": 2,
          "ID": 2
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/tenantappcatalog/downloadteamssolution(2)/$value`) {
        return {
          data: responseStream
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { appName: 'm365-spfx-wellbeing.sppkg' } });
    assert(fsStub.calledWith('m365-spfx-wellbeing.zip'));
  });

  it('saves the app package downloaded using appItemUniqueId to the specified file', async () => {
    const mockResponse = `{"data": 123}`;
    const responseStream = new PassThrough();
    responseStream.write(mockResponse);
    responseStream.end(); //Mark that we pushed all the data.

    const writeStream = new PassThrough();
    const fsStub = sinon.stub(fs, 'createWriteStream').returns(writeStream as any);

    setTimeout(() => {
      writeStream.emit('close');
    }, 5);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/appcatalog" };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetList('/sites/appcatalog/AppCatalog')/GetItemByUniqueId('335a5612-3e85-462d-9d5b-c014b5abeac4')?$expand=File&$select=Id,File/Name`) {
        return {
          "File": {
            "Name": "m365-spfx-wellbeing.sppkg"
          },
          "Id": 2,
          "ID": 2
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/tenantappcatalog/downloadteamssolution(2)/$value`) {
        return {
          data: responseStream
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { appItemUniqueId: '335a5612-3e85-462d-9d5b-c014b5abeac4', fileName: 'package.zip' } });
    assert(fsStub.calledWith('package.zip'));
  });

  it('saves the app package downloaded using appItemId to the specified file', async () => {
    const mockResponse = `{"data": 123}`;
    const responseStream = new PassThrough();
    responseStream.write(mockResponse);
    responseStream.end(); //Mark that we pushed all the data.

    const writeStream = new PassThrough();
    const fsStub = sinon.stub(fs, 'createWriteStream').returns(writeStream as any);

    setTimeout(() => {
      writeStream.emit('close');
    }, 5);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/appcatalog" };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetList('/sites/appcatalog/AppCatalog')/GetItemById(2)?$expand=File&$select=File/Name`) {
        return {
          "File": {
            "Name": "m365-spfx-wellbeing.sppkg"
          }
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/tenantappcatalog/downloadteamssolution(2)/$value`) {
        return {
          data: responseStream
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { appItemId: 2, fileName: 'package.zip' } });
    assert(fsStub.calledWith('package.zip'));
  });

  it('saves the app package downloaded using appName to the specified file', async () => {
    const mockResponse = `{"data": 123}`;
    const responseStream = new PassThrough();
    responseStream.write(mockResponse);
    responseStream.end(); //Mark that we pushed all the data.

    const writeStream = new PassThrough();
    const fsStub = sinon.stub(fs, 'createWriteStream').returns(writeStream as any);

    setTimeout(() => {
      writeStream.emit('close');
    }, 5);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/appcatalog" };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('m365-spfx-wellbeing.sppkg')/ListItemAllFields?$select=Id`) {
        return {
          "Id": 2,
          "ID": 2
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/tenantappcatalog/downloadteamssolution(2)/$value`) {
        return {
          data: responseStream
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { appName: 'm365-spfx-wellbeing.sppkg', fileName: 'package.zip' } });
    assert(fsStub.calledWith('package.zip'));
  });

  it(`doesn't detect app catalog URL when specified`, async () => {
    const mockResponse = `{"data": 123}`;
    const responseStream = new PassThrough();
    responseStream.write(mockResponse);
    responseStream.end(); //Mark that we pushed all the data.

    const writeStream = new PassThrough();
    const fsStub = sinon.stub(fs, 'createWriteStream').returns(writeStream as any);

    setTimeout(() => {
      writeStream.emit('close');
    }, 5);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetList('/sites/appcatalog/AppCatalog')/GetItemById(2)?$expand=File&$select=File/Name`) {
        return {
          "File": {
            "Name": "m365-spfx-wellbeing.sppkg"
          }
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/tenantappcatalog/downloadteamssolution(2)/$value`) {
        return {
          data: responseStream
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { appItemId: 2, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } });
    assert(fsStub.calledOnce);
  });

  it(`handles error when the specified app catalog doesn't exist`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('m365-spfx-wellbeing.sppkg')/ListItemAllFields?$select=Id`) {
        return Promise.reject('404 FILE NOT FOUND');
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { appName: 'm365-spfx-wellbeing.sppkg', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }),
      new CommandError(`404 FILE NOT FOUND`));
  });

  it(`handles error when the specified appItemUniqueId doesn't exist`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/appcatalog" };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetList('/sites/appcatalog/AppCatalog')/GetItemByUniqueId('335a5612-3e85-462d-9d5b-c014b5abeac4')?$expand=File&$select=Id,File/Name`) {
        throw {
          error: {
            "odata.error": {
              "code": "-2147024809, System.ArgumentException",
              "message": {
                "lang": "en-US",
                "value": "Value does not fall within the expected range."
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { appItemUniqueId: '335a5612-3e85-462d-9d5b-c014b5abeac4' } }),
      new CommandError(`Value does not fall within the expected range.`));
  });

  it(`handles error when the specified appItemId doesn't exist`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/appcatalog" };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetList('/sites/appcatalog/AppCatalog')/GetItemById(2)?$expand=File&$select=File/Name`) {
        throw {
          error: {
            "odata.error": {
              "code": "-2147024809, System.ArgumentException",
              "message": {
                "lang": "en-US",
                "value": "Item does not exist. It may have been deleted by another user."
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { appItemId: 2 } }),
      new CommandError('Item does not exist. It may have been deleted by another user.'));
  });

  it(`handles error when the specified appName doesn't exist`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/appcatalog" };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('m365-spfx-wellbeing.sppkg')/ListItemAllFields?$select=Id`) {
        throw {
          error: {
            "odata.error": {
              "code": "-2147024894, System.IO.FileNotFoundException",
              "message": {
                "lang": "en-US",
                "value": "File Not Found."
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { appName: 'm365-spfx-wellbeing.sppkg' } }),
      new CommandError('File Not Found.'));
  });

  it(`handles error when the package doesn't support syncing to Teams`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/appcatalog" };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetList('/sites/appcatalog/AppCatalog')/GetItemById(2)?$expand=File&$select=File/Name`) {
        return {
          "File": {
            "Name": "m365-spfx-wellbeing.sppkg"
          }
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/tenantappcatalog/downloadteamssolution(2)/$value`) {
        throw 'Request failed with status code 404';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { appItemId: 2 } }),
      new CommandError('Request failed with status code 404'));
  });

  it(`handles error when saving the package to file fails`, async () => {
    const mockResponse = `{"data": 123}`;
    const responseStream = new PassThrough();
    responseStream.write(mockResponse);
    responseStream.end(); //Mark that we pushed all the data.

    const writeStream = new PassThrough();
    sinon.stub(fs, 'createWriteStream').returns(writeStream as any);

    setTimeout(() => {
      writeStream.emit('error', "An error has occurred");
    }, 5);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/appcatalog" };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/GetList('/sites/appcatalog/AppCatalog')/GetItemById(2)?$expand=File&$select=File/Name`) {
        return {
          "File": {
            "Name": "m365-spfx-wellbeing.sppkg"
          }
        };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/appcatalog/_api/web/tenantappcatalog/downloadteamssolution(2)/$value`) {
        return {
          data: responseStream
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { appItemId: 2 } }),
      new CommandError('An error has occurred'));
  });

  it('fails validation if the appCatalogUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { appItemId: 1, appCatalogUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the appCatalogUrl is a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { appItemId: 1, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the appCatalogUrl is not specified', async () => {
    const actual = await command.validate({ options: { appItemId: 1 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when the appItemId is not a number', async () => {
    const actual = await command.validate({ options: { appItemId: 'a' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the appItemId is a number', async () => {
    const actual = await command.validate({ options: { appItemId: 1 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when the appItemUniqueId is not a GUID', async () => {
    const actual = await command.validate({ options: { appItemUniqueId: 'a' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the appItemUniqueId is a GUID', async () => {
    const actual = await command.validate({ options: { appItemUniqueId: '335a5612-3e85-462d-9d5b-c014b5abeac4' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when the specified file already exists', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = await command.validate({ options: { appItemId: 1, fileName: 'file.zip' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the specified file does not exist', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = await command.validate({ options: { appItemId: 1, fileName: 'file.zip' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when the appItemUniqueId and appItemId specified', async () => {
    const actual = await command.validate({ options: { appItemUniqueId: '335a5612-3e85-462d-9d5b-c014b5abeac4', appItemId: 1 } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the appItemUniqueId and appName specified', async () => {
    const actual = await command.validate({ options: { appItemUniqueId: '335a5612-3e85-462d-9d5b-c014b5abeac4', appName: 'app.sppkg' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the appItemId and appName specified', async () => {
    const actual = await command.validate({ options: { appItemId: 1, appName: 'app.sppkg' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when no app identifier specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
