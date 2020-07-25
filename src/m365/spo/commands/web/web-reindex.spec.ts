import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./web-reindex');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { SpoPropertyBagBaseCommand } from '../propertybag/propertybag-base';
import * as chalk from 'chalk';

describe(commands.WEB_REINDEX, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
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
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.post,
      SpoPropertyBagBaseCommand.isNoScriptSite,
      SpoPropertyBagBaseCommand.setProperty
    ]);
    (command as any).reindexedLists = false
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
    assert.strictEqual(command.name.startsWith(commands.WEB_REINDEX), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('requests reindexing site that is not a no-script site for the first time', (done) => {
    let propertyName: string = '';
    let propertyValue: string = '';

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body.indexOf(`<Query Id="1" ObjectPathId="5">`) > -1) {
          return Promise.resolve(JSON.stringify([{
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7331.1206",
            "ErrorInfo": null,
            "TraceCorrelationId": "38e4499e-10a2-5000-ce25-77d4ccc2bd96"
          }, 7, {
            "_ObjectType_": "SP.Web",
            "_ObjectIdentity_": "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
            "ServerRelativeUrl": "\u002fsites\u002fteam-a"
          }]));
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/allproperties') > -1) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(SpoPropertyBagBaseCommand, 'isNoScriptSite').callsFake(() => Promise.resolve(false));
    sinon.stub(SpoPropertyBagBaseCommand, 'setProperty').callsFake((_propertyName, _propertyValue) => {
      propertyName = _propertyName;
      propertyValue = _propertyValue;
      return Promise.resolve(JSON.stringify({}));
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled, 'Something has been logged');
        assert.strictEqual(propertyName, 'vti_searchversion', 'Incorrect property stored in the property bag');
        assert.strictEqual(propertyValue, '1', 'Incorrect property value stored in the property bag');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('requests reindexing site that is not a no-script site for the second time', (done) => {
    let propertyName: string = '';
    let propertyValue: string = '';

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body.indexOf(`<Query Id="1" ObjectPathId="5">`) > -1) {
          return Promise.resolve(JSON.stringify([{
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7331.1206",
            "ErrorInfo": null,
            "TraceCorrelationId": "38e4499e-10a2-5000-ce25-77d4ccc2bd96"
          }, 7, {
            "_ObjectType_": "SP.Web",
            "_ObjectIdentity_": "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
            "ServerRelativeUrl": "\u002fsites\u002fteam-a"
          }]));
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/allproperties') > -1) {
        return Promise.resolve({
          vti_x005f_searchversion: '1'
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(SpoPropertyBagBaseCommand, 'isNoScriptSite').callsFake(() => Promise.resolve(false));
    sinon.stub(SpoPropertyBagBaseCommand, 'setProperty').callsFake((_propertyName, _propertyValue) => {
      propertyName = _propertyName;
      propertyValue = _propertyValue;
      return Promise.resolve(JSON.stringify({}));
    });

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        assert.strictEqual(propertyName, 'vti_searchversion', 'Incorrect property stored in the property bag');
        assert.strictEqual(propertyValue, '2', 'Incorrect property value stored in the property bag');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('requests reindexing no-script site', (done) => {
    const propertyName: string[] = [];
    const propertyValue: string[] = [];

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body.indexOf(`<Query Id="1" ObjectPathId="5">`) > -1) {
          return Promise.resolve(JSON.stringify([{
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7331.1206",
            "ErrorInfo": null,
            "TraceCorrelationId": "38e4499e-10a2-5000-ce25-77d4ccc2bd96"
          }, 7, {
            "_ObjectType_": "SP.Web",
            "_ObjectIdentity_": "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
            "ServerRelativeUrl": "\u002fsites\u002fteam-a"
          }]));
        }

        if (opts.body.indexOf(`<ObjectPath Id="10" ObjectPathId="9" />`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1206", "ErrorInfo": null, "TraceCorrelationId": "93e5499e-00f1-5000-1f36-3ab12512a7e9"
            }, 18, {
              "IsNull": false
            }, 19, {
              "_ObjectIdentity_": "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1"
            }, 20, {
              "_ObjectType_": "SP.Folder", "_ObjectIdentity_": "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1", "Properties": {
                "_ObjectType_": "SP.PropertyValues", "vti_folderitemcount$  Int32": 0, "vti_level$  Int32": 1, "vti_parentid": "{1C5271C8-DB93-459E-9C18-68FC33EFD856}", "vti_winfileattribs": "00000012", "vti_candeleteversion": "true", "vti_foldersubfolderitemcount$  Int32": 0, "vti_timelastmodified": "\/Date(2017,10,7,11,29,31,0)\/", "vti_dirlateststamp": "\/Date(2018,1,12,22,34,31,0)\/", "vti_isscriptable": "false", "vti_isexecutable": "false", "vti_metainfoversion$  Int32": 1, "vti_isbrowsable": "true", "vti_timecreated": "\/Date(2017,10,7,11,29,31,0)\/", "vti_etag": "\"{DF4291DE-226F-4C39-BBCC-DF21915F5FC1},256\"", "vti_hassubdirs": "true", "vti_docstoreversion$  Int32": 256, "vti_rtag": "rt:DF4291DE-226F-4C39-BBCC-DF21915F5FC1@00000000256", "vti_docstoretype$  Int32": 1, "vti_replid": "rid:{DF4291DE-226F-4C39-BBCC-DF21915F5FC1}"
              }
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists') > -1) {
        return Promise.resolve({
          value: [
            {
              NoCrawl: true,
              Title: 'Excluded from search'
            },
            {
              NoCrawl: false,
              Title: 'Included in search',
              RootFolder: {
                Properties: {},
                ServerRelativeUrl: '/sites/team-a/included-in-search'
              }
            },
            {
              NoCrawl: false,
              Title: 'Previously crawled',
              RootFolder: {
                Properties: {
                  vti_x005f_searchversion: 1
                },
                ServerRelativeUrl: '/sites/team-a/included-in-search'
              }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(SpoPropertyBagBaseCommand, 'isNoScriptSite').callsFake(() => Promise.resolve(true));
    sinon.stub(SpoPropertyBagBaseCommand, 'setProperty').callsFake((_propertyName, _propertyValue) => {
      propertyName.push(_propertyName);
      propertyValue.push(_propertyValue);
      return Promise.resolve(JSON.stringify({}));
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled, 'Something has been logged');
        assert.strictEqual(propertyName[0], 'vti_searchversion');
        assert.strictEqual(propertyName[1], 'vti_searchversion');
        assert.strictEqual(propertyValue[0], '1');
        assert.strictEqual(propertyValue[1], '2');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('requests reindexing no-script site (debug)', (done) => {
    const propertyName: string[] = [];
    const propertyValue: string[] = [];

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body.indexOf(`<Query Id="1" ObjectPathId="5">`) > -1) {
          return Promise.resolve(JSON.stringify([{
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7331.1206",
            "ErrorInfo": null,
            "TraceCorrelationId": "38e4499e-10a2-5000-ce25-77d4ccc2bd96"
          }, 7, {
            "_ObjectType_": "SP.Web",
            "_ObjectIdentity_": "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
            "ServerRelativeUrl": "\u002fsites\u002fteam-a"
          }]));
        }

        if (opts.body.indexOf(`<ObjectPath Id="10" ObjectPathId="9" />`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1206", "ErrorInfo": null, "TraceCorrelationId": "93e5499e-00f1-5000-1f36-3ab12512a7e9"
            }, 18, {
              "IsNull": false
            }, 19, {
              "_ObjectIdentity_": "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1"
            }, 20, {
              "_ObjectType_": "SP.Folder", "_ObjectIdentity_": "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1", "Properties": {
                "_ObjectType_": "SP.PropertyValues", "vti_folderitemcount$  Int32": 0, "vti_level$  Int32": 1, "vti_parentid": "{1C5271C8-DB93-459E-9C18-68FC33EFD856}", "vti_winfileattribs": "00000012", "vti_candeleteversion": "true", "vti_foldersubfolderitemcount$  Int32": 0, "vti_timelastmodified": "\/Date(2017,10,7,11,29,31,0)\/", "vti_dirlateststamp": "\/Date(2018,1,12,22,34,31,0)\/", "vti_isscriptable": "false", "vti_isexecutable": "false", "vti_metainfoversion$  Int32": 1, "vti_isbrowsable": "true", "vti_timecreated": "\/Date(2017,10,7,11,29,31,0)\/", "vti_etag": "\"{DF4291DE-226F-4C39-BBCC-DF21915F5FC1},256\"", "vti_hassubdirs": "true", "vti_docstoreversion$  Int32": 256, "vti_rtag": "rt:DF4291DE-226F-4C39-BBCC-DF21915F5FC1@00000000256", "vti_docstoretype$  Int32": 1, "vti_replid": "rid:{DF4291DE-226F-4C39-BBCC-DF21915F5FC1}"
              }
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists') > -1) {
        return Promise.resolve({
          value: [
            {
              NoCrawl: true,
              Title: 'Excluded from search'
            },
            {
              NoCrawl: false,
              Title: 'Included in search',
              RootFolder: {
                Properties: {},
                ServerRelativeUrl: '/sites/team-a/included-in-search'
              }
            },
            {
              NoCrawl: false,
              Title: 'Previously crawled',
              RootFolder: {
                Properties: {
                  vti_x005f_searchversion: 1
                },
                ServerRelativeUrl: '/sites/team-a/included-in-search'
              }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(SpoPropertyBagBaseCommand, 'isNoScriptSite').callsFake(() => Promise.resolve(true));
    sinon.stub(SpoPropertyBagBaseCommand, 'setProperty').callsFake((_propertyName, _propertyValue) => {
      propertyName.push(_propertyName);
      propertyValue.push(_propertyValue);
      return Promise.resolve(JSON.stringify({}));
    });

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.called, 'Nothing has been logged');
        assert.strictEqual(propertyName[0], 'vti_searchversion');
        assert.strictEqual(propertyName[1], 'vti_searchversion');
        assert.strictEqual(propertyValue[0], '1');
        assert.strictEqual(propertyValue[1], '2');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error while requiring reindexing a list', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.body.indexOf(`<Query Id="1" ObjectPathId="5">`) > -1) {
          return Promise.resolve(JSON.stringify([{
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.7331.1206",
            "ErrorInfo": null,
            "TraceCorrelationId": "38e4499e-10a2-5000-ce25-77d4ccc2bd96"
          }, 7, {
            "_ObjectType_": "SP.Web",
            "_ObjectIdentity_": "38e4499e-10a2-5000-ce25-77d4ccc2bd96|740c6a0b-85e2-48a0-a494-e0f1759d4a77:site:f3806c23-0c9f-42d3-bc7d-3895acc06d73:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d275",
            "ServerRelativeUrl": "\u002fsites\u002fteam-a"
          }]));
        }

        if (opts.body.indexOf(`<ObjectPath Id="10" ObjectPathId="9" />`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1206", "ErrorInfo": null, "TraceCorrelationId": "93e5499e-00f1-5000-1f36-3ab12512a7e9"
            }, 18, {
              "IsNull": false
            }, 19, {
              "_ObjectIdentity_": "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1"
            }, 20, {
              "_ObjectType_": "SP.Folder", "_ObjectIdentity_": "93e5499e-00f1-5000-1f36-3ab12512a7e9|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:f3806c23-0c9f-42d3-bc7d-3895acc06dc3:web:5a39e548-b3d7-4090-9cb9-0ce7cd85d2c5:folder:df4291de-226f-4c39-bbcc-df21915f5fc1", "Properties": {
                "_ObjectType_": "SP.PropertyValues", "vti_folderitemcount$  Int32": 0, "vti_level$  Int32": 1, "vti_parentid": "{1C5271C8-DB93-459E-9C18-68FC33EFD856}", "vti_winfileattribs": "00000012", "vti_candeleteversion": "true", "vti_foldersubfolderitemcount$  Int32": 0, "vti_timelastmodified": "\/Date(2017,10,7,11,29,31,0)\/", "vti_dirlateststamp": "\/Date(2018,1,12,22,34,31,0)\/", "vti_isscriptable": "false", "vti_isexecutable": "false", "vti_metainfoversion$  Int32": 1, "vti_isbrowsable": "true", "vti_timecreated": "\/Date(2017,10,7,11,29,31,0)\/", "vti_etag": "\"{DF4291DE-226F-4C39-BBCC-DF21915F5FC1},256\"", "vti_hassubdirs": "true", "vti_docstoreversion$  Int32": 256, "vti_rtag": "rt:DF4291DE-226F-4C39-BBCC-DF21915F5FC1@00000000256", "vti_docstoretype$  Int32": 1, "vti_replid": "rid:{DF4291DE-226F-4C39-BBCC-DF21915F5FC1}"
              }
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists') > -1) {
        return Promise.resolve({
          value: [
            {
              NoCrawl: true,
              Title: 'Excluded from search'
            },
            {
              NoCrawl: false,
              Title: 'Included in search',
              RootFolder: {
                Properties: {},
                ServerRelativeUrl: '/sites/team-a/included-in-search'
              }
            },
            {
              NoCrawl: false,
              Title: 'Previously crawled',
              RootFolder: {
                Properties: {
                  vti_x005f_searchversion: 1
                },
                ServerRelativeUrl: '/sites/team-a/included-in-search'
              }
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(SpoPropertyBagBaseCommand, 'isNoScriptSite').callsFake(() => Promise.resolve(true));
    sinon.stub(SpoPropertyBagBaseCommand, 'setProperty').callsFake(() => Promise.reject('ClientSvc unknown error'));

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('ClientSvc unknown error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if webUrl is valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });
});