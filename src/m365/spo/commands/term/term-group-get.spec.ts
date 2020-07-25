import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./term-group-get');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.TERM_GROUP_GET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TERM_GROUP_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets taxonomy term group by id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="25" ObjectPathId="24" /><ObjectIdentityQuery Id="26" ObjectPathId="24" /><ObjectPath Id="28" ObjectPathId="27" /><ObjectIdentityQuery Id="29" ObjectPathId="27" /><ObjectPath Id="31" ObjectPathId="30" /><ObjectPath Id="33" ObjectPathId="32" /><ObjectIdentityQuery Id="34" ObjectPathId="32" /><Query Id="35" ObjectPathId="32"><Query SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticMethod Id="24" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="27" ParentId="24" Name="GetDefaultSiteCollectionTermStore" /><Property Id="30" ParentId="27" Name="Groups" /><Method Id="32" ParentId="30" Name="GetById"><Parameters><Parameter Type="Guid">{36a62501-17ea-455a-bed4-eff862242def}</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.8105.1217",
            "ErrorInfo": null,
            "TraceCorrelationId": "aa58909e-60c1-0000-29c7-003b321d02d1"
          },
          25,
          {
            "IsNull": false
          },
          26,
          {
            "_ObjectIdentity_": "aa58909e-60c1-0000-29c7-003b321d02d1|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          },
          28,
          {
            "IsNull": false
          },
          29,
          {
            "_ObjectIdentity_": "aa58909e-60c1-0000-29c7-003b321d02d1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          },
          31,
          {
            "IsNull": false
          },
          33,
          {
            "IsNull": false
          },
          34,
          {
            "_ObjectIdentity_": "aa58909e-60c1-0000-29c7-003b321d02d1|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUQElpjbqF1pFvtTv+GIkLe8="
          },
          35,
          {
            "_ObjectType_": "SP.Taxonomy.TermGroup",
            "_ObjectIdentity_": "aa58909e-60c1-0000-29c7-003b321d02d1|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUQElpjbqF1pFvtTv+GIkLe8=",
            "CreatedDate": "\/Date(1529479401033)\/",
            "Id": "\/Guid(36a62501-17ea-455a-bed4-eff862242def)\/",
            "LastModifiedDate": "\/Date(1529479401033)\/",
            "Name": "People",
            "Description": "",
            "IsSiteCollectionGroup": false,
            "IsSystemGroup": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, id: '36a62501-17ea-455a-bed4-eff862242def' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "CreatedDate": "2018-06-20T07:23:21.033Z",
          "Id": "36a62501-17ea-455a-bed4-eff862242def",
          "LastModifiedDate": "2018-06-20T07:23:21.033Z",
          "Name": "People",
          "Description": "",
          "IsSiteCollectionGroup": false,
          "IsSystemGroup": false
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets taxonomy term group by name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="25" ObjectPathId="24" /><ObjectIdentityQuery Id="26" ObjectPathId="24" /><ObjectPath Id="28" ObjectPathId="27" /><ObjectIdentityQuery Id="29" ObjectPathId="27" /><ObjectPath Id="31" ObjectPathId="30" /><ObjectPath Id="33" ObjectPathId="32" /><ObjectIdentityQuery Id="34" ObjectPathId="32" /><Query Id="35" ObjectPathId="32"><Query SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticMethod Id="24" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="27" ParentId="24" Name="GetDefaultSiteCollectionTermStore" /><Property Id="30" ParentId="27" Name="Groups" /><Method Id="32" ParentId="30" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.8105.1217",
            "ErrorInfo": null,
            "TraceCorrelationId": "aa58909e-60c1-0000-29c7-003b321d02d1"
          },
          25,
          {
            "IsNull": false
          },
          26,
          {
            "_ObjectIdentity_": "aa58909e-60c1-0000-29c7-003b321d02d1|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          },
          28,
          {
            "IsNull": false
          },
          29,
          {
            "_ObjectIdentity_": "aa58909e-60c1-0000-29c7-003b321d02d1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          },
          31,
          {
            "IsNull": false
          },
          33,
          {
            "IsNull": false
          },
          34,
          {
            "_ObjectIdentity_": "aa58909e-60c1-0000-29c7-003b321d02d1|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUQElpjbqF1pFvtTv+GIkLe8="
          },
          35,
          {
            "_ObjectType_": "SP.Taxonomy.TermGroup",
            "_ObjectIdentity_": "aa58909e-60c1-0000-29c7-003b321d02d1|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUQElpjbqF1pFvtTv+GIkLe8=",
            "CreatedDate": "\/Date(1529479401033)\/",
            "Id": "\/Guid(36a62501-17ea-455a-bed4-eff862242def)\/",
            "LastModifiedDate": "\/Date(1529479401033)\/",
            "Name": "People",
            "Description": "",
            "IsSiteCollectionGroup": false,
            "IsSystemGroup": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, name: 'People' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "CreatedDate": "2018-06-20T07:23:21.033Z",
          "Id": "36a62501-17ea-455a-bed4-eff862242def",
          "LastModifiedDate": "2018-06-20T07:23:21.033Z",
          "Name": "People",
          "Description": "",
          "IsSiteCollectionGroup": false,
          "IsSystemGroup": false
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles term group not found via id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="25" ObjectPathId="24" /><ObjectIdentityQuery Id="26" ObjectPathId="24" /><ObjectPath Id="28" ObjectPathId="27" /><ObjectIdentityQuery Id="29" ObjectPathId="27" /><ObjectPath Id="31" ObjectPathId="30" /><ObjectPath Id="33" ObjectPathId="32" /><ObjectIdentityQuery Id="34" ObjectPathId="32" /><Query Id="35" ObjectPathId="32"><Query SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticMethod Id="24" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="27" ParentId="24" Name="GetDefaultSiteCollectionTermStore" /><Property Id="30" ParentId="27" Name="Groups" /><Method Id="32" ParentId="30" Name="GetById"><Parameters><Parameter Type="Guid">{36a62501-17ea-455a-bed4-eff862242def}</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8105.1217", "ErrorInfo": {
              "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "3105909e-e037-0000-29c7-078ce31cbc78", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
            }, "TraceCorrelationId": "3105909e-e037-0000-29c7-078ce31cbc78"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, id: '36a62501-17ea-455a-bed4-eff862242def' } }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles term group not found via name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="25" ObjectPathId="24" /><ObjectIdentityQuery Id="26" ObjectPathId="24" /><ObjectPath Id="28" ObjectPathId="27" /><ObjectIdentityQuery Id="29" ObjectPathId="27" /><ObjectPath Id="31" ObjectPathId="30" /><ObjectPath Id="33" ObjectPathId="32" /><ObjectIdentityQuery Id="34" ObjectPathId="32" /><Query Id="35" ObjectPathId="32"><Query SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticMethod Id="24" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="27" ParentId="24" Name="GetDefaultSiteCollectionTermStore" /><Property Id="30" ParentId="27" Name="Groups" /><Method Id="32" ParentId="30" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8105.1217", "ErrorInfo": {
              "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "3105909e-e037-0000-29c7-078ce31cbc78", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
            }, "TraceCorrelationId": "3105909e-e037-0000-29c7-078ce31cbc78"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'People' } }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving taxonomy term groups', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
            "ErrorMessage": "File Not Found.", "ErrorValue": null, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26", "ErrorCode": -2147024894, "ErrorTypeName": "System.IO.FileNotFoundException"
          }, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26"
        }
      ]));
    });
    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('File Not Found.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if neither id nor name specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and name specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9e54299e-208a-4000-8546-cc4139091b26', name: 'People' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9e54299e-208a-4000-8546-cc4139091b26' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when name specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'People' } });
    assert.strictEqual(actual, true);
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

  it('handles promise rejection', (done) => {
    Utils.restore((command as any).getRequestDigest);
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.reject('getRequestDigest error'));
    
    cmdInstance.action({
      options: { debug: false, id: '36a62501-17ea-455a-bed4-eff862242def' }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('getRequestDigest error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});