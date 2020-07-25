import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./term-get');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.TERM_GET, () => {
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
    assert.strictEqual(command.name.startsWith(commands.TERM_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets taxonomy term by id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="6" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="7" ParentId="6" Name="GetDefaultSiteCollectionTermStore" /><Method Id="13" ParentId="7" Name="GetTerm"><Parameters><Parameter Type="Guid">{16573ae2-0cc4-42fa-a2ff-8bf0407bd385}</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8203.1211", "ErrorInfo": null, "TraceCorrelationId": "1d6f979e-2031-0000-37ae-18b2ccd62d87" }, 14, { "IsNull": false }, 15, { "_ObjectIdentity_": "1d6f979e-2031-0000-37ae-18b2ccd62d87|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv4jpXFsQM+kKi/4vwQHvThQ==" }, 16, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "1d6f979e-2031-0000-37ae-18b2ccd62d87|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv4jpXFsQM+kKi/4vwQHvThQ==", "CreatedDate": "/Date(1534946707600)/", "Id": "/Guid(16573ae2-0cc4-42fa-a2ff-8bf0407bd385)/", "LastModifiedDate": "/Date(1534946707600)/", "Name": "Engineering", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "DProdMGD104\\_SPOFrm_187262", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "Engineering", "TermsCount": 0 }]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '16573ae2-0cc4-42fa-a2ff-8bf0407bd385' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-08-22T14:05:07.600Z", "Id": "16573ae2-0cc4-42fa-a2ff-8bf0407bd385", "LastModifiedDate": "2018-08-22T14:05:07.600Z", "Name": "Engineering", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "DProdMGD104\\_SPOFrm_187262", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "Engineering", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets taxonomy term by name, term group by id, term set by id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="91" ObjectPathId="90" /><ObjectIdentityQuery Id="92" ObjectPathId="90" /><ObjectPath Id="94" ObjectPathId="93" /><ObjectIdentityQuery Id="95" ObjectPathId="93" /><ObjectPath Id="97" ObjectPathId="96" /><ObjectPath Id="99" ObjectPathId="98" /><ObjectIdentityQuery Id="100" ObjectPathId="98" /><ObjectPath Id="102" ObjectPathId="101" /><ObjectPath Id="104" ObjectPathId="103" /><ObjectIdentityQuery Id="105" ObjectPathId="103" /><ObjectPath Id="107" ObjectPathId="106" /><ObjectPath Id="109" ObjectPathId="108" /><ObjectIdentityQuery Id="110" ObjectPathId="108" /><Query Id="111" ObjectPathId="108"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="90" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="93" ParentId="90" Name="GetDefaultSiteCollectionTermStore" /><Property Id="96" ParentId="93" Name="Groups" /><Method Id="98" ParentId="96" Name="GetById"><Parameters><Parameter Type="Guid">{5c928151-c140-4d48-aab9-54da901c7fef}</Parameter></Parameters></Method><Property Id="101" ParentId="98" Name="TermSets" /><Method Id="103" ParentId="101" Name="GetById"><Parameters><Parameter Type="Guid">{8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f}</Parameter></Parameters></Method><Property Id="106" ParentId="103" Name="Terms" /><Method Id="108" ParentId="106" Name="GetByName"><Parameters><Parameter Type="String">IT</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8203.1211", "ErrorInfo": null, "TraceCorrelationId": "4cb8979e-f0e5-0000-37ae-196c257d7d8f" }, 91, { "IsNull": false }, 92, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 94, { "IsNull": false }, 95, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 97, { "IsNull": false }, 99, { "IsNull": false }, 100, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 102, { "IsNull": false }, 104, { "IsNull": false }, 105, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 107, { "IsNull": false }, 109, { "IsNull": false }, 110, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pvaDrOATi/9UueqPwTsTjfjw==" }, 111, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pvaDrOATi/9UueqPwTsTjfjw==", "CreatedDate": "/Date(1534946407973)/", "Id": "/Guid(01ce3a68-bf38-4bf5-9ea8-fc13b138df8f)/", "LastModifiedDate": "/Date(1534946407973)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "DProdMGD104\\_SPOFrm_187262", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, name: 'IT', termGroupId: '5c928151-c140-4d48-aab9-54da901c7fef', termSetId: '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-08-22T14:00:07.973Z", "Id": "01ce3a68-bf38-4bf5-9ea8-fc13b138df8f", "LastModifiedDate": "2018-08-22T14:00:07.973Z", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "DProdMGD104\\_SPOFrm_187262", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets taxonomy term by name, term group by id, term set by name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="91" ObjectPathId="90" /><ObjectIdentityQuery Id="92" ObjectPathId="90" /><ObjectPath Id="94" ObjectPathId="93" /><ObjectIdentityQuery Id="95" ObjectPathId="93" /><ObjectPath Id="97" ObjectPathId="96" /><ObjectPath Id="99" ObjectPathId="98" /><ObjectIdentityQuery Id="100" ObjectPathId="98" /><ObjectPath Id="102" ObjectPathId="101" /><ObjectPath Id="104" ObjectPathId="103" /><ObjectIdentityQuery Id="105" ObjectPathId="103" /><ObjectPath Id="107" ObjectPathId="106" /><ObjectPath Id="109" ObjectPathId="108" /><ObjectIdentityQuery Id="110" ObjectPathId="108" /><Query Id="111" ObjectPathId="108"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="90" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="93" ParentId="90" Name="GetDefaultSiteCollectionTermStore" /><Property Id="96" ParentId="93" Name="Groups" /><Method Id="98" ParentId="96" Name="GetById"><Parameters><Parameter Type="Guid">{5c928151-c140-4d48-aab9-54da901c7fef}</Parameter></Parameters></Method><Property Id="101" ParentId="98" Name="TermSets" /><Method Id="103" ParentId="101" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Property Id="106" ParentId="103" Name="Terms" /><Method Id="108" ParentId="106" Name="GetByName"><Parameters><Parameter Type="String">IT</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8203.1211", "ErrorInfo": null, "TraceCorrelationId": "4cb8979e-f0e5-0000-37ae-196c257d7d8f" }, 91, { "IsNull": false }, 92, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 94, { "IsNull": false }, 95, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 97, { "IsNull": false }, 99, { "IsNull": false }, 100, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 102, { "IsNull": false }, 104, { "IsNull": false }, 105, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 107, { "IsNull": false }, 109, { "IsNull": false }, 110, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pvaDrOATi/9UueqPwTsTjfjw==" }, 111, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pvaDrOATi/9UueqPwTsTjfjw==", "CreatedDate": "/Date(1534946407973)/", "Id": "/Guid(01ce3a68-bf38-4bf5-9ea8-fc13b138df8f)/", "LastModifiedDate": "/Date(1534946407973)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "DProdMGD104\\_SPOFrm_187262", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', termGroupId: '5c928151-c140-4d48-aab9-54da901c7fef', termSetName: 'Department' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-08-22T14:00:07.973Z", "Id": "01ce3a68-bf38-4bf5-9ea8-fc13b138df8f", "LastModifiedDate": "2018-08-22T14:00:07.973Z", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "DProdMGD104\\_SPOFrm_187262", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets taxonomy term by name, term group by name, term set by id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="91" ObjectPathId="90" /><ObjectIdentityQuery Id="92" ObjectPathId="90" /><ObjectPath Id="94" ObjectPathId="93" /><ObjectIdentityQuery Id="95" ObjectPathId="93" /><ObjectPath Id="97" ObjectPathId="96" /><ObjectPath Id="99" ObjectPathId="98" /><ObjectIdentityQuery Id="100" ObjectPathId="98" /><ObjectPath Id="102" ObjectPathId="101" /><ObjectPath Id="104" ObjectPathId="103" /><ObjectIdentityQuery Id="105" ObjectPathId="103" /><ObjectPath Id="107" ObjectPathId="106" /><ObjectPath Id="109" ObjectPathId="108" /><ObjectIdentityQuery Id="110" ObjectPathId="108" /><Query Id="111" ObjectPathId="108"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="90" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="93" ParentId="90" Name="GetDefaultSiteCollectionTermStore" /><Property Id="96" ParentId="93" Name="Groups" /><Method Id="98" ParentId="96" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="101" ParentId="98" Name="TermSets" /><Method Id="103" ParentId="101" Name="GetById"><Parameters><Parameter Type="Guid">{8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f}</Parameter></Parameters></Method><Property Id="106" ParentId="103" Name="Terms" /><Method Id="108" ParentId="106" Name="GetByName"><Parameters><Parameter Type="String">IT</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8203.1211", "ErrorInfo": null, "TraceCorrelationId": "4cb8979e-f0e5-0000-37ae-196c257d7d8f" }, 91, { "IsNull": false }, 92, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 94, { "IsNull": false }, 95, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 97, { "IsNull": false }, 99, { "IsNull": false }, 100, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 102, { "IsNull": false }, 104, { "IsNull": false }, 105, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 107, { "IsNull": false }, 109, { "IsNull": false }, 110, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pvaDrOATi/9UueqPwTsTjfjw==" }, 111, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pvaDrOATi/9UueqPwTsTjfjw==", "CreatedDate": "/Date(1534946407973)/", "Id": "/Guid(01ce3a68-bf38-4bf5-9ea8-fc13b138df8f)/", "LastModifiedDate": "/Date(1534946407973)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "DProdMGD104\\_SPOFrm_187262", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', termGroupName: 'People', termSetId: '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-08-22T14:00:07.973Z", "Id": "01ce3a68-bf38-4bf5-9ea8-fc13b138df8f", "LastModifiedDate": "2018-08-22T14:00:07.973Z", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "DProdMGD104\\_SPOFrm_187262", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets taxonomy term by name, term group by name, term set by name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="91" ObjectPathId="90" /><ObjectIdentityQuery Id="92" ObjectPathId="90" /><ObjectPath Id="94" ObjectPathId="93" /><ObjectIdentityQuery Id="95" ObjectPathId="93" /><ObjectPath Id="97" ObjectPathId="96" /><ObjectPath Id="99" ObjectPathId="98" /><ObjectIdentityQuery Id="100" ObjectPathId="98" /><ObjectPath Id="102" ObjectPathId="101" /><ObjectPath Id="104" ObjectPathId="103" /><ObjectIdentityQuery Id="105" ObjectPathId="103" /><ObjectPath Id="107" ObjectPathId="106" /><ObjectPath Id="109" ObjectPathId="108" /><ObjectIdentityQuery Id="110" ObjectPathId="108" /><Query Id="111" ObjectPathId="108"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="90" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="93" ParentId="90" Name="GetDefaultSiteCollectionTermStore" /><Property Id="96" ParentId="93" Name="Groups" /><Method Id="98" ParentId="96" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="101" ParentId="98" Name="TermSets" /><Method Id="103" ParentId="101" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Property Id="106" ParentId="103" Name="Terms" /><Method Id="108" ParentId="106" Name="GetByName"><Parameters><Parameter Type="String">IT</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8203.1211", "ErrorInfo": null, "TraceCorrelationId": "4cb8979e-f0e5-0000-37ae-196c257d7d8f" }, 91, { "IsNull": false }, 92, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 94, { "IsNull": false }, 95, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 97, { "IsNull": false }, 99, { "IsNull": false }, 100, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 102, { "IsNull": false }, 104, { "IsNull": false }, 105, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 107, { "IsNull": false }, 109, { "IsNull": false }, 110, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pvaDrOATi/9UueqPwTsTjfjw==" }, 111, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pvaDrOATi/9UueqPwTsTjfjw==", "CreatedDate": "/Date(1534946407973)/", "Id": "/Guid(01ce3a68-bf38-4bf5-9ea8-fc13b138df8f)/", "LastModifiedDate": "/Date(1534946407973)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "DProdMGD104\\_SPOFrm_187262", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', termGroupName: 'People', termSetName: 'Department' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-08-22T14:00:07.973Z", "Id": "01ce3a68-bf38-4bf5-9ea8-fc13b138df8f", "LastModifiedDate": "2018-08-22T14:00:07.973Z", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "DProdMGD104\\_SPOFrm_187262", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles term not found by name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="91" ObjectPathId="90" /><ObjectIdentityQuery Id="92" ObjectPathId="90" /><ObjectPath Id="94" ObjectPathId="93" /><ObjectIdentityQuery Id="95" ObjectPathId="93" /><ObjectPath Id="97" ObjectPathId="96" /><ObjectPath Id="99" ObjectPathId="98" /><ObjectIdentityQuery Id="100" ObjectPathId="98" /><ObjectPath Id="102" ObjectPathId="101" /><ObjectPath Id="104" ObjectPathId="103" /><ObjectIdentityQuery Id="105" ObjectPathId="103" /><ObjectPath Id="107" ObjectPathId="106" /><ObjectPath Id="109" ObjectPathId="108" /><ObjectIdentityQuery Id="110" ObjectPathId="108" /><Query Id="111" ObjectPathId="108"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="90" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="93" ParentId="90" Name="GetDefaultSiteCollectionTermStore" /><Property Id="96" ParentId="93" Name="Groups" /><Method Id="98" ParentId="96" Name="GetById"><Parameters><Parameter Type="Guid">{5c928151-c140-4d48-aab9-54da901c7fef}</Parameter></Parameters></Method><Property Id="101" ParentId="98" Name="TermSets" /><Method Id="103" ParentId="101" Name="GetById"><Parameters><Parameter Type="Guid">{8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f}</Parameter></Parameters></Method><Property Id="106" ParentId="103" Name="Terms" /><Method Id="108" ParentId="106" Name="GetByName"><Parameters><Parameter Type="String">IT</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8203.1211", "ErrorInfo": { "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "e2b8979e-307e-0000-37ae-164629627460", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException" }, "TraceCorrelationId": "e2b8979e-307e-0000-37ae-164629627460" }]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', termGroupId: '5c928151-c140-4d48-aab9-54da901c7fef', termSetId: '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles term not found by id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="6" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="7" ParentId="6" Name="GetDefaultSiteCollectionTermStore" /><Method Id="13" ParentId="7" Name="GetTerm"><Parameters><Parameter Type="Guid">{16573ae2-0cc4-42fa-a2ff-8bf0407bd385}</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8203.1211", "ErrorInfo": null, "TraceCorrelationId": "1bb9979e-a09b-0000-37d0-4314e2f545e4" }, 14, { "IsNull": true }, 15, null, 16, null]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '16573ae2-0cc4-42fa-a2ff-8bf0407bd385' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles term group not found', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="91" ObjectPathId="90" /><ObjectIdentityQuery Id="92" ObjectPathId="90" /><ObjectPath Id="94" ObjectPathId="93" /><ObjectIdentityQuery Id="95" ObjectPathId="93" /><ObjectPath Id="97" ObjectPathId="96" /><ObjectPath Id="99" ObjectPathId="98" /><ObjectIdentityQuery Id="100" ObjectPathId="98" /><ObjectPath Id="102" ObjectPathId="101" /><ObjectPath Id="104" ObjectPathId="103" /><ObjectIdentityQuery Id="105" ObjectPathId="103" /><ObjectPath Id="107" ObjectPathId="106" /><ObjectPath Id="109" ObjectPathId="108" /><ObjectIdentityQuery Id="110" ObjectPathId="108" /><Query Id="111" ObjectPathId="108"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="90" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="93" ParentId="90" Name="GetDefaultSiteCollectionTermStore" /><Property Id="96" ParentId="93" Name="Groups" /><Method Id="98" ParentId="96" Name="GetById"><Parameters><Parameter Type="Guid">{5c928151-c140-4d48-aab9-54da901c7fef}</Parameter></Parameters></Method><Property Id="101" ParentId="98" Name="TermSets" /><Method Id="103" ParentId="101" Name="GetById"><Parameters><Parameter Type="Guid">{8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f}</Parameter></Parameters></Method><Property Id="106" ParentId="103" Name="Terms" /><Method Id="108" ParentId="106" Name="GetByName"><Parameters><Parameter Type="String">IT</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8203.1211", "ErrorInfo": { "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "2eb9979e-802f-0000-37ae-13af10dc354b", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException" }, "TraceCorrelationId": "2eb9979e-802f-0000-37ae-13af10dc354b" }]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', termGroupId: '5c928151-c140-4d48-aab9-54da901c7fef', termSetId: '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles term set not found', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="91" ObjectPathId="90" /><ObjectIdentityQuery Id="92" ObjectPathId="90" /><ObjectPath Id="94" ObjectPathId="93" /><ObjectIdentityQuery Id="95" ObjectPathId="93" /><ObjectPath Id="97" ObjectPathId="96" /><ObjectPath Id="99" ObjectPathId="98" /><ObjectIdentityQuery Id="100" ObjectPathId="98" /><ObjectPath Id="102" ObjectPathId="101" /><ObjectPath Id="104" ObjectPathId="103" /><ObjectIdentityQuery Id="105" ObjectPathId="103" /><ObjectPath Id="107" ObjectPathId="106" /><ObjectPath Id="109" ObjectPathId="108" /><ObjectIdentityQuery Id="110" ObjectPathId="108" /><Query Id="111" ObjectPathId="108"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="90" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="93" ParentId="90" Name="GetDefaultSiteCollectionTermStore" /><Property Id="96" ParentId="93" Name="Groups" /><Method Id="98" ParentId="96" Name="GetById"><Parameters><Parameter Type="Guid">{5c928151-c140-4d48-aab9-54da901c7fef}</Parameter></Parameters></Method><Property Id="101" ParentId="98" Name="TermSets" /><Method Id="103" ParentId="101" Name="GetById"><Parameters><Parameter Type="Guid">{8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f}</Parameter></Parameters></Method><Property Id="106" ParentId="103" Name="Terms" /><Method Id="108" ParentId="106" Name="GetByName"><Parameters><Parameter Type="String">IT</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8203.1211", "ErrorInfo": { "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "45b9979e-00af-0000-37ae-1098a489e222", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException" }, "TraceCorrelationId": "45b9979e-00af-0000-37ae-1098a489e222" }]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', termGroupId: '5c928151-c140-4d48-aab9-54da901c7fef', termSetId: '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in term name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="91" ObjectPathId="90" /><ObjectIdentityQuery Id="92" ObjectPathId="90" /><ObjectPath Id="94" ObjectPathId="93" /><ObjectIdentityQuery Id="95" ObjectPathId="93" /><ObjectPath Id="97" ObjectPathId="96" /><ObjectPath Id="99" ObjectPathId="98" /><ObjectIdentityQuery Id="100" ObjectPathId="98" /><ObjectPath Id="102" ObjectPathId="101" /><ObjectPath Id="104" ObjectPathId="103" /><ObjectIdentityQuery Id="105" ObjectPathId="103" /><ObjectPath Id="107" ObjectPathId="106" /><ObjectPath Id="109" ObjectPathId="108" /><ObjectIdentityQuery Id="110" ObjectPathId="108" /><Query Id="111" ObjectPathId="108"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="90" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="93" ParentId="90" Name="GetDefaultSiteCollectionTermStore" /><Property Id="96" ParentId="93" Name="Groups" /><Method Id="98" ParentId="96" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="101" ParentId="98" Name="TermSets" /><Method Id="103" ParentId="101" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Property Id="106" ParentId="103" Name="Terms" /><Method Id="108" ParentId="106" Name="GetByName"><Parameters><Parameter Type="String">IT&gt;</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8203.1211", "ErrorInfo": null, "TraceCorrelationId": "4cb8979e-f0e5-0000-37ae-196c257d7d8f" }, 91, { "IsNull": false }, 92, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 94, { "IsNull": false }, 95, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 97, { "IsNull": false }, 99, { "IsNull": false }, 100, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 102, { "IsNull": false }, 104, { "IsNull": false }, 105, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 107, { "IsNull": false }, 109, { "IsNull": false }, 110, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pvaDrOATi/9UueqPwTsTjfjw==" }, 111, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pvaDrOATi/9UueqPwTsTjfjw==", "CreatedDate": "/Date(1534946407973)/", "Id": "/Guid(01ce3a68-bf38-4bf5-9ea8-fc13b138df8f)/", "LastModifiedDate": "/Date(1534946407973)/", "Name": "IT>", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "DProdMGD104\\_SPOFrm_187262", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT>", "TermsCount": 0 }]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT>', termGroupName: 'People', termSetName: 'Department' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-08-22T14:00:07.973Z", "Id": "01ce3a68-bf38-4bf5-9ea8-fc13b138df8f", "LastModifiedDate": "2018-08-22T14:00:07.973Z", "Name": "IT>", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "DProdMGD104\\_SPOFrm_187262", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT>", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in term group name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="91" ObjectPathId="90" /><ObjectIdentityQuery Id="92" ObjectPathId="90" /><ObjectPath Id="94" ObjectPathId="93" /><ObjectIdentityQuery Id="95" ObjectPathId="93" /><ObjectPath Id="97" ObjectPathId="96" /><ObjectPath Id="99" ObjectPathId="98" /><ObjectIdentityQuery Id="100" ObjectPathId="98" /><ObjectPath Id="102" ObjectPathId="101" /><ObjectPath Id="104" ObjectPathId="103" /><ObjectIdentityQuery Id="105" ObjectPathId="103" /><ObjectPath Id="107" ObjectPathId="106" /><ObjectPath Id="109" ObjectPathId="108" /><ObjectIdentityQuery Id="110" ObjectPathId="108" /><Query Id="111" ObjectPathId="108"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="90" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="93" ParentId="90" Name="GetDefaultSiteCollectionTermStore" /><Property Id="96" ParentId="93" Name="Groups" /><Method Id="98" ParentId="96" Name="GetByName"><Parameters><Parameter Type="String">People&gt;</Parameter></Parameters></Method><Property Id="101" ParentId="98" Name="TermSets" /><Method Id="103" ParentId="101" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Property Id="106" ParentId="103" Name="Terms" /><Method Id="108" ParentId="106" Name="GetByName"><Parameters><Parameter Type="String">IT</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8203.1211", "ErrorInfo": null, "TraceCorrelationId": "4cb8979e-f0e5-0000-37ae-196c257d7d8f" }, 91, { "IsNull": false }, 92, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 94, { "IsNull": false }, 95, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 97, { "IsNull": false }, 99, { "IsNull": false }, 100, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 102, { "IsNull": false }, 104, { "IsNull": false }, 105, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 107, { "IsNull": false }, 109, { "IsNull": false }, 110, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pvaDrOATi/9UueqPwTsTjfjw==" }, 111, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pvaDrOATi/9UueqPwTsTjfjw==", "CreatedDate": "/Date(1534946407973)/", "Id": "/Guid(01ce3a68-bf38-4bf5-9ea8-fc13b138df8f)/", "LastModifiedDate": "/Date(1534946407973)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "DProdMGD104\\_SPOFrm_187262", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', termGroupName: 'People>', termSetName: 'Department' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-08-22T14:00:07.973Z", "Id": "01ce3a68-bf38-4bf5-9ea8-fc13b138df8f", "LastModifiedDate": "2018-08-22T14:00:07.973Z", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "DProdMGD104\\_SPOFrm_187262", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in term set name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="91" ObjectPathId="90" /><ObjectIdentityQuery Id="92" ObjectPathId="90" /><ObjectPath Id="94" ObjectPathId="93" /><ObjectIdentityQuery Id="95" ObjectPathId="93" /><ObjectPath Id="97" ObjectPathId="96" /><ObjectPath Id="99" ObjectPathId="98" /><ObjectIdentityQuery Id="100" ObjectPathId="98" /><ObjectPath Id="102" ObjectPathId="101" /><ObjectPath Id="104" ObjectPathId="103" /><ObjectIdentityQuery Id="105" ObjectPathId="103" /><ObjectPath Id="107" ObjectPathId="106" /><ObjectPath Id="109" ObjectPathId="108" /><ObjectIdentityQuery Id="110" ObjectPathId="108" /><Query Id="111" ObjectPathId="108"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="90" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="93" ParentId="90" Name="GetDefaultSiteCollectionTermStore" /><Property Id="96" ParentId="93" Name="Groups" /><Method Id="98" ParentId="96" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="101" ParentId="98" Name="TermSets" /><Method Id="103" ParentId="101" Name="GetByName"><Parameters><Parameter Type="String">Department&gt;</Parameter></Parameters></Method><Property Id="106" ParentId="103" Name="Terms" /><Method Id="108" ParentId="106" Name="GetByName"><Parameters><Parameter Type="String">IT</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8203.1211", "ErrorInfo": null, "TraceCorrelationId": "4cb8979e-f0e5-0000-37ae-196c257d7d8f" }, 91, { "IsNull": false }, 92, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 94, { "IsNull": false }, 95, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 97, { "IsNull": false }, 99, { "IsNull": false }, 100, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 102, { "IsNull": false }, 104, { "IsNull": false }, 105, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 107, { "IsNull": false }, 109, { "IsNull": false }, 110, { "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pvaDrOATi/9UueqPwTsTjfjw==" }, 111, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "4cb8979e-f0e5-0000-37ae-196c257d7d8f|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pvaDrOATi/9UueqPwTsTjfjw==", "CreatedDate": "/Date(1534946407973)/", "Id": "/Guid(01ce3a68-bf38-4bf5-9ea8-fc13b138df8f)/", "LastModifiedDate": "/Date(1534946407973)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "DProdMGD104\\_SPOFrm_187262", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', termGroupName: 'People', termSetName: 'Department>' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-08-22T14:00:07.973Z", "Id": "01ce3a68-bf38-4bf5-9ea8-fc13b138df8f", "LastModifiedDate": "2018-08-22T14:00:07.973Z", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "DProdMGD104\\_SPOFrm_187262", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if neither id nor name specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { termGroupName: 'People', termSetName: 'Department' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and name specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9e54299e-208a-4000-8546-cc4139091b26', name: 'IT', termGroupName: 'People', termSetName: 'Department' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if only id specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '9e54299e-208a-4000-8546-cc4139091b26' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation when only name specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when only name and termGroupId specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupId: '9e54299e-208a-4000-8546-cc4139091b26' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both termGroupId and termGroupName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupId: '9e54299e-208a-4000-8546-cc4139091b27', termGroupName: 'People', termSetName: 'Department' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if termGroupId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupId: 'invalid', termSetName: 'Department' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both termSetId and termSetName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupId: '9e54299e-208a-4000-8546-cc4139091b27', termSetName: 'Department', termSetId: '9e54299e-208a-4000-8546-cc4139091b2a' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if termSetId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termSetId: 'invalid', termGroupName: 'People' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name, termGroupName and termSetName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupName: 'People', termSetName: 'Department' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when name, termGroupId and termSetId specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupId: '9e54299e-208a-4000-8546-cc4139091b27', termSetId: '9e54299e-208a-4000-8546-cc4139091b2a' } });
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
      options: { debug: false, name: 'IT', termGroupName: 'People', termSetName: 'Department>' }
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