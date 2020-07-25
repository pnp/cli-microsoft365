import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./term-add');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.TERM_ADD, () => {
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
    assert.strictEqual(command.name.startsWith(commands.TERM_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds term with the specified name to the term set and term group specified by name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1205", "ErrorInfo": null, "TraceCorrelationId": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6" }, 4, { "IsNull": false }, 5, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 7, { "IsNull": false }, 8, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 10, { "IsNull": false }, 12, { "IsNull": false }, 13, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 15, { "IsNull": false }, 17, { "IsNull": false }, 18, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 20, { "IsNull": false }, 21, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" }, 22, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==", "CreatedDate": "/Date(1540235503669)/", "Id": "/Guid(47fdacfe-ff64-4a05-b611-e84e767f04de)/", "LastModifiedDate": "/Date(1540235503669)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'IT', termSetName: 'Department', termGroupName: 'People' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-10-22T19:11:43.669Z", "Id": "47fdacfe-ff64-4a05-b611-e84e767f04de", "LastModifiedDate": "2018-10-22T19:11:43.669Z", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds term with the specified name and id to the term set and term group specified by id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetById"><Parameters><Parameter Type="Guid">{5c928151-c140-4d48-aab9-54da901c7fef}</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetById"><Parameters><Parameter Type="Guid">{8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f}</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{47fdacfe-ff64-4a05-b611-e84e767f04de}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1205", "ErrorInfo": null, "TraceCorrelationId": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6" }, 4, { "IsNull": false }, 5, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 7, { "IsNull": false }, 8, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 10, { "IsNull": false }, 12, { "IsNull": false }, 13, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 15, { "IsNull": false }, 17, { "IsNull": false }, 18, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 20, { "IsNull": false }, 21, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" }, 22, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==", "CreatedDate": "/Date(1540235503669)/", "Id": "/Guid(47fdacfe-ff64-4a05-b611-e84e767f04de)/", "LastModifiedDate": "/Date(1540235503669)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', id: '47fdacfe-ff64-4a05-b611-e84e767f04de', termSetId: '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f', termGroupId: '5c928151-c140-4d48-aab9-54da901c7fef' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-10-22T19:11:43.669Z", "Id": "47fdacfe-ff64-4a05-b611-e84e767f04de", "LastModifiedDate": "2018-10-22T19:11:43.669Z", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds term with the specified name and id below the specified term', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetById"><Parameters><Parameter Type="Guid">{5c928151-c140-4d48-aab9-54da901c7fef}</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="6" Name="GetTerm"><Parameters><Parameter Type="Guid">{8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f}</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{47fdacfe-ff64-4a05-b611-e84e767f04de}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1205", "ErrorInfo": null, "TraceCorrelationId": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6" }, 4, { "IsNull": false }, 5, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 7, { "IsNull": false }, 8, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 10, { "IsNull": false }, 12, { "IsNull": false }, 13, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 15, { "IsNull": false }, 17, { "IsNull": false }, 18, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 20, { "IsNull": false }, 21, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" }, 22, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==", "CreatedDate": "/Date(1540235503669)/", "Id": "/Guid(47fdacfe-ff64-4a05-b611-e84e767f04de)/", "LastModifiedDate": "/Date(1540235503669)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', id: '47fdacfe-ff64-4a05-b611-e84e767f04de', parentTermId: '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f', termGroupId: '5c928151-c140-4d48-aab9-54da901c7fef' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-10-22T19:11:43.669Z", "Id": "47fdacfe-ff64-4a05-b611-e84e767f04de", "LastModifiedDate": "2018-10-22T19:11:43.669Z", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds term with description', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1205", "ErrorInfo": null, "TraceCorrelationId": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6" }, 4, { "IsNull": false }, 5, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 7, { "IsNull": false }, 8, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 10, { "IsNull": false }, 12, { "IsNull": false }, 13, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 15, { "IsNull": false }, 17, { "IsNull": false }, 18, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 20, { "IsNull": false }, 21, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" }, 22, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==", "CreatedDate": "/Date(1540235503669)/", "Id": "/Guid(47fdacfe-ff64-4a05-b611-e84e767f04de)/", "LastModifiedDate": "/Date(1540235503669)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetDescription" Id="127" ObjectPathId="117"><Parameters><Parameter Type="String">IT term</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method><Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" /><Identity Id="109" Name="d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1221", "ErrorInfo": null, "TraceCorrelationId": "8b409b9e-b003-0000-37ae-1d4bfff0edf2"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, name: 'IT', description: 'IT term', termSetName: 'Department', termGroupName: 'People' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-10-22T19:11:43.669Z", "Id": "47fdacfe-ff64-4a05-b611-e84e767f04de", "LastModifiedDate": "2018-10-22T19:11:43.669Z", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "IT term", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds term with local and custom local properties', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1205", "ErrorInfo": null, "TraceCorrelationId": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6" }, 4, { "IsNull": false }, 5, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 7, { "IsNull": false }, 8, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 10, { "IsNull": false }, 12, { "IsNull": false }, 13, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 15, { "IsNull": false }, 17, { "IsNull": false }, 18, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 20, { "IsNull": false }, 21, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" }, 22, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==", "CreatedDate": "/Date(1540235503669)/", "Id": "/Guid(47fdacfe-ff64-4a05-b611-e84e767f04de)/", "LastModifiedDate": "/Date(1540235503669)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetCustomProperty" Id="127" ObjectPathId="117"><Parameters><Parameter Type="String">Prop1</Parameter><Parameter Type="String">Value1</Parameter></Parameters></Method><Method Name="SetLocalCustomProperty" Id="128" ObjectPathId="117"><Parameters><Parameter Type="String">LocalProp1</Parameter><Parameter Type="String">LocalValue1</Parameter></Parameters></Method><Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" /><Identity Id="109" Name="d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1221", "ErrorInfo": null, "TraceCorrelationId": "8b409b9e-b003-0000-37ae-1d4bfff0edf2"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', customProperties: '{"Prop1": "Value1"}', localCustomProperties: '{"LocalProp1": "LocalValue1"}', termSetName: 'Department', termGroupName: 'People' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-10-22T19:11:43.669Z", "Id": "47fdacfe-ff64-4a05-b611-e84e767f04de", "LastModifiedDate": "2018-10-22T19:11:43.669Z", "Name": "IT", "CustomProperties": { "Prop1": "Value1" }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": { "LocalProp1": "LocalValue1" }, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving the term store', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": {
                "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', termSetName: 'Department', termGroupName: 'People' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the term group specified by id doesn\'t exist', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetById"><Parameters><Parameter Type="Guid">{5c928151-c140-4d48-aab9-54da901c7fef}</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetById"><Parameters><Parameter Type="Guid">{8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f}</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{47fdacfe-ff64-4a05-b611-e84e767f04de}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8105.1217", "ErrorInfo": {
                "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "3105909e-e037-0000-29c7-078ce31cbc78", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
              }, "TraceCorrelationId": "3105909e-e037-0000-29c7-078ce31cbc78"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', id: '47fdacfe-ff64-4a05-b611-e84e767f04de', termSetId: '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f', termGroupId: '5c928151-c140-4d48-aab9-54da901c7fef' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the term group specified by name doesn\'t exist', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8105.1217", "ErrorInfo": {
                "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "3105909e-e037-0000-29c7-078ce31cbc78", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
              }, "TraceCorrelationId": "3105909e-e037-0000-29c7-078ce31cbc78"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', termSetName: 'Department', termGroupName: 'People' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the term set specified by name doesn\'t exist', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8105.1217", "ErrorInfo": {
                "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "3105909e-e037-0000-29c7-078ce31cbc78", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
              }, "TraceCorrelationId": "3105909e-e037-0000-29c7-078ce31cbc78"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', termSetName: 'Department', termGroupName: 'People' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the term set specified by id doesn\'t exist', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetById"><Parameters><Parameter Type="Guid">{5c928151-c140-4d48-aab9-54da901c7fef}</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetById"><Parameters><Parameter Type="Guid">{8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f}</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{47fdacfe-ff64-4a05-b611-e84e767f04de}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8105.1217", "ErrorInfo": {
                "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "3105909e-e037-0000-29c7-078ce31cbc78", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
              }, "TraceCorrelationId": "3105909e-e037-0000-29c7-078ce31cbc78"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', id: '47fdacfe-ff64-4a05-b611-e84e767f04de', termSetId: '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f', termGroupId: '5c928151-c140-4d48-aab9-54da901c7fef' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified name already exists', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1221", "ErrorInfo": { "ErrorMessage": "There is already a term with the same default label and parent term.", "ErrorValue": null, "TraceCorrelationId": "5c419b9e-5074-0000-3292-b5fe42f75fd1", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Taxonomy.TermStoreOperationException" }, "TraceCorrelationId": "5c419b9e-5074-0000-3292-b5fe42f75fd1" }]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', termSetName: 'Department', termGroupName: 'People' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('There is already a term with the same default label and parent term.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified id already exists', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetById"><Parameters><Parameter Type="Guid">{5c928151-c140-4d48-aab9-54da901c7fef}</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetById"><Parameters><Parameter Type="Guid">{8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f}</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{47fdacfe-ff64-4a05-b611-e84e767f04de}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1221", "ErrorInfo": { "ErrorMessage": "Failed to read from or write to database. Refresh and try again. If the problem persists, please contact the administrator.", "ErrorValue": null, "TraceCorrelationId": "8f419b9e-b042-0000-37ae-164c0c311c0a", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Taxonomy.TermStoreOperationException" }, "TraceCorrelationId": "8f419b9e-b042-0000-37ae-164c0c311c0a" }]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', id: '47fdacfe-ff64-4a05-b611-e84e767f04de', termSetId: '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f', termGroupId: '5c928151-c140-4d48-aab9-54da901c7fef' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Failed to read from or write to database. Refresh and try again. If the problem persists, please contact the administrator.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when setting the description', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1205", "ErrorInfo": null, "TraceCorrelationId": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6" }, 4, { "IsNull": false }, 5, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 7, { "IsNull": false }, 8, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 10, { "IsNull": false }, 12, { "IsNull": false }, 13, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 15, { "IsNull": false }, 17, { "IsNull": false }, 18, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 20, { "IsNull": false }, 21, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" }, 22, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==", "CreatedDate": "/Date(1540235503669)/", "Id": "/Guid(47fdacfe-ff64-4a05-b611-e84e767f04de)/", "LastModifiedDate": "/Date(1540235503669)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetDescription" Id="127" ObjectPathId="117"><Parameters><Parameter Type="String">IT term</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method><Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" /><Identity Id="109" Name="d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": {
                "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', description: 'IT term', termSetName: 'Department', termGroupName: 'People' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when setting custom properties', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1205", "ErrorInfo": null, "TraceCorrelationId": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6" }, 4, { "IsNull": false }, 5, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 7, { "IsNull": false }, 8, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 10, { "IsNull": false }, 12, { "IsNull": false }, 13, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 15, { "IsNull": false }, 17, { "IsNull": false }, 18, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 20, { "IsNull": false }, 21, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" }, 22, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==", "CreatedDate": "/Date(1540235503669)/", "Id": "/Guid(47fdacfe-ff64-4a05-b611-e84e767f04de)/", "LastModifiedDate": "/Date(1540235503669)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetCustomProperty" Id="127" ObjectPathId="117"><Parameters><Parameter Type="String">Prop1</Parameter><Parameter Type="String">Value1</Parameter></Parameters></Method><Method Name="SetLocalCustomProperty" Id="128" ObjectPathId="117"><Parameters><Parameter Type="String">LocalProp1</Parameter><Parameter Type="String">LocalValue1</Parameter></Parameters></Method><Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" /><Identity Id="109" Name="d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": {
                "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', customProperties: '{"Prop1": "Value1"}', localCustomProperties: '{"LocalProp1": "LocalValue1"}', termSetName: 'Department', termGroupName: 'People' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when setting local custom properties', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1205", "ErrorInfo": null, "TraceCorrelationId": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6" }, 4, { "IsNull": false }, 5, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 7, { "IsNull": false }, 8, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 10, { "IsNull": false }, 12, { "IsNull": false }, 13, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 15, { "IsNull": false }, 17, { "IsNull": false }, 18, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 20, { "IsNull": false }, 21, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" }, 22, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==", "CreatedDate": "/Date(1540235503669)/", "Id": "/Guid(47fdacfe-ff64-4a05-b611-e84e767f04de)/", "LastModifiedDate": "/Date(1540235503669)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetCustomProperty" Id="127" ObjectPathId="117"><Parameters><Parameter Type="String">Prop1</Parameter><Parameter Type="String">Value1</Parameter></Parameters></Method><Method Name="SetLocalCustomProperty" Id="128" ObjectPathId="117"><Parameters><Parameter Type="String">LocalProp1</Parameter><Parameter Type="String">LocalValue1</Parameter></Parameters></Method><Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" /><Identity Id="109" Name="d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": {
                "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', customProperties: '{"Prop1": "Value1"}', localCustomProperties: '{"LocalProp1": "LocalValue1"}', termSetName: 'Department', termGroupName: 'People' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly escapes XML in term group name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">People&gt;</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1205", "ErrorInfo": null, "TraceCorrelationId": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6" }, 4, { "IsNull": false }, 5, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 7, { "IsNull": false }, 8, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 10, { "IsNull": false }, 12, { "IsNull": false }, 13, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 15, { "IsNull": false }, 17, { "IsNull": false }, 18, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 20, { "IsNull": false }, 21, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" }, 22, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==", "CreatedDate": "/Date(1540235503669)/", "Id": "/Guid(47fdacfe-ff64-4a05-b611-e84e767f04de)/", "LastModifiedDate": "/Date(1540235503669)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', termSetName: 'Department', termGroupName: 'People>' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-10-22T19:11:43.669Z", "Id": "47fdacfe-ff64-4a05-b611-e84e767f04de", "LastModifiedDate": "2018-10-22T19:11:43.669Z", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly escapes XML in term set name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">Department&gt;</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1205", "ErrorInfo": null, "TraceCorrelationId": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6" }, 4, { "IsNull": false }, 5, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 7, { "IsNull": false }, 8, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 10, { "IsNull": false }, 12, { "IsNull": false }, 13, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 15, { "IsNull": false }, 17, { "IsNull": false }, 18, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 20, { "IsNull": false }, 21, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" }, 22, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==", "CreatedDate": "/Date(1540235503669)/", "Id": "/Guid(47fdacfe-ff64-4a05-b611-e84e767f04de)/", "LastModifiedDate": "/Date(1540235503669)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', termSetName: 'Department>', termGroupName: 'People' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-10-22T19:11:43.669Z", "Id": "47fdacfe-ff64-4a05-b611-e84e767f04de", "LastModifiedDate": "2018-10-22T19:11:43.669Z", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly escapes XML in term name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT&gt;</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1205", "ErrorInfo": null, "TraceCorrelationId": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6" }, 4, { "IsNull": false }, 5, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 7, { "IsNull": false }, 8, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 10, { "IsNull": false }, 12, { "IsNull": false }, 13, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 15, { "IsNull": false }, 17, { "IsNull": false }, 18, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 20, { "IsNull": false }, 21, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" }, 22, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==", "CreatedDate": "/Date(1540235503669)/", "Id": "/Guid(47fdacfe-ff64-4a05-b611-e84e767f04de)/", "LastModifiedDate": "/Date(1540235503669)/", "Name": "IT>", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT>', termSetName: 'Department', termGroupName: 'People' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-10-22T19:11:43.669Z", "Id": "47fdacfe-ff64-4a05-b611-e84e767f04de", "LastModifiedDate": "2018-10-22T19:11:43.669Z", "Name": "IT>", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly escapes XML in term description', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1205", "ErrorInfo": null, "TraceCorrelationId": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6" }, 4, { "IsNull": false }, 5, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 7, { "IsNull": false }, 8, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 10, { "IsNull": false }, 12, { "IsNull": false }, 13, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 15, { "IsNull": false }, 17, { "IsNull": false }, 18, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 20, { "IsNull": false }, 21, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" }, 22, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==", "CreatedDate": "/Date(1540235503669)/", "Id": "/Guid(47fdacfe-ff64-4a05-b611-e84e767f04de)/", "LastModifiedDate": "/Date(1540235503669)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetDescription" Id="127" ObjectPathId="117"><Parameters><Parameter Type="String">IT term&gt;</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method><Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" /><Identity Id="109" Name="d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1221", "ErrorInfo": null, "TraceCorrelationId": "8b409b9e-b003-0000-37ae-1d4bfff0edf2"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', description: 'IT term>', termSetName: 'Department', termGroupName: 'People' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-10-22T19:11:43.669Z", "Id": "47fdacfe-ff64-4a05-b611-e84e767f04de", "LastModifiedDate": "2018-10-22T19:11:43.669Z", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "IT term>", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly escapes XML in term custom properties', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1205", "ErrorInfo": null, "TraceCorrelationId": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6" }, 4, { "IsNull": false }, 5, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 7, { "IsNull": false }, 8, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 10, { "IsNull": false }, 12, { "IsNull": false }, 13, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 15, { "IsNull": false }, 17, { "IsNull": false }, 18, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 20, { "IsNull": false }, 21, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" }, 22, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==", "CreatedDate": "/Date(1540235503669)/", "Id": "/Guid(47fdacfe-ff64-4a05-b611-e84e767f04de)/", "LastModifiedDate": "/Date(1540235503669)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetCustomProperty" Id="127" ObjectPathId="117"><Parameters><Parameter Type="String">Prop1&gt;</Parameter><Parameter Type="String">Value1&gt;</Parameter></Parameters></Method><Method Name="SetLocalCustomProperty" Id="128" ObjectPathId="117"><Parameters><Parameter Type="String">LocalProp1</Parameter><Parameter Type="String">LocalValue1</Parameter></Parameters></Method><Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" /><Identity Id="109" Name="d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1221", "ErrorInfo": null, "TraceCorrelationId": "8b409b9e-b003-0000-37ae-1d4bfff0edf2"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', customProperties: '{"Prop1>": "Value1>"}', localCustomProperties: '{"LocalProp1": "LocalValue1"}', termSetName: 'Department', termGroupName: 'People' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-10-22T19:11:43.669Z", "Id": "47fdacfe-ff64-4a05-b611-e84e767f04de", "LastModifiedDate": "2018-10-22T19:11:43.669Z", "Name": "IT", "CustomProperties": { "Prop1>": "Value1>" }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": { "LocalProp1": "LocalValue1" }, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly escapes XML in term local custom properties', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /><Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">People</Parameter></Parameters></Method><Property Id="14" ParentId="11" Name="TermSets" /><Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">Department</Parameter></Parameters></Method><Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">IT</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1205", "ErrorInfo": null, "TraceCorrelationId": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6" }, 4, { "IsNull": false }, 5, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:ss:" }, 7, { "IsNull": false }, 8, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" }, 10, { "IsNull": false }, 12, { "IsNull": false }, 13, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:gr:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+8=" }, 15, { "IsNull": false }, 17, { "IsNull": false }, 18, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:se:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv" }, 20, { "IsNull": false }, 21, { "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" }, 22, { "_ObjectType_": "SP.Taxonomy.Term", "_ObjectIdentity_": "d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==", "CreatedDate": "/Date(1540235503669)/", "Id": "/Guid(47fdacfe-ff64-4a05-b611-e84e767f04de)/", "LastModifiedDate": "/Date(1540235503669)/", "Name": "IT", "CustomProperties": {}, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": {}, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetCustomProperty" Id="127" ObjectPathId="117"><Parameters><Parameter Type="String">Prop1</Parameter><Parameter Type="String">Value1</Parameter></Parameters></Method><Method Name="SetLocalCustomProperty" Id="128" ObjectPathId="117"><Parameters><Parameter Type="String">LocalProp1&gt;</Parameter><Parameter Type="String">LocalValue1&gt;</Parameter></Parameters></Method><Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:te:MvRe/3xHkEqrmEXxmJ7Lx1GBklxAwUhNqrlU2pAcf+/qydiOUnAdTKTXucEL/+pv/qz9R2T/BUq2EehOdn8E3g==" /><Identity Id="109" Name="d7f59a9e-a0f5-0000-37ae-17ef5f03c2e6|fec14c62-7c3b-481b-851b-c80d7802b224:st:MvRe/3xHkEqrmEXxmJ7Lxw==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8210.1221", "ErrorInfo": null, "TraceCorrelationId": "8b409b9e-b003-0000-37ae-1d4bfff0edf2"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, name: 'IT', customProperties: '{"Prop1": "Value1"}', localCustomProperties: '{"LocalProp1>": "LocalValue1>"}', termSetName: 'Department', termGroupName: 'People' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ "CreatedDate": "2018-10-22T19:11:43.669Z", "Id": "47fdacfe-ff64-4a05-b611-e84e767f04de", "LastModifiedDate": "2018-10-22T19:11:43.669Z", "Name": "IT", "CustomProperties": { "Prop1": "Value1" }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Description": "", "IsDeprecated": false, "IsKeyword": false, "IsPinned": false, "IsPinnedRoot": false, "IsReused": false, "IsRoot": true, "IsSourceTerm": true, "LocalCustomProperties": { "LocalProp1>": "LocalValue1>" }, "MergedTermIds": [], "PathOfTerm": "IT", "TermsCount": 0 }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { termGroupName: 'PnPTermSets', name: 'PnP-Organizations', id: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither termGroupId nor termGroupName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'PnP-Organizations' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both termGroupId and termGroupName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'PnP-Organizations', termGroupName: 'PnPTermSets', termGroupId: 'aca21974-139c-44fd-813c-6bbe6f25e658' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if termGroupId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupId: 'invalid', termSetName: 'Department' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither termSetId nor termSetName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupId: '9e54299e-208a-4000-8546-cc4139091b27' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both termSetId and termSetName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupId: '9e54299e-208a-4000-8546-cc4139091b27', termSetId: '9e54299e-208a-4000-8546-cc4139091b28', termSetName: 'Department' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both parentTermId and termSetName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupId: '9e54299e-208a-4000-8546-cc4139091b27', parentTermId: '9e54299e-208a-4000-8546-cc4139091b28', termSetName: 'Department' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both parentTermId and termSetId specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupId: '9e54299e-208a-4000-8546-cc4139091b27', parentTermId: '9e54299e-208a-4000-8546-cc4139091b28', termSetId: '9e54299e-208a-4000-8546-cc4139091b29' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both parentTermId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupId: '9e54299e-208a-4000-8546-cc4139091b27', parentTermId: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if termSetId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupId: '9e54299e-208a-4000-8546-cc4139091b27', termSetId: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if custom properties is not a valid JSON string', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupName: 'People', termSetName: 'Department', customProperties: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if local custom properties is not a valid JSON string', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupName: 'People', termSetName: 'Department', localCustomProperties: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id, termSetId and termGroupId specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', id: '9e54299e-208a-4000-8546-cc4139091b26', termGroupId: '9e54299e-208a-4000-8546-cc4139091b27', termSetId: '9e54299e-208a-4000-8546-cc4139091b28' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when id, termSetName and termGroupName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', id: '9e54299e-208a-4000-8546-cc4139091b26', termGroupName: 'People', termSetName: 'Department' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when id, parentTermId and termGroupName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', id: '9e54299e-208a-4000-8546-cc4139091b26', termGroupName: 'People', parentTermId: '9e54299e-208a-4000-8546-cc4139091b26' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when custom properties is a valid JSON string', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupName: 'People', termSetName: 'Department', customProperties: '{}' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when local custom properties is a valid JSON string', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'IT', termGroupName: 'People', termSetName: 'Department', localCustomProperties: '{}' } });
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
});