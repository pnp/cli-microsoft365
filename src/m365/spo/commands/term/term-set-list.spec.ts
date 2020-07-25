import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./term-set-list');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.TERM_SET_LIST, () => {
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
    assert.strictEqual(command.name.startsWith(commands.TERM_SET_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists taxonomy term sets from the term group specified using id (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54" /><ObjectIdentityQuery Id="56" ObjectPathId="54" /><ObjectPath Id="58" ObjectPathId="57" /><ObjectIdentityQuery Id="59" ObjectPathId="57" /><ObjectPath Id="61" ObjectPathId="60" /><ObjectPath Id="63" ObjectPathId="62" /><ObjectIdentityQuery Id="64" ObjectPathId="62" /><ObjectPath Id="66" ObjectPathId="65" /><Query Id="67" ObjectPathId="65"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="54" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="57" ParentId="54" Name="GetDefaultSiteCollectionTermStore" /><Property Id="60" ParentId="57" Name="Groups" /><Method Id="62" ParentId="60" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Property Id="65" ParentId="62" Name="TermSets" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "ec3f929e-2007-0000-2cdb-ebdf7451c224"
          }, 55, {
            "IsNull": false
          }, 56, {
            "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          }, 58, {
            "IsNull": false
          }, 59, {
            "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          }, 61, {
            "IsNull": false
          }, 63, {
            "IsNull": false
          }, 64, {
            "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
          }, 66, {
            "IsNull": false
          }, 67, {
            "_ObjectType_": "SP.Taxonomy.TermSetCollection", "_Child_Items_": [
              {
                "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+tHfBZ6NyvQQZTQ6WLBpPLt", "CreatedDate": "\/Date(1536839573337)\/", "Id": "\/Guid(7a167c47-2b37-41d0-94d0-e962c1a4f2ed)\/", "LastModifiedDate": "\/Date(1536840826883)\/", "Name": "PnP-CollabFooter-SharedLinks", "CustomProperties": {
                  "_Sys_Nav_IsNavigationTermSet": "True"
                }, "CustomSortOrder": "a359ee29-cf72-4235-a4ef-1ed96bf4eaea:60d165e6-8cb1-4c20-8fad-80067c4ca767:da7bfb84-008b-48ff-b61f-bfe40da2602f", "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@m365x035040.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                  "1033": "PnP-CollabFooter-SharedLinks"
                }, "Stakeholders": [

                ]
              }, {
                "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7", "CreatedDate": "\/Date(1536839575147)\/", "Id": "\/Guid(1479e26c-1380-41a8-9183-72bc5a9651bb)\/", "LastModifiedDate": "\/Date(1536840827383)\/", "Name": "PnP-Organizations", "CustomProperties": {

                }, "CustomSortOrder": "02cf219e-8ce9-4e85-ac04-a913a44a5d2b:247543b6-45f2-4232-b9e8-66c5bf53c31e:ffc3608f-1250-4d28-b388-381fad8d4602", "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@m365x035040.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                  "1033": "PnP-Organizations"
                }, "Stakeholders": [

                ]
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          { "Id": "7a167c47-2b37-41d0-94d0-e962c1a4f2ed", "Name": "PnP-CollabFooter-SharedLinks" },
          { "Id": "1479e26c-1380-41a8-9183-72bc5a9651bb", "Name": "PnP-Organizations" }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists taxonomy term sets from the term group specified using name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54" /><ObjectIdentityQuery Id="56" ObjectPathId="54" /><ObjectPath Id="58" ObjectPathId="57" /><ObjectIdentityQuery Id="59" ObjectPathId="57" /><ObjectPath Id="61" ObjectPathId="60" /><ObjectPath Id="63" ObjectPathId="62" /><ObjectIdentityQuery Id="64" ObjectPathId="62" /><ObjectPath Id="66" ObjectPathId="65" /><Query Id="67" ObjectPathId="65"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="54" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="57" ParentId="54" Name="GetDefaultSiteCollectionTermStore" /><Property Id="60" ParentId="57" Name="Groups" /><Method Id="62" ParentId="60" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="65" ParentId="62" Name="TermSets" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "ec3f929e-2007-0000-2cdb-ebdf7451c224"
          }, 55, {
            "IsNull": false
          }, 56, {
            "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          }, 58, {
            "IsNull": false
          }, 59, {
            "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          }, 61, {
            "IsNull": false
          }, 63, {
            "IsNull": false
          }, 64, {
            "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
          }, 66, {
            "IsNull": false
          }, 67, {
            "_ObjectType_": "SP.Taxonomy.TermSetCollection", "_Child_Items_": [
              {
                "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+tHfBZ6NyvQQZTQ6WLBpPLt", "CreatedDate": "\/Date(1536839573337)\/", "Id": "\/Guid(7a167c47-2b37-41d0-94d0-e962c1a4f2ed)\/", "LastModifiedDate": "\/Date(1536840826883)\/", "Name": "PnP-CollabFooter-SharedLinks", "CustomProperties": {
                  "_Sys_Nav_IsNavigationTermSet": "True"
                }, "CustomSortOrder": "a359ee29-cf72-4235-a4ef-1ed96bf4eaea:60d165e6-8cb1-4c20-8fad-80067c4ca767:da7bfb84-008b-48ff-b61f-bfe40da2602f", "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@m365x035040.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                  "1033": "PnP-CollabFooter-SharedLinks"
                }, "Stakeholders": [

                ]
              }, {
                "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7", "CreatedDate": "\/Date(1536839575147)\/", "Id": "\/Guid(1479e26c-1380-41a8-9183-72bc5a9651bb)\/", "LastModifiedDate": "\/Date(1536840827383)\/", "Name": "PnP-Organizations", "CustomProperties": {

                }, "CustomSortOrder": "02cf219e-8ce9-4e85-ac04-a913a44a5d2b:247543b6-45f2-4232-b9e8-66c5bf53c31e:ffc3608f-1250-4d28-b388-381fad8d4602", "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@m365x035040.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                  "1033": "PnP-Organizations"
                }, "Stakeholders": [

                ]
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, termGroupName: 'PnPTermSets' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          { "Id": "7a167c47-2b37-41d0-94d0-e962c1a4f2ed", "Name": "PnP-CollabFooter-SharedLinks" },
          { "Id": "1479e26c-1380-41a8-9183-72bc5a9651bb", "Name": "PnP-Organizations" }
        ]));
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
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54" /><ObjectIdentityQuery Id="56" ObjectPathId="54" /><ObjectPath Id="58" ObjectPathId="57" /><ObjectIdentityQuery Id="59" ObjectPathId="57" /><ObjectPath Id="61" ObjectPathId="60" /><ObjectPath Id="63" ObjectPathId="62" /><ObjectIdentityQuery Id="64" ObjectPathId="62" /><ObjectPath Id="66" ObjectPathId="65" /><Query Id="67" ObjectPathId="65"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="54" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="57" ParentId="54" Name="GetDefaultSiteCollectionTermStore" /><Property Id="60" ParentId="57" Name="Groups" /><Method Id="62" ParentId="60" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets&gt;</Parameter></Parameters></Method><Property Id="65" ParentId="62" Name="TermSets" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "ec3f929e-2007-0000-2cdb-ebdf7451c224"
          }, 55, {
            "IsNull": false
          }, 56, {
            "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          }, 58, {
            "IsNull": false
          }, 59, {
            "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          }, 61, {
            "IsNull": false
          }, 63, {
            "IsNull": false
          }, 64, {
            "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
          }, 66, {
            "IsNull": false
          }, 67, {
            "_ObjectType_": "SP.Taxonomy.TermSetCollection", "_Child_Items_": [
              {
                "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+tHfBZ6NyvQQZTQ6WLBpPLt", "CreatedDate": "\/Date(1536839573337)\/", "Id": "\/Guid(7a167c47-2b37-41d0-94d0-e962c1a4f2ed)\/", "LastModifiedDate": "\/Date(1536840826883)\/", "Name": "PnP-CollabFooter-SharedLinks", "CustomProperties": {
                  "_Sys_Nav_IsNavigationTermSet": "True"
                }, "CustomSortOrder": "a359ee29-cf72-4235-a4ef-1ed96bf4eaea:60d165e6-8cb1-4c20-8fad-80067c4ca767:da7bfb84-008b-48ff-b61f-bfe40da2602f", "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@m365x035040.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                  "1033": "PnP-CollabFooter-SharedLinks"
                }, "Stakeholders": [

                ]
              }, {
                "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7", "CreatedDate": "\/Date(1536839575147)\/", "Id": "\/Guid(1479e26c-1380-41a8-9183-72bc5a9651bb)\/", "LastModifiedDate": "\/Date(1536840827383)\/", "Name": "PnP-Organizations", "CustomProperties": {

                }, "CustomSortOrder": "02cf219e-8ce9-4e85-ac04-a913a44a5d2b:247543b6-45f2-4232-b9e8-66c5bf53c31e:ffc3608f-1250-4d28-b388-381fad8d4602", "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@m365x035040.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                  "1033": "PnP-Organizations"
                }, "Stakeholders": [

                ]
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, termGroupName: 'PnPTermSets>' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          { "Id": "7a167c47-2b37-41d0-94d0-e962c1a4f2ed", "Name": "PnP-CollabFooter-SharedLinks" },
          { "Id": "1479e26c-1380-41a8-9183-72bc5a9651bb", "Name": "PnP-Organizations" }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists taxonomy term sets with all properties when output is JSON', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54" /><ObjectIdentityQuery Id="56" ObjectPathId="54" /><ObjectPath Id="58" ObjectPathId="57" /><ObjectIdentityQuery Id="59" ObjectPathId="57" /><ObjectPath Id="61" ObjectPathId="60" /><ObjectPath Id="63" ObjectPathId="62" /><ObjectIdentityQuery Id="64" ObjectPathId="62" /><ObjectPath Id="66" ObjectPathId="65" /><Query Id="67" ObjectPathId="65"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="54" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="57" ParentId="54" Name="GetDefaultSiteCollectionTermStore" /><Property Id="60" ParentId="57" Name="Groups" /><Method Id="62" ParentId="60" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Property Id="65" ParentId="62" Name="TermSets" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "ec3f929e-2007-0000-2cdb-ebdf7451c224"
          }, 55, {
            "IsNull": false
          }, 56, {
            "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          }, 58, {
            "IsNull": false
          }, 59, {
            "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          }, 61, {
            "IsNull": false
          }, 63, {
            "IsNull": false
          }, 64, {
            "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
          }, 66, {
            "IsNull": false
          }, 67, {
            "_ObjectType_": "SP.Taxonomy.TermSetCollection", "_Child_Items_": [
              {
                "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+tHfBZ6NyvQQZTQ6WLBpPLt", "CreatedDate": "\/Date(1536839573337)\/", "Id": "\/Guid(7a167c47-2b37-41d0-94d0-e962c1a4f2ed)\/", "LastModifiedDate": "\/Date(1536840826883)\/", "Name": "PnP-CollabFooter-SharedLinks", "CustomProperties": {
                  "_Sys_Nav_IsNavigationTermSet": "True"
                }, "CustomSortOrder": "a359ee29-cf72-4235-a4ef-1ed96bf4eaea:60d165e6-8cb1-4c20-8fad-80067c4ca767:da7bfb84-008b-48ff-b61f-bfe40da2602f", "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@m365x035040.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                  "1033": "PnP-CollabFooter-SharedLinks"
                }, "Stakeholders": [

                ]
              }, {
                "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7", "CreatedDate": "\/Date(1536839575147)\/", "Id": "\/Guid(1479e26c-1380-41a8-9183-72bc5a9651bb)\/", "LastModifiedDate": "\/Date(1536840827383)\/", "Name": "PnP-Organizations", "CustomProperties": {

                }, "CustomSortOrder": "02cf219e-8ce9-4e85-ac04-a913a44a5d2b:247543b6-45f2-4232-b9e8-66c5bf53c31e:ffc3608f-1250-4d28-b388-381fad8d4602", "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@m365x035040.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                  "1033": "PnP-Organizations"
                }, "Stakeholders": [

                ]
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { output: 'json', termGroupName: 'PnPTermSets' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([{ "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+tHfBZ6NyvQQZTQ6WLBpPLt", "CreatedDate": "2018-09-13T11:52:53.337Z", "Id": "7a167c47-2b37-41d0-94d0-e962c1a4f2ed", "LastModifiedDate": "2018-09-13T12:13:46.883Z", "Name": "PnP-CollabFooter-SharedLinks", "CustomProperties": { "_Sys_Nav_IsNavigationTermSet": "True" }, "CustomSortOrder": "a359ee29-cf72-4235-a4ef-1ed96bf4eaea:60d165e6-8cb1-4c20-8fad-80067c4ca767:da7bfb84-008b-48ff-b61f-bfe40da2602f", "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@m365x035040.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": { "1033": "PnP-CollabFooter-SharedLinks" }, "Stakeholders": [] }, { "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "ec3f929e-2007-0000-2cdb-ebdf7451c224|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+ts4nkUgBOoQZGDcrxallG7", "CreatedDate": "2018-09-13T11:52:55.147Z", "Id": "1479e26c-1380-41a8-9183-72bc5a9651bb", "LastModifiedDate": "2018-09-13T12:13:47.383Z", "Name": "PnP-Organizations", "CustomProperties": {}, "CustomSortOrder": "02cf219e-8ce9-4e85-ac04-a913a44a5d2b:247543b6-45f2-4232-b9e8-66c5bf53c31e:ffc3608f-1250-4d28-b388-381fad8d4602", "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@m365x035040.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": { "1033": "PnP-Organizations" }, "Stakeholders": [] }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no term sets found', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54" /><ObjectIdentityQuery Id="56" ObjectPathId="54" /><ObjectPath Id="58" ObjectPathId="57" /><ObjectIdentityQuery Id="59" ObjectPathId="57" /><ObjectPath Id="61" ObjectPathId="60" /><ObjectPath Id="63" ObjectPathId="62" /><ObjectIdentityQuery Id="64" ObjectPathId="62" /><ObjectPath Id="66" ObjectPathId="65" /><Query Id="67" ObjectPathId="65"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="54" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="57" ParentId="54" Name="GetDefaultSiteCollectionTermStore" /><Property Id="60" ParentId="57" Name="Groups" /><Method Id="62" ParentId="60" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Property Id="65" ParentId="62" Name="TermSets" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.8105.1215",
            "ErrorInfo": null,
            "TraceCorrelationId": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9"
          },
          4,
          {
            "IsNull": false
          },
          5,
          {
            "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
          },
          7,
          {
            "IsNull": false
          },
          8,
          {
            "_ObjectIdentity_": "40bc8e9e-c0f3-0000-2b65-64d3c82fb3d9|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
          },
          10,
          {
            "IsNull": false
          },
          11,
          {
            "_ObjectType_": "SP.Taxonomy.TermSetCollection",
            "_Child_Items_": []
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving taxonomy term sets', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
            "ErrorMessage": "File Not Found.", "ErrorValue": null, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26", "ErrorCode": -2147024894, "ErrorTypeName": "System.IO.FileNotFoundException"
          }, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26"
        }
      ]));
    });
    cmdInstance.action({ options: { debug: false, termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('File Not Found.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified term group id doesn\'t exist', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": {
            "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "0140929e-a0f5-0000-2cdb-ea8d3db8259b", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
          }, "TraceCorrelationId": "0140929e-a0f5-0000-2cdb-ea8d3db8259b"
        }
      ]));
    });
    cmdInstance.action({ options: { debug: false, termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified term group name doesn\'t exist', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": {
            "ErrorMessage": "Specified argument was out of the range of valid values.\r\nParameter name: index", "ErrorValue": null, "TraceCorrelationId": "0c40929e-00f7-0000-2cdb-e77493720fa6", "ErrorCode": -2146233086, "ErrorTypeName": "System.ArgumentOutOfRangeException"
          }, "TraceCorrelationId": "0c40929e-00f7-0000-2cdb-e77493720fa6"
        }
      ]));
    });
    cmdInstance.action({ options: { debug: false, termGroupName: 'PnPTermSets' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Specified argument was out of the range of valid values.\r\nParameter name: index')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when neither id nor name specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both id and name specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { termGroupId: '9e54299e-208a-4000-8546-cc4139091b26', termGroupName: 'PnPTermSets' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { termGroupId: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { termGroupId: '9e54299e-208a-4000-8546-cc4139091b26' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when name specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { termGroupName: 'PnPTermSets' } });
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
      options: { debug: false, termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb' }
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