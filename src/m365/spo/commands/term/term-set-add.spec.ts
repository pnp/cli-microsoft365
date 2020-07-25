import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./term-set-add');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.TERM_SET_ADD, () => {
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
    assert.strictEqual(command.name.startsWith(commands.TERM_SET_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds term set to term group specified with id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" /><Method Id="42" ParentId="40" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">PnP-Organizations</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8119.1219", "ErrorInfo": null, "TraceCorrelationId": "3231949e-109d-0000-2cdb-ef525ee6aff1"
            }, 35, {
              "IsNull": false
            }, 36, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 38, {
              "IsNull": false
            }, 39, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 41, {
              "IsNull": false
            }, 43, {
              "IsNull": false
            }, 44, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
            }, 46, {
              "IsNull": false
            }, 47, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB"
            }, 48, {
              "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB", "CreatedDate": "\/Date(1538418692608)\/", "Id": "\/Guid(b53f9aa1-1d35-4b39-8498-7e4705e57301)\/", "LastModifiedDate": "\/Date(1538418692608)\/", "Name": "PnP-Organizations", "CustomProperties": {

              }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                "1033": "PnP-Organizations"
              }, "Stakeholders": [

              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'PnP-Organizations', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          CreatedDate: '2018-10-01T18:31:32.608Z',
          Id: 'b53f9aa1-1d35-4b39-8498-7e4705e57301',
          LastModifiedDate: '2018-10-01T18:31:32.608Z',
          Name: 'PnP-Organizations',
          CustomProperties: {},
          CustomSortOrder: null,
          IsAvailableForTagging: true,
          Owner: 'i:0#.f|membership|admin@contoso.onmicrosoft.com',
          Contact: '',
          Description: '',
          IsOpenForTermCreation: false,
          Names: { '1033': 'PnP-Organizations' },
          Stakeholders: []
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds term set to term group specified with name', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" /><Method Id="42" ParentId="40" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">PnP-Organizations</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8119.1219", "ErrorInfo": null, "TraceCorrelationId": "3231949e-109d-0000-2cdb-ef525ee6aff1"
            }, 35, {
              "IsNull": false
            }, 36, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 38, {
              "IsNull": false
            }, 39, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 41, {
              "IsNull": false
            }, 43, {
              "IsNull": false
            }, 44, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
            }, 46, {
              "IsNull": false
            }, 47, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB"
            }, 48, {
              "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB", "CreatedDate": "\/Date(1538418692608)\/", "Id": "\/Guid(b53f9aa1-1d35-4b39-8498-7e4705e57301)\/", "LastModifiedDate": "\/Date(1538418692608)\/", "Name": "PnP-Organizations", "CustomProperties": {

              }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                "1033": "PnP-Organizations"
              }, "Stakeholders": [

              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'PnP-Organizations', termGroupName: 'PnPTermSets' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          CreatedDate: '2018-10-01T18:31:32.608Z',
          Id: 'b53f9aa1-1d35-4b39-8498-7e4705e57301',
          LastModifiedDate: '2018-10-01T18:31:32.608Z',
          Name: 'PnP-Organizations',
          CustomProperties: {},
          CustomSortOrder: null,
          IsAvailableForTagging: true,
          Owner: 'i:0#.f|membership|admin@contoso.onmicrosoft.com',
          Contact: '',
          Description: '',
          IsOpenForTermCreation: false,
          Names: { '1033': 'PnP-Organizations' },
          Stakeholders: []
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds term set with a specified id to term group specified with id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" /><Method Id="42" ParentId="40" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">PnP-Organizations</Parameter><Parameter Type="Guid">{b53f9aa1-1d35-4b39-8498-7e4705e57301}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8119.1219", "ErrorInfo": null, "TraceCorrelationId": "3231949e-109d-0000-2cdb-ef525ee6aff1"
            }, 35, {
              "IsNull": false
            }, 36, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 38, {
              "IsNull": false
            }, 39, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 41, {
              "IsNull": false
            }, 43, {
              "IsNull": false
            }, 44, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
            }, 46, {
              "IsNull": false
            }, 47, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB"
            }, 48, {
              "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB", "CreatedDate": "\/Date(1538418692608)\/", "Id": "\/Guid(b53f9aa1-1d35-4b39-8498-7e4705e57301)\/", "LastModifiedDate": "\/Date(1538418692608)\/", "Name": "PnP-Organizations", "CustomProperties": {

              }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                "1033": "PnP-Organizations"
              }, "Stakeholders": [

              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'PnP-Organizations', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb', id: 'b53f9aa1-1d35-4b39-8498-7e4705e57301' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          CreatedDate: '2018-10-01T18:31:32.608Z',
          Id: 'b53f9aa1-1d35-4b39-8498-7e4705e57301',
          LastModifiedDate: '2018-10-01T18:31:32.608Z',
          Name: 'PnP-Organizations',
          CustomProperties: {},
          CustomSortOrder: null,
          IsAvailableForTagging: true,
          Owner: 'i:0#.f|membership|admin@contoso.onmicrosoft.com',
          Contact: '',
          Description: '',
          IsOpenForTermCreation: false,
          Names: { '1033': 'PnP-Organizations' },
          Stakeholders: []
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds term set with a specified description to term group specified with id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" /><Method Id="42" ParentId="40" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">PnP-Organizations</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8119.1219", "ErrorInfo": null, "TraceCorrelationId": "3231949e-109d-0000-2cdb-ef525ee6aff1"
            }, 35, {
              "IsNull": false
            }, 36, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 38, {
              "IsNull": false
            }, 39, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 41, {
              "IsNull": false
            }, 43, {
              "IsNull": false
            }, 44, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
            }, 46, {
              "IsNull": false
            }, 47, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB"
            }, 48, {
              "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB", "CreatedDate": "\/Date(1538418692608)\/", "Id": "\/Guid(b53f9aa1-1d35-4b39-8498-7e4705e57301)\/", "LastModifiedDate": "\/Date(1538418692608)\/", "Name": "PnP-Organizations", "CustomProperties": {

              }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                "1033": "PnP-Organizations"
              }, "Stakeholders": [

              ]
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="127" ObjectPathId="117" Name="Description"><Parameter Type="String">List of organizations</Parameter></SetProperty><Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB" /><Identity Id="109" Name="3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8119.1219", "ErrorInfo": null, "TraceCorrelationId": "1332949e-609b-0000-2cdb-e238bddae823"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, name: 'PnP-Organizations', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb', description: 'List of organizations' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          CreatedDate: '2018-10-01T18:31:32.608Z',
          Id: 'b53f9aa1-1d35-4b39-8498-7e4705e57301',
          LastModifiedDate: '2018-10-01T18:31:32.608Z',
          Name: 'PnP-Organizations',
          CustomProperties: {},
          CustomSortOrder: null,
          IsAvailableForTagging: true,
          Owner: 'i:0#.f|membership|admin@contoso.onmicrosoft.com',
          Contact: '',
          Description: 'List of organizations',
          IsOpenForTermCreation: false,
          Names: { '1033': 'PnP-Organizations' },
          Stakeholders: []
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds term set with custom properties to term group specified with id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" /><Method Id="42" ParentId="40" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">PnP-Organizations</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8119.1219", "ErrorInfo": null, "TraceCorrelationId": "3231949e-109d-0000-2cdb-ef525ee6aff1"
            }, 35, {
              "IsNull": false
            }, 36, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 38, {
              "IsNull": false
            }, 39, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 41, {
              "IsNull": false
            }, 43, {
              "IsNull": false
            }, 44, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
            }, 46, {
              "IsNull": false
            }, 47, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB"
            }, 48, {
              "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB", "CreatedDate": "\/Date(1538418692608)\/", "Id": "\/Guid(b53f9aa1-1d35-4b39-8498-7e4705e57301)\/", "LastModifiedDate": "\/Date(1538418692608)\/", "Name": "PnP-Organizations", "CustomProperties": {

              }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                "1033": "PnP-Organizations"
              }, "Stakeholders": [

              ]
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetCustomProperty" Id="127" ObjectPathId="117"><Parameters><Parameter Type="String">Prop1</Parameter><Parameter Type="String">Value 1</Parameter></Parameters></Method><Method Name="SetCustomProperty" Id="128" ObjectPathId="117"><Parameters><Parameter Type="String">Prop2</Parameter><Parameter Type="String">Value 2</Parameter></Parameters></Method><Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB" /><Identity Id="109" Name="3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8119.1219", "ErrorInfo": null, "TraceCorrelationId": "1332949e-609b-0000-2cdb-e238bddae823"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'PnP-Organizations', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb', customProperties: JSON.stringify({ Prop1: 'Value 1', Prop2: 'Value 2' }) } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          CreatedDate: '2018-10-01T18:31:32.608Z',
          Id: 'b53f9aa1-1d35-4b39-8498-7e4705e57301',
          LastModifiedDate: '2018-10-01T18:31:32.608Z',
          Name: 'PnP-Organizations',
          CustomProperties: {
            Prop1: 'Value 1',
            Prop2: 'Value 2'
          },
          CustomSortOrder: null,
          IsAvailableForTagging: true,
          Owner: 'i:0#.f|membership|admin@contoso.onmicrosoft.com',
          Contact: '',
          Description: '',
          IsOpenForTermCreation: false,
          Names: { '1033': 'PnP-Organizations' },
          Stakeholders: []
        }));
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
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" /><Method Id="42" ParentId="40" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">PnP-Organizations</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
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
    cmdInstance.action({ options: { debug: false, name: 'PnP-Organizations', termGroupName: 'PnPTermSets' } }, (err?: any) => {
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
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" /><Method Id="42" ParentId="40" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">PnP-Organizations</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
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
    cmdInstance.action({ options: { debug: false, name: 'PnP-Organizations', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb' } }, (err?: any) => {
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
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" /><Method Id="42" ParentId="40" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">PnP-Organizations</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
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
    cmdInstance.action({ options: { debug: false, name: 'PnP-Organizations', termGroupName: 'PnPTermSets' } }, (err?: any) => {
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
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" /><Method Id="42" ParentId="40" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">PnP-Organizations</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8105.1217", "ErrorInfo": {
                "ErrorMessage": "A term set already exists with the name specified.", "ErrorValue": null, "TraceCorrelationId": "3105909e-e037-0000-29c7-078ce31cbc78", "ErrorCode": -2146233086, "ErrorTypeName": "Microsoft.SharePoint.Taxonomy.TermStoreOperationException"
              }, "TraceCorrelationId": "3105909e-e037-0000-29c7-078ce31cbc78"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'PnP-Organizations', termGroupName: 'PnPTermSets' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('A term set already exists with the name specified.')));
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
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" /><Method Id="42" ParentId="40" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">PnP-Organizations</Parameter><Parameter Type="Guid">{aca21974-139c-44fd-813c-6bbe6f25e658}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8105.1217", "ErrorInfo": {
                "ErrorMessage": "Failed to read from or write to database. Refresh and try again. If the problem persists, please contact the administrator.", "ErrorValue": null, "TraceCorrelationId": "3105909e-e037-0000-29c7-078ce31cbc78", "ErrorCode": -2146233086, "ErrorTypeName": "Microsoft.SharePoint.Taxonomy.TermStoreOperationException"
              }, "TraceCorrelationId": "3105909e-e037-0000-29c7-078ce31cbc78"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'PnP-Organizations', id: 'aca21974-139c-44fd-813c-6bbe6f25e658', termGroupName: 'PnPTermSets' } }, (err?: any) => {
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
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" /><Method Id="42" ParentId="40" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">PnP-Organizations</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8119.1219", "ErrorInfo": null, "TraceCorrelationId": "3231949e-109d-0000-2cdb-ef525ee6aff1"
            }, 35, {
              "IsNull": false
            }, 36, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 38, {
              "IsNull": false
            }, 39, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 41, {
              "IsNull": false
            }, 43, {
              "IsNull": false
            }, 44, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
            }, 46, {
              "IsNull": false
            }, 47, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB"
            }, 48, {
              "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB", "CreatedDate": "\/Date(1538418692608)\/", "Id": "\/Guid(b53f9aa1-1d35-4b39-8498-7e4705e57301)\/", "LastModifiedDate": "\/Date(1538418692608)\/", "Name": "PnP-Organizations", "CustomProperties": {

              }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                "1033": "PnP-Organizations"
              }, "Stakeholders": [

              ]
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="127" ObjectPathId="117" Name="Description"><Parameter Type="String">List of organizations</Parameter></SetProperty><Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB" /><Identity Id="109" Name="3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
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
    cmdInstance.action({ options: { debug: false, name: 'PnP-Organizations', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb', description: 'List of organizations' } }, (err?: any) => {
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
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" /><Method Id="42" ParentId="40" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">PnP-Organizations</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8119.1219", "ErrorInfo": null, "TraceCorrelationId": "3231949e-109d-0000-2cdb-ef525ee6aff1"
            }, 35, {
              "IsNull": false
            }, 36, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 38, {
              "IsNull": false
            }, 39, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 41, {
              "IsNull": false
            }, 43, {
              "IsNull": false
            }, 44, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
            }, 46, {
              "IsNull": false
            }, 47, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB"
            }, 48, {
              "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB", "CreatedDate": "\/Date(1538418692608)\/", "Id": "\/Guid(b53f9aa1-1d35-4b39-8498-7e4705e57301)\/", "LastModifiedDate": "\/Date(1538418692608)\/", "Name": "PnP-Organizations", "CustomProperties": {

              }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                "1033": "PnP-Organizations"
              }, "Stakeholders": [

              ]
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetCustomProperty" Id="127" ObjectPathId="117"><Parameters><Parameter Type="String">Prop1</Parameter><Parameter Type="String">Value 1</Parameter></Parameters></Method><Method Name="SetCustomProperty" Id="128" ObjectPathId="117"><Parameters><Parameter Type="String">Prop2</Parameter><Parameter Type="String">Value 2</Parameter></Parameters></Method><Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB" /><Identity Id="109" Name="3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
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
    cmdInstance.action({ options: { debug: false, name: 'PnP-Organizations', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb', customProperties: JSON.stringify({ Prop1: 'Value 1', Prop2: 'Value 2' }) } }, (err?: any) => {
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
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" /><Method Id="42" ParentId="40" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets&gt;</Parameter></Parameters></Method><Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">PnP-Organizations</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8119.1219", "ErrorInfo": null, "TraceCorrelationId": "3231949e-109d-0000-2cdb-ef525ee6aff1"
            }, 35, {
              "IsNull": false
            }, 36, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 38, {
              "IsNull": false
            }, 39, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 41, {
              "IsNull": false
            }, 43, {
              "IsNull": false
            }, 44, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
            }, 46, {
              "IsNull": false
            }, 47, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB"
            }, 48, {
              "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB", "CreatedDate": "\/Date(1538418692608)\/", "Id": "\/Guid(b53f9aa1-1d35-4b39-8498-7e4705e57301)\/", "LastModifiedDate": "\/Date(1538418692608)\/", "Name": "PnP-Organizations", "CustomProperties": {

              }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                "1033": "PnP-Organizations"
              }, "Stakeholders": [

              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'PnP-Organizations', termGroupName: 'PnPTermSets>' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          CreatedDate: '2018-10-01T18:31:32.608Z',
          Id: 'b53f9aa1-1d35-4b39-8498-7e4705e57301',
          LastModifiedDate: '2018-10-01T18:31:32.608Z',
          Name: 'PnP-Organizations',
          CustomProperties: {},
          CustomSortOrder: null,
          IsAvailableForTagging: true,
          Owner: 'i:0#.f|membership|admin@contoso.onmicrosoft.com',
          Contact: '',
          Description: '',
          IsOpenForTermCreation: false,
          Names: { '1033': 'PnP-Organizations' },
          Stakeholders: []
        }));
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
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" /><Method Id="42" ParentId="40" Name="GetByName"><Parameters><Parameter Type="String">PnPTermSets</Parameter></Parameters></Method><Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">PnP-Organizations&gt;</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8119.1219", "ErrorInfo": null, "TraceCorrelationId": "3231949e-109d-0000-2cdb-ef525ee6aff1"
            }, 35, {
              "IsNull": false
            }, 36, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 38, {
              "IsNull": false
            }, 39, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 41, {
              "IsNull": false
            }, 43, {
              "IsNull": false
            }, 44, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
            }, 46, {
              "IsNull": false
            }, 47, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB"
            }, 48, {
              "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB", "CreatedDate": "\/Date(1538418692608)\/", "Id": "\/Guid(b53f9aa1-1d35-4b39-8498-7e4705e57301)\/", "LastModifiedDate": "\/Date(1538418692608)\/", "Name": "PnP-Organizations>", "CustomProperties": {

              }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                "1033": "PnP-Organizations>"
              }, "Stakeholders": [

              ]
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'PnP-Organizations>', termGroupName: 'PnPTermSets' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          CreatedDate: '2018-10-01T18:31:32.608Z',
          Id: 'b53f9aa1-1d35-4b39-8498-7e4705e57301',
          LastModifiedDate: '2018-10-01T18:31:32.608Z',
          Name: 'PnP-Organizations>',
          CustomProperties: {},
          CustomSortOrder: null,
          IsAvailableForTagging: true,
          Owner: 'i:0#.f|membership|admin@contoso.onmicrosoft.com',
          Contact: '',
          Description: '',
          IsOpenForTermCreation: false,
          Names: { '1033': 'PnP-Organizations>' },
          Stakeholders: []
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly escapes XML in term set description', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" /><Method Id="42" ParentId="40" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">PnP-Organizations</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8119.1219", "ErrorInfo": null, "TraceCorrelationId": "3231949e-109d-0000-2cdb-ef525ee6aff1"
            }, 35, {
              "IsNull": false
            }, 36, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 38, {
              "IsNull": false
            }, 39, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 41, {
              "IsNull": false
            }, 43, {
              "IsNull": false
            }, 44, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
            }, 46, {
              "IsNull": false
            }, 47, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB"
            }, 48, {
              "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB", "CreatedDate": "\/Date(1538418692608)\/", "Id": "\/Guid(b53f9aa1-1d35-4b39-8498-7e4705e57301)\/", "LastModifiedDate": "\/Date(1538418692608)\/", "Name": "PnP-Organizations", "CustomProperties": {

              }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                "1033": "PnP-Organizations"
              }, "Stakeholders": [

              ]
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="127" ObjectPathId="117" Name="Description"><Parameter Type="String">List of organizations&gt;</Parameter></SetProperty><Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB" /><Identity Id="109" Name="3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8119.1219", "ErrorInfo": null, "TraceCorrelationId": "1332949e-609b-0000-2cdb-e238bddae823"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'PnP-Organizations', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb', description: 'List of organizations>' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          CreatedDate: '2018-10-01T18:31:32.608Z',
          Id: 'b53f9aa1-1d35-4b39-8498-7e4705e57301',
          LastModifiedDate: '2018-10-01T18:31:32.608Z',
          Name: 'PnP-Organizations',
          CustomProperties: {},
          CustomSortOrder: null,
          IsAvailableForTagging: true,
          Owner: 'i:0#.f|membership|admin@contoso.onmicrosoft.com',
          Contact: '',
          Description: 'List of organizations>',
          IsOpenForTermCreation: false,
          Names: { '1033': 'PnP-Organizations' },
          Stakeholders: []
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly escapes XML in term set custom properties', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" /><Method Id="42" ParentId="40" Name="GetById"><Parameters><Parameter Type="Guid">{0e8f395e-ff58-4d45-9ff7-e331ab728beb}</Parameter></Parameters></Method><Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">PnP-Organizations</Parameter><Parameter Type="Guid">{`) > -1 && opts.body.indexOf(`}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8119.1219", "ErrorInfo": null, "TraceCorrelationId": "3231949e-109d-0000-2cdb-ef525ee6aff1"
            }, 35, {
              "IsNull": false
            }, 36, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 38, {
              "IsNull": false
            }, 39, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 41, {
              "IsNull": false
            }, 43, {
              "IsNull": false
            }, 44, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+s="
            }, 46, {
              "IsNull": false
            }, 47, {
              "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB"
            }, 48, {
              "_ObjectType_": "SP.Taxonomy.TermSet", "_ObjectIdentity_": "3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB", "CreatedDate": "\/Date(1538418692608)\/", "Id": "\/Guid(b53f9aa1-1d35-4b39-8498-7e4705e57301)\/", "LastModifiedDate": "\/Date(1538418692608)\/", "Name": "PnP-Organizations", "CustomProperties": {

              }, "CustomSortOrder": null, "IsAvailableForTagging": true, "Owner": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "Contact": "", "Description": "", "IsOpenForTermCreation": false, "Names": {
                "1033": "PnP-Organizations"
              }, "Stakeholders": [

              ]
            }
          ]));
        }

        if (opts.body.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetCustomProperty" Id="127" ObjectPathId="117"><Parameters><Parameter Type="String">Prop1</Parameter><Parameter Type="String">&lt;Value 1</Parameter></Parameters></Method><Method Name="SetCustomProperty" Id="128" ObjectPathId="117"><Parameters><Parameter Type="String">Prop2</Parameter><Parameter Type="String">Value 2&gt;</Parameter></Parameters></Method><Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:se:YU1+cBy9wUuh\u002ffzgFZGpUV45jw5Y\u002f0VNn\u002ffjMatyi+uhmj+1NR05S4SYfkcF5XMB" /><Identity Id="109" Name="3231949e-109d-0000-2cdb-ef525ee6aff1|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8119.1219", "ErrorInfo": null, "TraceCorrelationId": "1332949e-609b-0000-2cdb-e238bddae823"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, name: 'PnP-Organizations', termGroupId: '0e8f395e-ff58-4d45-9ff7-e331ab728beb', customProperties: JSON.stringify({ Prop1: '<Value 1', Prop2: 'Value 2>' }) } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          CreatedDate: '2018-10-01T18:31:32.608Z',
          Id: 'b53f9aa1-1d35-4b39-8498-7e4705e57301',
          LastModifiedDate: '2018-10-01T18:31:32.608Z',
          Name: 'PnP-Organizations',
          CustomProperties: {
            Prop1: '<Value 1',
            Prop2: 'Value 2>'
          },
          CustomSortOrder: null,
          IsAvailableForTagging: true,
          Owner: 'i:0#.f|membership|admin@contoso.onmicrosoft.com',
          Contact: '',
          Description: '',
          IsOpenForTermCreation: false,
          Names: { '1033': 'PnP-Organizations' },
          Stakeholders: []
        }));
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
    const actual = (command.validate() as CommandValidate)({ options: { name: 'PnP-Organizations', termGroupId: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if custom properties is not a valid JSON string', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'PnP-Organizations', termGroupName: 'PnPTermSets', customProperties: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id, name and termGroupId specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'PnP-Organizations', id: '9e54299e-208a-4000-8546-cc4139091b26', termGroupId: '9e54299e-208a-4000-8546-cc4139091b27' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when id, name and termGroupName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'PnP-Organizations', id: '9e54299e-208a-4000-8546-cc4139091b26', termGroupName: 'PnPTermSets' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when name and termGroupId specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'People', termGroupId: '9e54299e-208a-4000-8546-cc4139091b26' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when name and termGroupName specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'People', termGroupName: 'PnPTermSets' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when custom properties is a valid JSON string', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'PnP-Organizations', termGroupName: 'PnPTermSets', customProperties: '{}' } });
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