import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./hubsite-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.HUBSITE_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      request.get,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.HUBSITE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists hub sites', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/hubsites`) > -1) {
        return Promise.resolve({
          value: [
            {
              "Description": null,
              "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
              "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
              "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
              "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
              "Targets": null,
              "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
              "Title": "Sales"
            },
            {
              "Description": null,
              "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
              "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
              "SiteId": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
              "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
              "Targets": null,
              "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
              "Title": "Travel Programs"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
            "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
            "Title": "Sales"
          },
          {
            "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
            "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
            "Title": "Travel Programs"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists hub sites (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/hubsites`) > -1) {
        return Promise.resolve({
          value: [
            {
              "Description": null,
              "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
              "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
              "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
              "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
              "Targets": null,
              "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
              "Title": "Sales"
            },
            {
              "Description": null,
              "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
              "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
              "SiteId": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
              "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
              "Targets": null,
              "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
              "Title": "Travel Programs"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
            "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
            "Title": "Sales"
          },
          {
            "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
            "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
            "Title": "Travel Programs"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists hub sites with all properties for JSON output', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/hubsites`) > -1) {
        return Promise.resolve({
          value: [
            {
              "Description": null,
              "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
              "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
              "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
              "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
              "Targets": null,
              "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
              "Title": "Sales"
            },
            {
              "Description": null,
              "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
              "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
              "SiteId": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
              "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
              "Targets": null,
              "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
              "Title": "Travel Programs"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "Description": null,
            "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
            "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
            "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
            "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
            "Targets": null,
            "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
            "Title": "Sales"
          },
          {
            "Description": null,
            "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
            "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
            "SiteId": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
            "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
            "Targets": null,
            "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
            "Title": "Travel Programs"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('does not list associated sites allong the hub sites, if the includeAssociatedSites option is provided, if the output is TEXT', (done) => {
    sinon.stub(request, 'get').resolves({
      value: [
        {
          "Description": null,
          "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Sales"
        },
        {
          "Description": null,
          "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
          "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Travel Programs"
        }
      ]
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/RenderListDataAsStream`) > -1
        && JSON.stringify(opts.body) === JSON.stringify({
          parameters: {
            ViewXml: "<View><Query><Where><And><And><IsNull><FieldRef Name=\"TimeDeleted\"/></IsNull><Neq><FieldRef Name=\"State\"/><Value Type='Integer'>0</Value></Neq></And><Neq><FieldRef Name=\"HubSiteId\"/><Value Type='Text'>{00000000-0000-0000-0000-000000000000}</Value></Neq></And></Where><OrderBy><FieldRef Name='Title' Ascending='true' /></OrderBy></Query><ViewFields><FieldRef Name=\"Title\"/><FieldRef Name=\"SiteUrl\"/><FieldRef Name=\"SiteId\"/><FieldRef Name=\"HubSiteId\"/></ViewFields><RowLimit Paged=\"TRUE\">30</RowLimit></View>",
            DatesInUtc: true
          }
        })
      ) {
        return Promise.resolve({
          FilterLink: "?",
          FirstRow: 1,
          FolderPermissions: "0x7fffffffffffffff",
          ForceNoHierarchy: 1,
          HierarchyHasIndention: null,
          LastRow: 5,
          Row: [{
            "ID": "25",
            "PermMask": "0x7fffffffffffffff",
            "FSObjType": "0",
            "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3",
            "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/25_.000",
            "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F25%5F%2E000",
            "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/25_.000",
            "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/25_.000",
            "ItemChildCount": "0",
            "FolderChildCount": "0",
            "SMTotalSize": "494",
            "Title": "North",
            "SiteUrl": "https://contoso.sharepoint.com/sites/north",
            "HubSiteId": "{389D0D83-40BB-40AD-B92A-534B7CB37D0B}",
            "TimeDeleted": "",
            "State": "",
            "State.": ""
          }, {
            "ID": "28",
            "PermMask": "0x7fffffffffffffff",
            "FSObjType": "0",
            "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3",
            "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/28_.000",
            "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F28%5F%2E000",
            "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/28_.000",
            "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/28_.000",
            "ItemChildCount": "0",
            "FolderChildCount": "0",
            "SMTotalSize": "526",
            "Title": "South",
            "SiteUrl": "https://contoso.sharepoint.com/sites/south",
            "HubSiteId": "{389D0D83-40BB-40AD-B92A-534B7CB37D0B}",
            "TimeDeleted": "",
            "State": "",
            "State.": ""
          }, {
            "ID": "29",
            "PermMask": "0x7fffffffffffffff",
            "FSObjType": "0",
            "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3",
            "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/29_.000",
            "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F29%5F%2E000",
            "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/29_.000",
            "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/29_.000",
            "ItemChildCount": "0",
            "FolderChildCount": "0",
            "SMTotalSize": "494",
            "Title": "Europe",
            "SiteUrl": "https://contoso.sharepoint.com/sites/europe",
            "HubSiteId": "{B2C94CA1-0957-4BDD-B549-B7D365EDC10F}",
            "TimeDeleted": "",
            "State": "",
            "State.": ""
          }, {
            "ID": "27",
            "PermMask": "0x7fffffffffffffff",
            "FSObjType": "0",
            "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3",
            "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/27_.000",
            "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F27%5F%2E000",
            "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/27_.000",
            "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/27_.000",
            "ItemChildCount": "0",
            "FolderChildCount": "0",
            "SMTotalSize": "526",
            "Title": "Asia",
            "SiteUrl": "https://contoso.sharepoint.com/sites/asia",
            "HubSiteId": "{B2C94CA1-0957-4BDD-B549-B7D365EDC10F}",
            "TimeDeleted": "",
            "State": "",
            "State.": ""
          }, {
            "ID": "24",
            "PermMask": "0x7fffffffffffffff",
            "FSObjType": "0",
            "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3",
            "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/24_.000",
            "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F24%5F%2E000",
            "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/24_.000",
            "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/24_.000",
            "ItemChildCount": "0",
            "FolderChildCount": "0",
            "SMTotalSize": "490",
            "Title": "America",
            "SiteUrl": "https://contoso.sharepoint.com/sites/america",
            "HubSiteId": "{B2C94CA1-0957-4BDD-B549-B7D365EDC10F}",
            "TimeDeleted": "",
            "State": "",
            "State.": ""
          }],
          RowLimit: 100
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, includeAssociatedSites: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
            "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
            "Title": "Sales"
          },
          {
            "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
            "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
            "Title": "Travel Programs"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists hub sites, including associated sites, with all properties for JSON output', (done) => {
    sinon.stub(request, 'get').resolves({
      value: [
        {
          "Description": null,
          "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Sales"
        },
        {
          "Description": null,
          "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
          "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Travel Programs"
        }
      ]
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/RenderListDataAsStream`) > -1
        && JSON.stringify(opts.body) === JSON.stringify({
          parameters: {
            ViewXml: "<View><Query><Where><And><And><IsNull><FieldRef Name=\"TimeDeleted\"/></IsNull><Neq><FieldRef Name=\"State\"/><Value Type='Integer'>0</Value></Neq></And><Neq><FieldRef Name=\"HubSiteId\"/><Value Type='Text'>{00000000-0000-0000-0000-000000000000}</Value></Neq></And></Where><OrderBy><FieldRef Name='Title' Ascending='true' /></OrderBy></Query><ViewFields><FieldRef Name=\"Title\"/><FieldRef Name=\"SiteUrl\"/><FieldRef Name=\"SiteId\"/><FieldRef Name=\"HubSiteId\"/></ViewFields><RowLimit Paged=\"TRUE\">30</RowLimit></View>",
            DatesInUtc: true
          }
        })
      ) {
        return Promise.resolve({
          FilterLink: "?",
          FirstRow: 1,
          FolderPermissions: "0x7fffffffffffffff",
          ForceNoHierarchy: 1,
          HierarchyHasIndention: null,
          LastRow: 5,
          Row: [{
            "ID": "25",
            "PermMask": "0x7fffffffffffffff",
            "FSObjType": "0",
            "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3",
            "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/25_.000",
            "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F25%5F%2E000",
            "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/25_.000",
            "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/25_.000",
            "ItemChildCount": "0",
            "FolderChildCount": "0",
            "SMTotalSize": "494",
            "Title": "North",
            "SiteUrl": "https://contoso.sharepoint.com/sites/north",
            "HubSiteId": "{389D0D83-40BB-40AD-B92A-534B7CB37D0B}",
            "TimeDeleted": "",
            "State": "",
            "State.": ""
          }, {
            "ID": "28",
            "PermMask": "0x7fffffffffffffff",
            "FSObjType": "0",
            "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3",
            "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/28_.000",
            "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F28%5F%2E000",
            "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/28_.000",
            "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/28_.000",
            "ItemChildCount": "0",
            "FolderChildCount": "0",
            "SMTotalSize": "526",
            "Title": "South",
            "SiteUrl": "https://contoso.sharepoint.com/sites/south",
            "HubSiteId": "{389D0D83-40BB-40AD-B92A-534B7CB37D0B}",
            "TimeDeleted": "",
            "State": "",
            "State.": ""
          }, {
            "ID": "29",
            "PermMask": "0x7fffffffffffffff",
            "FSObjType": "0",
            "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3",
            "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/29_.000",
            "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F29%5F%2E000",
            "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/29_.000",
            "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/29_.000",
            "ItemChildCount": "0",
            "FolderChildCount": "0",
            "SMTotalSize": "494",
            "Title": "Europe",
            "SiteUrl": "https://contoso.sharepoint.com/sites/europe",
            "HubSiteId": "{B2C94CA1-0957-4BDD-B549-B7D365EDC10F}",
            "TimeDeleted": "",
            "State": "",
            "State.": ""
          }, {
            "ID": "27",
            "PermMask": "0x7fffffffffffffff",
            "FSObjType": "0",
            "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3",
            "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/27_.000",
            "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F27%5F%2E000",
            "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/27_.000",
            "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/27_.000",
            "ItemChildCount": "0",
            "FolderChildCount": "0",
            "SMTotalSize": "526",
            "Title": "Asia",
            "SiteUrl": "https://contoso.sharepoint.com/sites/asia",
            "HubSiteId": "{B2C94CA1-0957-4BDD-B549-B7D365EDC10F}",
            "TimeDeleted": "",
            "State": "",
            "State.": ""
          }, {
            "ID": "24",
            "PermMask": "0x7fffffffffffffff",
            "FSObjType": "0",
            "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3",
            "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/24_.000",
            "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F24%5F%2E000",
            "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/24_.000",
            "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/24_.000",
            "ItemChildCount": "0",
            "FolderChildCount": "0",
            "SMTotalSize": "490",
            "Title": "America",
            "SiteUrl": "https://contoso.sharepoint.com/sites/america",
            "HubSiteId": "{B2C94CA1-0957-4BDD-B549-B7D365EDC10F}",
            "TimeDeleted": "",
            "State": "",
            "State.": ""
          }],
          RowLimit: 100
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, includeAssociatedSites: true, output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "Description": null,
            "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
            "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
            "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
            "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
            "Targets": null,
            "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
            "Title": "Sales",
            "AssociatedSites": [
              {
                "Title": "North",
                "SiteUrl": "https://contoso.sharepoint.com/sites/north"
              }
              , {
                "Title": "South",
                "SiteUrl": "https://contoso.sharepoint.com/sites/south"
              }
            ]
          },
          {
            "Description": null,
            "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
            "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
            "SiteId": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
            "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
            "Targets": null,
            "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
            "Title": "Travel Programs",
            "AssociatedSites": [
              {
                "Title": "Europe",
                "SiteUrl": "https://contoso.sharepoint.com/sites/europe"
              },
              {
                "Title": "Asia",
                "SiteUrl": "https://contoso.sharepoint.com/sites/asia"
              },
              {
                "Title": "America",
                "SiteUrl": "https://contoso.sharepoint.com/sites/america"
              }
            ]
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly retrieves the associated sites in batches', (done) => {
    // Cast the command class instance to any so we can set the private
    // property 'batchSize' to a small number for easier testing
    const newBatchSize = 3;
    (command as any).batchSize = newBatchSize;
    let firstPagedRequest: boolean = false;
    let secondPagedRequest: boolean = false;
    let thirdPagedRequest: boolean = false;
    sinon.stub(request, 'get').resolves({
      value: [
        {
          "Description": null,
          "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Sales"
        },
        {
          "Description": null,
          "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
          "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Travel Programs"
        }
      ]
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/RenderListDataAsStream`) > -1
        && opts.body.parameters.ViewXml.indexOf('<RowLimit Paged="TRUE">' + newBatchSize + '</RowLimit>') > -1) {
        if ((opts.url as string).indexOf('?Paged=TRUE') == -1) {
          firstPagedRequest = true;
          return Promise.resolve({
            FilterLink: "?",
            FirstRow: 1,
            FolderPermissions: "0x7fffffffffffffff",
            ForceNoHierarchy: 1,
            HierarchyHasIndention: null,
            LastRow: 3,
            NextHref: "?Paged=TRUE&p_Title=Another%20Hub%20Sub%202&p_ID=32&PageFirstRow=4&View=00000000-0000-0000-0000-00000000000",
            Row: [{ "ID": "30", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/30_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F30%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/30_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/30_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "554", "Title": "Another Hub Root", "SiteUrl": "https://bloemium.sharepoint.com/sites/AnotherHubRoot", "SiteId": "{E9E2090B-1F51-47EA-8466-75D5A244217E}", "HubSiteId": "{E9E2090B-1F51-47EA-8466-75D5A244217E}", "TimeDeleted": "", "State": "", "State.": "" }, { "ID": "31", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/31_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F31%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/31_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/31_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "556", "Title": "Another Hub Sub 1", "SiteUrl": "https://bloemium.sharepoint.com/sites/AnotherHubSub1", "SiteId": "{3A569D44-D3CD-45AB-9AB8-87675D18AF63}", "HubSiteId": "{E9E2090B-1F51-47EA-8466-75D5A244217E}", "TimeDeleted": "", "State": "", "State.": "" }, { "ID": "32", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/32_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F32%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/32_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/32_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "556", "Title": "Another Hub Sub 2", "SiteUrl": "https://bloemium.sharepoint.com/sites/AnotherHubSub2", "SiteId": "{794FE8EC-458F-444B-A799-E179AB786784}", "HubSiteId": "{E9E2090B-1F51-47EA-8466-75D5A244217E}", "TimeDeleted": "", "State": "", "State.": "" }],
            RowLimit: 3
          });
        }
        if ((opts.url as string).indexOf('?Paged=TRUE&p_Title=Another%20Hub%20Sub%202&p_ID=32&PageFirstRow=4&View=00000000-0000-0000-0000-00000000000') > -1) {
          secondPagedRequest = true
          return Promise.resolve({
            FilterLink: "?",
            FirstRow: 4,
            FolderPermissions: "0x7fffffffffffffff",
            ForceNoHierarchy: 1,
            HierarchyHasIndention: null,
            LastRow: 6,
            NextHref: "?Paged=TRUE&p_Title=Hub%20sub%204&p_ID=29&PageFirstRow=7&View=00000000-0000-0000-0000-00000000000",
            PrevHref: "?&&p_Title=Hub%20sub%201&&PageFirstRow=1&View=00000000-0000-0000-0000-000000000000",
            Row: [{ "ID": "25", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/25_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F25%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/25_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/25_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "518", "Title": "Hub sub 1", "SiteUrl": "https://bloemium.sharepoint.com/sites/Hubsub1", "SiteId": "{83C2E5B0-DC64-4040-AB1F-A6A9A8169E46}", "HubSiteId": "{77F50C57-C40A-4666-83F5-D325567512BE}", "TimeDeleted": "", "State": "", "State.": "" }, { "ID": "28", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/28_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F28%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/28_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/28_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "550", "Title": "Hub sub 3", "SiteUrl": "https://bloemium.sharepoint.com/sites/Hubsub3", "SiteId": "{5509F9AC-ECF8-488A-B960-BEDF4D8FB321}", "HubSiteId": "{77F50C57-C40A-4666-83F5-D325567512BE}", "TimeDeleted": "", "State": "", "State.": "" }, { "ID": "29", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/29_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F29%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/29_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/29_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "518", "Title": "Hub sub 4", "SiteUrl": "https://bloemium.sharepoint.com/sites/Hubsub4", "SiteId": "{8AC9E1ED-29B8-4342-AF30-11F597731F8A}", "HubSiteId": "{77F50C57-C40A-4666-83F5-D325567512BE}", "TimeDeleted": "", "State": "", "State.": "" }],
            RowLimit: 3
          });
        }
        if ((opts.url as string).indexOf('?Paged=TRUE&p_Title=Hub%20sub%204&p_ID=29&PageFirstRow=7&View=00000000-0000-0000-0000-00000000000') > -1) {
          thirdPagedRequest = true;
          return Promise.resolve({
            FilterLink: "?",
            FirstRow: 7,
            FolderPermissions: "0x7fffffffffffffff",
            ForceNoHierarchy: 1,
            HierarchyHasIndention: null,
            LastRow: 8,
            PrevHref: "?Paged=TRUE&PagedPrev=TRUE&p_Title=Hub%20sub%20x&p_ID=27&PageFirstRow=4&View=00000000-0000-0000-0000-000000000000",
            Row: [{ "ID": "27", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/27_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F27%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/27_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/27_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "550", "Title": "Hub sub x", "SiteUrl": "https://bloemium.sharepoint.com/sites/Hubsubx", "SiteId": "{DC0D0D79-1B0D-45A7-A8EE-7B97679B79DE}", "HubSiteId": "{77F50C57-C40A-4666-83F5-D325567512BE}", "TimeDeleted": "", "State": "", "State.": "" }, { "ID": "24", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/24_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F24%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/24_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/24_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "514", "Title": "Root Hub", "SiteUrl": "https://bloemium.sharepoint.com/sites/RootHub", "SiteId": "{77F50C57-C40A-4666-83F5-D325567512BE}", "HubSiteId": "{77F50C57-C40A-4666-83F5-D325567512BE}", "TimeDeleted": "", "State": "", "State.": "" }],
            RowLimit: 3
          });
        }
        return Promise.reject('Invalid request');
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, includeAssociatedSites: true, output: 'json' } }, () => {
      try {
        assert.strictEqual((firstPagedRequest && secondPagedRequest && thirdPagedRequest), true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly retrieves the associated sites in batches (debug)', (done) => {
    // Cast the command class instance to any so we can set the private
    // property 'batchSize' to a small number for easier testing
    const newBatchSize = 3;
    (command as any).batchSize = newBatchSize;
    let firstPagedRequest: boolean = false;
    let secondPagedRequest: boolean = false;
    let thirdPagedRequest: boolean = false;
    sinon.stub(request, 'get').resolves({
      value: [
        {
          "Description": null,
          "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Sales"
        },
        {
          "Description": null,
          "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
          "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Travel Programs"
        }
      ]
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/RenderListDataAsStream`) > -1
        && opts.body.parameters.ViewXml.indexOf('<RowLimit Paged="TRUE">' + newBatchSize + '</RowLimit>') > -1) {
        if ((opts.url as string).indexOf('?Paged=TRUE') == -1) {
          firstPagedRequest = true;
          return Promise.resolve({
            FilterLink: "?",
            FirstRow: 1,
            FolderPermissions: "0x7fffffffffffffff",
            ForceNoHierarchy: 1,
            HierarchyHasIndention: null,
            LastRow: 3,
            NextHref: "?Paged=TRUE&p_Title=Another%20Hub%20Sub%202&p_ID=32&PageFirstRow=4&View=00000000-0000-0000-0000-00000000000",
            Row: [{ "ID": "30", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/30_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F30%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/30_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/30_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "554", "Title": "Another Hub Root", "SiteUrl": "https://bloemium.sharepoint.com/sites/AnotherHubRoot", "SiteId": "{E9E2090B-1F51-47EA-8466-75D5A244217E}", "HubSiteId": "{E9E2090B-1F51-47EA-8466-75D5A244217E}", "TimeDeleted": "", "State": "", "State.": "" }, { "ID": "31", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/31_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F31%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/31_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/31_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "556", "Title": "Another Hub Sub 1", "SiteUrl": "https://bloemium.sharepoint.com/sites/AnotherHubSub1", "SiteId": "{3A569D44-D3CD-45AB-9AB8-87675D18AF63}", "HubSiteId": "{E9E2090B-1F51-47EA-8466-75D5A244217E}", "TimeDeleted": "", "State": "", "State.": "" }, { "ID": "32", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/32_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F32%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/32_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/32_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "556", "Title": "Another Hub Sub 2", "SiteUrl": "https://bloemium.sharepoint.com/sites/AnotherHubSub2", "SiteId": "{794FE8EC-458F-444B-A799-E179AB786784}", "HubSiteId": "{E9E2090B-1F51-47EA-8466-75D5A244217E}", "TimeDeleted": "", "State": "", "State.": "" }],
            RowLimit: 3
          });
        }
        if ((opts.url as string).indexOf('?Paged=TRUE&p_Title=Another%20Hub%20Sub%202&p_ID=32&PageFirstRow=4&View=00000000-0000-0000-0000-00000000000') > -1) {
          secondPagedRequest = true
          return Promise.resolve({
            FilterLink: "?",
            FirstRow: 4,
            FolderPermissions: "0x7fffffffffffffff",
            ForceNoHierarchy: 1,
            HierarchyHasIndention: null,
            LastRow: 6,
            NextHref: "?Paged=TRUE&p_Title=Hub%20sub%204&p_ID=29&PageFirstRow=7&View=00000000-0000-0000-0000-00000000000",
            PrevHref: "?&&p_Title=Hub%20sub%201&&PageFirstRow=1&View=00000000-0000-0000-0000-000000000000",
            Row: [{ "ID": "25", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/25_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F25%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/25_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/25_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "518", "Title": "Hub sub 1", "SiteUrl": "https://bloemium.sharepoint.com/sites/Hubsub1", "SiteId": "{83C2E5B0-DC64-4040-AB1F-A6A9A8169E46}", "HubSiteId": "{77F50C57-C40A-4666-83F5-D325567512BE}", "TimeDeleted": "", "State": "", "State.": "" }, { "ID": "28", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/28_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F28%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/28_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/28_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "550", "Title": "Hub sub 3", "SiteUrl": "https://bloemium.sharepoint.com/sites/Hubsub3", "SiteId": "{5509F9AC-ECF8-488A-B960-BEDF4D8FB321}", "HubSiteId": "{77F50C57-C40A-4666-83F5-D325567512BE}", "TimeDeleted": "", "State": "", "State.": "" }, { "ID": "29", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/29_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F29%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/29_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/29_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "518", "Title": "Hub sub 4", "SiteUrl": "https://bloemium.sharepoint.com/sites/Hubsub4", "SiteId": "{8AC9E1ED-29B8-4342-AF30-11F597731F8A}", "HubSiteId": "{77F50C57-C40A-4666-83F5-D325567512BE}", "TimeDeleted": "", "State": "", "State.": "" }],
            RowLimit: 3
          });
        }
        if ((opts.url as string).indexOf('?Paged=TRUE&p_Title=Hub%20sub%204&p_ID=29&PageFirstRow=7&View=00000000-0000-0000-0000-00000000000') > -1) {
          thirdPagedRequest = true;
          return Promise.resolve({
            FilterLink: "?",
            FirstRow: 7,
            FolderPermissions: "0x7fffffffffffffff",
            ForceNoHierarchy: 1,
            HierarchyHasIndention: null,
            LastRow: 8,
            PrevHref: "?Paged=TRUE&PagedPrev=TRUE&p_Title=Hub%20sub%20x&p_ID=27&PageFirstRow=4&View=00000000-0000-0000-0000-000000000000",
            Row: [{ "ID": "27", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/27_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F27%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/27_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/27_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "550", "Title": "Hub sub x", "SiteUrl": "https://bloemium.sharepoint.com/sites/Hubsubx", "SiteId": "{DC0D0D79-1B0D-45A7-A8EE-7B97679B79DE}", "HubSiteId": "{77F50C57-C40A-4666-83F5-D325567512BE}", "TimeDeleted": "", "State": "", "State.": "" }, { "ID": "24", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/24_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F24%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/24_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/24_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "514", "Title": "Root Hub", "SiteUrl": "https://bloemium.sharepoint.com/sites/RootHub", "SiteId": "{77F50C57-C40A-4666-83F5-D325567512BE}", "HubSiteId": "{77F50C57-C40A-4666-83F5-D325567512BE}", "TimeDeleted": "", "State": "", "State.": "" }],
            RowLimit: 3
          });
        }
        return Promise.reject('Invalid request');
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, includeAssociatedSites: true, output: 'json' } }, () => {
      try {
        assert.strictEqual((firstPagedRequest && secondPagedRequest && thirdPagedRequest), true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles empty result while retrieving associated sites in batches', (done) => {
    // Cast the command class instance to any so we can set the private
    // property 'batchSize' to a small number for easier testing
    const newBatchSize = 3;
    (command as any).batchSize = newBatchSize;
    let firstPagedRequest: boolean = false;
    sinon.stub(request, 'get').resolves({
      value: [
        {
          "Description": null,
          "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Sales"
        },
        {
          "Description": null,
          "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
          "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Travel Programs"
        }
      ]
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/RenderListDataAsStream`) > -1
        && opts.body.parameters.ViewXml.indexOf('<RowLimit Paged="TRUE">' + newBatchSize + '</RowLimit>') > -1) {
        if ((opts.url as string).indexOf('?Paged=TRUE') == -1) {
          firstPagedRequest = true;
          return Promise.resolve({
            FilterLink: "?",
            FirstRow: 1,
            FolderPermissions: "0x7fffffffffffffff",
            ForceNoHierarchy: 1,
            HierarchyHasIndention: null,
            LastRow: 0,
            Row: [],
            RowLimit: 3
          });
        }
        return Promise.reject('Invalid request');
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, includeAssociatedSites: true, output: 'json' } }, () => {
      try {
        assert.strictEqual(firstPagedRequest, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when retrieving hub sites', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error in the first batch of associated sites', (done) => {
    // Cast the command class instance to any so we can set the private
    // property 'batchSize' to a small number for easier testing
    const newBatchSize = 3;
    (command as any).batchSize = newBatchSize;
    sinon.stub(request, 'get').resolves({
      value: [
        {
          "Description": null,
          "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Sales"
        },
        {
          "Description": null,
          "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
          "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Travel Programs"
        }
      ]
    });
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('An error has occurred');
    });
    cmdInstance.action({ options: { debug: false, includeAssociatedSites: true, output: 'json' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error in a subsequent batch of associated sites', (done) => {
    // Cast the command class instance to any so we can set the private
    // property 'batchSize' to a small number for easier testing
    const newBatchSize = 3;
    (command as any).batchSize = newBatchSize;
    sinon.stub(request, 'get').resolves({
      value: [
        {
          "Description": null,
          "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
          "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Sales"
        },
        {
          "Description": null,
          "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
          "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
          "SiteId": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
          "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
          "Targets": null,
          "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
          "Title": "Travel Programs"
        }
      ]
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/RenderListDataAsStream`) > -1
        && opts.body.parameters.ViewXml.indexOf('<RowLimit Paged="TRUE">' + newBatchSize + '</RowLimit>') > -1) {
        if ((opts.url as string).indexOf('?Paged=TRUE') == -1) {
          return Promise.resolve({
            FilterLink: "?",
            FirstRow: 1,
            FolderPermissions: "0x7fffffffffffffff",
            ForceNoHierarchy: 1,
            HierarchyHasIndention: null,
            LastRow: 3,
            NextHref: "?Paged=TRUE&p_Title=Another%20Hub%20Sub%202&p_ID=32&PageFirstRow=4&View=00000000-0000-0000-0000-00000000000",
            Row: [{ "ID": "30", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/30_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F30%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/30_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/30_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "554", "Title": "Another Hub Root", "SiteUrl": "https://bloemium.sharepoint.com/sites/AnotherHubRoot", "SiteId": "{E9E2090B-1F51-47EA-8466-75D5A244217E}", "HubSiteId": "{E9E2090B-1F51-47EA-8466-75D5A244217E}", "TimeDeleted": "", "State": "", "State.": "" }, { "ID": "31", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/31_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F31%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/31_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/31_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "556", "Title": "Another Hub Sub 1", "SiteUrl": "https://bloemium.sharepoint.com/sites/AnotherHubSub1", "SiteId": "{3A569D44-D3CD-45AB-9AB8-87675D18AF63}", "HubSiteId": "{E9E2090B-1F51-47EA-8466-75D5A244217E}", "TimeDeleted": "", "State": "", "State.": "" }, { "ID": "32", "PermMask": "0x7fffffffffffffff", "FSObjType": "0", "ContentTypeId": "0x0100F14AFE642BCF6347882B6B8ABA3E15E3", "FileRef": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/32_.000", "FileRef.urlencode": "%2FLists%2FDO%5FNOT%5FDELETE%5FSPLIST%5FTENANTADMIN%5FAGGREGATED%5FSITECO%2F32%5F%2E000", "FileRef.urlencodeasurl": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/32_.000", "FileRef.urlencoding": "/Lists/DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECO/32_.000", "ItemChildCount": "0", "FolderChildCount": "0", "SMTotalSize": "556", "Title": "Another Hub Sub 2", "SiteUrl": "https://bloemium.sharepoint.com/sites/AnotherHubSub2", "SiteId": "{794FE8EC-458F-444B-A799-E179AB786784}", "HubSiteId": "{E9E2090B-1F51-47EA-8466-75D5A244217E}", "TimeDeleted": "", "State": "", "State.": "" }],
            RowLimit: 3
          });
        }
        if ((opts.url as string).indexOf('?Paged=TRUE&p_Title=Another%20Hub%20Sub%202&p_ID=32&PageFirstRow=4&View=00000000-0000-0000-0000-00000000000') > -1) {
          return Promise.reject('An error has occurred');
        }

        return Promise.reject('Invalid request');
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, includeAssociatedSites: true, output: 'json' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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