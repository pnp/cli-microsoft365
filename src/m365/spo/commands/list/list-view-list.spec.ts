import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { urlUtil } from '../../../../utils/urlUtil';
import { formatting } from '../../../../utils/formatting';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./list-view-list');

describe(commands.LIST_VIEW_LIST, () => {
  //#region Mocked Responses
  const listViewResponse = {
    "value": [{ "Aggregations": null, "AggregationsStatus": null, "BaseViewId": "1", "ColumnWidth": null, "ContentTypeId": { "StringValue": "0x" }, "CustomFormatter": null, "DefaultView": true, "DefaultViewForContentType": false, "EditorModified": false, "Formats": null, "Hidden": false, "HtmlSchemaXml": "<View Name=\"{85907FB0-CDA1-43EC-832E-27118C89CF9F}\" DefaultView=\"TRUE\" MobileView=\"TRUE\" MobileDefaultView=\"TRUE\" Type=\"HTML\" DisplayName=\"All Documents\" Url=\"/sites/ninja/Shared Documents/Forms/AllItems.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/dlicon.png?rev=45\"><XslLink Default=\"TRUE\">main.xsl</XslLink><JSLink>clienttemplates.js</JSLink><RowLimit Paged=\"TRUE\">30</RowLimit><Toolbar Type=\"Standard\" /><ViewFields><FieldRef Name=\"DocIcon\" /><FieldRef Name=\"LinkFilename\" /><FieldRef Name=\"Modified\" /><FieldRef Name=\"Editor\" /></ViewFields><ParameterBindings><ParameterBinding Name=\"NoAnnouncements\" Location=\"Resource(wss,noitemsinview_doclibrary)\" /><ParameterBinding Name=\"NoAnnouncementsHowTo\" Location=\"Resource(wss,noitemsinview_doclibrary_howto2)\" /></ParameterBindings><Query><OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy></Query></View>", "Id": "85907fb0-cda1-43ec-832e-27118c89cf9f", "ImageUrl": "/_layouts/15/images/dlicon.png?rev=45", "IncludeRootFolder": false, "ViewJoins": null, "JSLink": "clienttemplates.js", "ListViewXml": "<View Name=\"{85907FB0-CDA1-43EC-832E-27118C89CF9F}\" DefaultView=\"TRUE\" MobileView=\"TRUE\" MobileDefaultView=\"TRUE\" Type=\"HTML\" DisplayName=\"All Documents\" Url=\"/sites/ninja/Shared Documents/Forms/AllItems.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/dlicon.png?rev=45\" ><Query><OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy></Query><ViewFields><FieldRef Name=\"DocIcon\" /><FieldRef Name=\"LinkFilename\" /><FieldRef Name=\"Modified\" /><FieldRef Name=\"Editor\" /></ViewFields><RowLimit Paged=\"TRUE\">30</RowLimit><JSLink>clienttemplates.js</JSLink><XslLink Default=\"TRUE\">main.xsl</XslLink><Toolbar Type=\"Standard\"/></View>", "Method": null, "MobileDefaultView": true, "MobileView": true, "ModerationType": null, "NewDocumentTemplates": null, "OrderedView": false, "Paged": true, "PersonalView": false, "ViewProjectedFields": null, "ViewQuery": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>", "ReadOnlyView": false, "RequiresClientIntegration": false, "RowLimit": 30, "Scope": 0, "ServerRelativePath": { "DecodedUrl": "/sites/ninja/Shared Documents/Forms/AllItems.aspx" }, "ServerRelativeUrl": "/sites/ninja/Shared Documents/Forms/AllItems.aspx", "StyleId": null, "TabularView": true, "Threaded": false, "Title": "All Documents", "Toolbar": "", "ToolbarTemplateName": null, "ViewType": "HTML", "ViewData": null, "VisualizationInfo": null }, { "Aggregations": null, "AggregationsStatus": null, "BaseViewId": "9", "ColumnWidth": null, "ContentTypeId": { "StringValue": "0x" }, "CustomFormatter": null, "DefaultView": false, "DefaultViewForContentType": false, "EditorModified": false, "Formats": null, "Hidden": true, "HtmlSchemaXml": "<View Name=\"{281E80FD-099B-4EBA-B622-A94FA03BF865}\" Type=\"HTML\" Hidden=\"TRUE\" TabularView=\"FALSE\" DisplayName=\"Relink Documents\" Url=\"/sites/ninja/Shared Documents/Forms/repair.aspx\" Level=\"1\" BaseViewID=\"9\" ContentTypeID=\"0x\" ToolbarTemplate=\"RelinkToolBar\" ImageUrl=\"/_layouts/15/images/dlicon.png?rev=45\"><XslLink Default=\"TRUE\">main.xsl</XslLink><RowLimit Paged=\"TRUE\">30</RowLimit><Toolbar Type=\"Standard\" /><ViewFields><FieldRef Name=\"DocIcon\" /><FieldRef Name=\"LinkFilenameNoMenu\" /><FieldRef Name=\"RepairDocument\" /><FieldRef Name=\"Modified\" /><FieldRef Name=\"Editor\" /><FieldRef Name=\"ContentType\" /><FieldRef Name=\"TemplateUrl\" /></ViewFields><ParameterBindings><ParameterBinding Name=\"NoAnnouncements\" Location=\"Resource(wss,noitemsinview_doclibrary)\" /><ParameterBinding Name=\"NoAnnouncementsHowTo\" Location=\"Resource(wss,noitemsinview_doclibrary_howto2)\" /></ParameterBindings><Query><OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy><Where><Neq><FieldRef Name=\"xd_Signature\" /><Value Type=\"Boolean\">1</Value></Neq></Where></Query></View>", "Id": "281e80fd-099b-4eba-b622-a94fa03bf865", "ImageUrl": "/_layouts/15/images/dlicon.png?rev=45", "IncludeRootFolder": false, "ViewJoins": null, "JSLink": null, "ListViewXml": "<View Name=\"{281E80FD-099B-4EBA-B622-A94FA03BF865}\" Type=\"HTML\" Hidden=\"TRUE\" TabularView=\"FALSE\" DisplayName=\"Relink Documents\" Url=\"/sites/ninja/Shared Documents/Forms/repair.aspx\" Level=\"1\" BaseViewID=\"9\" ContentTypeID=\"0x\" ToolbarTemplate=\"RelinkToolBar\" ImageUrl=\"/_layouts/15/images/dlicon.png?rev=45\" ><Query><OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy><Where><Neq><FieldRef Name=\"xd_Signature\" /><Value Type=\"Boolean\">1</Value></Neq></Where></Query><ViewFields><FieldRef Name=\"DocIcon\" /><FieldRef Name=\"LinkFilenameNoMenu\" /><FieldRef Name=\"RepairDocument\" /><FieldRef Name=\"Modified\" /><FieldRef Name=\"Editor\"/><FieldRef Name=\"ContentType\" /><FieldRef Name=\"TemplateUrl\" /></ViewFields><RowLimit Paged=\"TRUE\">30</RowLimit><XslLink Default=\"TRUE\">main.xsl</XslLink><Toolbar Type=\"Standard\"/></View>", "Method": null, "MobileDefaultView": false, "MobileView": false, "ModerationType": null, "NewDocumentTemplates": null, "OrderedView": false, "Paged": true, "PersonalView": false, "ViewProjectedFields": null, "ViewQuery": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy><Where><Neq><FieldRef Name=\"xd_Signature\" /><Value Type=\"Boolean\">1</Value></Neq></Where>", "ReadOnlyView": false, "RequiresClientIntegration": false, "RowLimit": 30, "Scope": 0, "ServerRelativePath": { "DecodedUrl": "/sites/ninja/Shared Documents/Forms/repair.aspx" }, "ServerRelativeUrl": "/sites/ninja/Shared Documents/Forms/repair.aspx", "StyleId": null, "TabularView": false, "Threaded": false, "Title": "Relink Documents", "Toolbar": "", "ToolbarTemplateName": "RelinkToolBar", "ViewType": "HTML", "ViewData": null, "VisualizationInfo": null }, { "Aggregations": null, "AggregationsStatus": null, "BaseViewId": "40", "ColumnWidth": null, "ContentTypeId": { "StringValue": "0x" }, "CustomFormatter": null, "DefaultView": false, "DefaultViewForContentType": false, "EditorModified": false, "Formats": null, "Hidden": true, "HtmlSchemaXml": "<View Name=\"{3D760127-982C-405E-9C93-E1F76E1A1110}\" Type=\"HTML\" Hidden=\"TRUE\" DisplayName=\"assetLibTemp\" Url=\"/sites/ninja/Shared Documents/Forms/Thumbnails.aspx\" Level=\"1\" BaseViewID=\"40\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/dlicon.png?rev=45\"><ViewFields><FieldRef Name=\"LinkFilename\" /></ViewFields><RowLimit>20</RowLimit><Query><OrderBy><FieldRef Name=\"LinkFilename\" /></OrderBy></Query></View>", "Id": "3d760127-982c-405e-9c93-e1f76e1a1110", "ImageUrl": "/_layouts/15/images/dlicon.png?rev=45", "IncludeRootFolder": false, "ViewJoins": null, "JSLink": null, "ListViewXml": "<View Name=\"{3D760127-982C-405E-9C93-E1F76E1A1110}\" Type=\"HTML\" Hidden=\"TRUE\" DisplayName=\"assetLibTemp\" Url=\"/sites/ninja/SharedDocuments/Forms/Thumbnails.aspx\" Level=\"1\" BaseViewID=\"40\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/dlicon.png?rev=45\" ><Query><OrderBy><FieldRef Name=\"LinkFilename\" /></OrderBy></Query><ViewFields><FieldRef Name=\"LinkFilename\" /></ViewFields><RowLimit>20</RowLimit><Toolbar Type=\"None\"/></View>", "Method": null, "MobileDefaultView": false, "MobileView": false, "ModerationType": null, "NewDocumentTemplates": null, "OrderedView": false, "Paged": false, "PersonalView": false, "ViewProjectedFields": null, "ViewQuery": "<OrderBy><FieldRef Name=\"LinkFilename\" /></OrderBy>", "ReadOnlyView": false, "RequiresClientIntegration": false, "RowLimit": 20, "Scope": 0, "ServerRelativePath": { "DecodedUrl": "/sites/ninja/Shared Documents/Forms/Thumbnails.aspx" }, "ServerRelativeUrl": "/sites/ninja/Shared Documents/Forms/Thumbnails.aspx", "StyleId": null, "TabularView": true, "Threaded": false, "Title": "assetLibTemp", "Toolbar": null, "ToolbarTemplateName": null, "ViewType": "HTML", "ViewData": null, "VisualizationInfo": null }, { "Aggregations": null, "AggregationsStatus": null, "BaseViewId": "7", "ColumnWidth": null, "ContentTypeId": { "StringValue": "0x" }, "CustomFormatter": null, "DefaultView": false, "DefaultViewForContentType": false, "EditorModified": false, "Formats": null, "Hidden": true, "HtmlSchemaXml": "<View Name=\"{1204C10A-978D-46C8-A6F2-65ECFA5E03D5}\" Type=\"HTML\" Hidden=\"TRUE\" TabularView=\"FALSE\" AggregateView=\"TRUE\" DisplayName=\"Merge Documents\" Url=\"/sites/ninja/Shared Documents/Forms/Combine.aspx\" Level=\"1\" BaseViewID=\"7\" ContentTypeID=\"0x\" ToolbarTemplate=\"MergeToolBar\" ImageUrl=\"/_layouts/15/images/dlicon.png?rev=45\"><XslLink Default=\"TRUE\">main.xsl</XslLink><JSLink>clienttemplates.js</JSLink><RowLimit Paged=\"TRUE\">30</RowLimit><Toolbar Type=\"Standard\" /><ViewFields><FieldRef Name=\"DocIcon\" /><FieldRef Name=\"LinkFilename\" /><FieldRef Name=\"Combine\" /><FieldRef Name=\"Modified\" /><FieldRef Name=\"Editor\" /></ViewFields><ParameterBindings><ParameterBinding Name=\"NoAnnouncements\" Location=\"Resource(wss,noitemsinview_doclibrary)\" /><ParameterBinding Name=\"NoAnnouncementsHowTo\" Location=\"Resource(wss,noitemsinview_doclibrary_howto2)\" /></ParameterBindings><Query><OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy></Query></View>", "Id": "1204c10a-978d-46c8-a6f2-65ecfa5e03d5", "ImageUrl": "/_layouts/15/images/dlicon.png?rev=45", "IncludeRootFolder": false, "ViewJoins": null, "JSLink": "clienttemplates.js", "ListViewXml": "<View Name=\"{1204C10A-978D-46C8-A6F2-65ECFA5E03D5}\" Type=\"HTML\" Hidden=\"TRUE\" TabularView=\"FALSE\" AggregateView=\"TRUE\" DisplayName=\"Merge Documents\" Url=\"/sites/ninja/Shared Documents/Forms/Combine.aspx\" Level=\"1\" BaseViewID=\"7\" ContentTypeID=\"0x\" ToolbarTemplate=\"MergeToolBar\" ImageUrl=\"/_layouts/15/images/dlicon.png?rev=45\" ><Query><OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy></Query><ViewFields><FieldRef Name=\"DocIcon\" /><FieldRef Name=\"LinkFilename\" /><FieldRef Name=\"Combine\" /><FieldRef Name=\"Modified\" /><FieldRef Name=\"Editor\" /></ViewFields><RowLimit Paged=\"TRUE\">30</RowLimit><JSLink>clienttemplates.js</JSLink><XslLink Default=\"TRUE\">main.xsl</XslLink><Toolbar Type=\"Standard\"/></View>", "Method": null, "MobileDefaultView": false, "MobileView": false, "ModerationType": null, "NewDocumentTemplates": null, "OrderedView": false, "Paged": true, "PersonalView": false, "ViewProjectedFields": null, "ViewQuery": "<OrderBy><FieldRef Name=\"FileLeafRef\" /></OrderBy>", "ReadOnlyView": false, "RequiresClientIntegration": false, "RowLimit": 30, "Scope": 0, "ServerRelativePath": { "DecodedUrl": "/sites/ninja/Shared Documents/Forms/Combine.aspx" }, "ServerRelativeUrl": "/sites/ninja/Shared Documents/Forms/Combine.aspx", "StyleId": null, "TabularView": false, "Threaded": false, "Title": "Merge Documents", "Toolbar": "", "ToolbarTemplateName": "MergeToolBar", "ViewType": "HTML", "ViewData": null, "VisualizationInfo": null }, { "Aggregations": null, "AggregationsStatus": null, "BaseViewId": "50", "ColumnWidth": null, "ContentTypeId": { "StringValue": "0x" }, "CustomFormatter": null, "DefaultView": false, "DefaultViewForContentType": false, "EditorModified": false, "Formats": null, "Hidden": true, "HtmlSchemaXml": "<View Name=\"{E021E923-0801-4A16-9775-545239356739}\" MobileView=\"TRUE\" Type=\"HTML\" Hidden=\"TRUE\" DisplayName=\"\" Url=\"/sites/ninja/SitePages/Home.aspx\" Level=\"1\" BaseViewID=\"50\" ContentTypeID=\"0x\"><XslLink Default=\"TRUE\">main.xsl</XslLink><JSLink>clienttemplates.js</JSLink><RowLimit Paged=\"TRUE\">15</RowLimit><Toolbar Type=\"Standard\" /><ViewFields><FieldRef Name=\"DocIcon\" /><FieldRef Name=\"LinkFilename\" /></ViewFields><ParameterBindings><ParameterBinding Name=\"NoAnnouncements\" Location=\"Resource(wss,noitemsinview_doclibrary)\" /><ParameterBinding Name=\"NoAnnouncementsHowTo\" Location=\"Resource(wss,noitemsinview_doclibrary_howto2)\" /></ParameterBindings><Query><OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy></Query></View>", "Id": "e021e923-0801-4a16-9775-545239356739", "ImageUrl": null, "IncludeRootFolder": false, "ViewJoins": null, "JSLink": "clienttemplates.js", "ListViewXml": "<View Name=\"{E021E923-0801-4A16-9775-545239356739}\" MobileView=\"TRUE\" Type=\"HTML\"Hidden=\"TRUE\" DisplayName=\"\" Url=\"/sites/ninja/SitePages/Home.aspx\" Level=\"1\" BaseViewID=\"50\" ContentTypeID=\"0x\" ><Query><OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy></Query><ViewFields><FieldRef Name=\"DocIcon\" /><FieldRef Name=\"LinkFilename\" /></ViewFields><RowLimit Paged=\"TRUE\">15</RowLimit><JSLink>clienttemplates.js</JSLink><XslLink Default=\"TRUE\">main.xsl</XslLink><Toolbar Type=\"Standard\"/></View>", "Method": null, "MobileDefaultView": false, "MobileView": true, "ModerationType": null, "NewDocumentTemplates": null, "OrderedView": false, "Paged": true, "PersonalView": false, "ViewProjectedFields": null, "ViewQuery": "<OrderBy><FieldRef Name=\"Modified\" Ascending=\"FALSE\" /></OrderBy>", "ReadOnlyView": false, "RequiresClientIntegration": false, "RowLimit": 15, "Scope": 0, "ServerRelativePath": { "DecodedUrl": "/sites/ninja/SitePages/Home.aspx" }, "ServerRelativeUrl": "/sites/ninja/SitePages/Home.aspx", "StyleId": null, "TabularView": true, "Threaded": false, "Title": "", "Toolbar": "", "ToolbarTemplateName": null, "ViewType": "HTML", "ViewData": null, "VisualizationInfo": null }]
  };
  //#endregion

  const listId = '1f187321-f086-4d3d-8523-517e94cc9df9';
  const webUrl = 'https://contoso.sharepoint.com/sites/ninja';
  const listTitle = 'Documents';
  const listUrl = '/sites/ninja/Shared Documents';

  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LIST_VIEW_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Id', 'Title', 'DefaultView', 'Hidden', 'BaseViewId']);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listId: listId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listId: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if valid options are specified', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, listTitle: listTitle } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves all views of the specific list if listTitle option is passed (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(listTitle)}')/views`) {
        return listViewResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        listTitle: listTitle,
        webUrl: webUrl
      }
    });
  });

  it('retrieves all views of the specific list if listId option is passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listId)}')/views`) {
        return listViewResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        listId: listId,
        webUrl: webUrl
      }
    });
  });

  it('retrieves all views of the specific list if listUrl option is passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const serverRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
      if (opts.url === `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')/views`) {
        return listViewResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        listUrl: listUrl,
        webUrl: webUrl
      }
    });
  });

  it('correctly handles error when the specified list doesn\'t exist', async () => {
    const errorMessage = `List '' does not exist at site with URL ''`;
    sinon.stub(request, 'get').callsFake(async () => {
      throw errorMessage;
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        listTitle: listTitle,
        webUrl: webUrl
      }
    }), new CommandError(errorMessage));
  });
});
