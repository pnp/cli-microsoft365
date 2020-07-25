import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError, CommandTypes } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./contenttype-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.CONTENTTYPE_GET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONTENTTYPE_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets information about a site content type', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.resolve({"Description":"Create a new list item.","DisplayFormTemplateName":"ListForm","DisplayFormUrl":"","DocumentTemplate":"","DocumentTemplateUrl":"","EditFormTemplateName":"ListForm","EditFormUrl":"","Group":"PnP Content Types","Hidden":false,"Id":{"StringValue":"0x0100558D85B7216F6A489A499DB361E1AE2F"},"JSLink":"","MobileDisplayFormUrl":"","MobileEditFormUrl":"","MobileNewFormUrl":"","Name":"PnP Alert","NewFormTemplateName":"ListForm","NewFormUrl":"","ReadOnly":false,"SchemaXml":"<ContentType ID=\"0x0100558D85B7216F6A489A499DB361E1AE2F\" Name=\"PnP Alert\" Group=\"PnP Content Types\" Description=\"Create a new list item.\" Version=\"1\"><Folder TargetName=\"_cts/PnP Alert\" /><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Name=\"ContentType\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" Type=\"Computed\" DisplayName=\"Content Type\" Sealed=\"TRUE\" Sortable=\"FALSE\" RenderXMLUsingPattern=\"TRUE\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" Customization=\"\"><FieldRefs><FieldRef ID=\"{03e45e84-1992-4d42-9116-26f756012634}\" Name=\"ContentTypeId\" /></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\" /></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Name=\"Title\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" Group=\"_Hidden\" Type=\"Text\" DisplayName=\"Title\" Required=\"TRUE\" FromBaseType=\"TRUE\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" Customization=\"\" ShowInNewForm=\"TRUE\" ShowInEditForm=\"TRUE\"></Field></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>","Scope":"/sites/portal","Sealed":false,"StringId":"0x0100558D85B7216F6A489A499DB361E1AE2F"});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '0x0100558D85B7216F6A489A499DB361E1AE2F' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({"Description":"Create a new list item.","DisplayFormTemplateName":"ListForm","DisplayFormUrl":"","DocumentTemplate":"","DocumentTemplateUrl":"","EditFormTemplateName":"ListForm","EditFormUrl":"","Group":"PnP Content Types","Hidden":false,"Id":{"StringValue":"0x0100558D85B7216F6A489A499DB361E1AE2F"},"JSLink":"","MobileDisplayFormUrl":"","MobileEditFormUrl":"","MobileNewFormUrl":"","Name":"PnP Alert","NewFormTemplateName":"ListForm","NewFormUrl":"","ReadOnly":false,"SchemaXml":"<ContentType ID=\"0x0100558D85B7216F6A489A499DB361E1AE2F\" Name=\"PnP Alert\" Group=\"PnP Content Types\" Description=\"Create a new list item.\" Version=\"1\"><Folder TargetName=\"_cts/PnP Alert\" /><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Name=\"ContentType\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" Type=\"Computed\" DisplayName=\"Content Type\" Sealed=\"TRUE\" Sortable=\"FALSE\" RenderXMLUsingPattern=\"TRUE\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" Customization=\"\"><FieldRefs><FieldRef ID=\"{03e45e84-1992-4d42-9116-26f756012634}\" Name=\"ContentTypeId\" /></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\" /></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Name=\"Title\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" Group=\"_Hidden\" Type=\"Text\" DisplayName=\"Title\" Required=\"TRUE\" FromBaseType=\"TRUE\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" Customization=\"\" ShowInNewForm=\"TRUE\" ShowInEditForm=\"TRUE\"></Field></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments></ContentType>","Scope":"/sites/portal","Sealed":false,"StringId":"0x0100558D85B7216F6A489A499DB361E1AE2F"}));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about a list content type', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Events')/contenttypes('0x010200973548ACFFDA0948BE80AF607C4E28F9')`) > -1) {
        return Promise.resolve({"Description":"Create a new meeting, deadline or other event.","DisplayFormTemplateName":"ListForm","DisplayFormUrl":"","DocumentTemplate":"","DocumentTemplateUrl":"","EditFormTemplateName":"ListForm","EditFormUrl":"","Group":"List Content Types","Hidden":false,"Id":{"StringValue":"0x010200973548ACFFDA0948BE80AF607C4E28F9"},"JSLink":"","MobileDisplayFormUrl":"","MobileEditFormUrl":"","MobileNewFormUrl":"","Name":"Event","NewFormTemplateName":"ListForm","NewFormUrl":"","ReadOnly":false,"SchemaXml":"<ContentType ID=\"0x010200973548ACFFDA0948BE80AF607C4E28F9\" Name=\"Event\" Group=\"List Content Types\" V2ListTemplateName=\"events\" Description=\"Create a new meeting, deadline or other event.\" Version=\"0\" FeatureId=\"{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}\"><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Type=\"Computed\" DisplayName=\"Content Type\" Name=\"ContentType\" DisplaceOnUpgrade=\"TRUE\" RenderXMLUsingPattern=\"TRUE\" Sortable=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ContentTypeId\"/></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\"/></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Type=\"Text\" Name=\"Title\" DisplayName=\"Title\" Required=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" FromBaseType=\"TRUE\" ColName=\"nvarchar1\" ShowInNewForm=\"TRUE\" ShowInEditForm=\"TRUE\"/><Field ID=\"{288f5f32-8462-4175-8f09-dd7ba29359a9}\" Type=\"Text\" Name=\"Location\" DisplayName=\"Location\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Location\" ColName=\"nvarchar4\"/><Field Type=\"DateTime\" ID=\"{64cd368d-2f95-4bfc-a1f9-8d4324ecb007}\" Name=\"EventDate\" DisplayName=\"Start Time\" Format=\"DateTime\" Sealed=\"TRUE\" Required=\"TRUE\" FromBaseType=\"TRUE\" Filterable=\"FALSE\" FilterableNoRecurrence=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"EventDate\" ColName=\"datetime1\"><Default>[today]</Default><FieldRefs><FieldRef Name=\"fAllDayEvent\" RefType=\"AllDayEvent\"/></FieldRefs></Field><Field ID=\"{2684f9f2-54be-429f-ba06-76754fc056bf}\" Type=\"DateTime\" Name=\"EndDate\" DisplayName=\"End Time\" Format=\"DateTime\" Sealed=\"TRUE\" Required=\"TRUE\" Filterable=\"FALSE\" FilterableNoRecurrence=\"TRUE\" Indexed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"EndDate\" ColName=\"datetime2\"><Default>[today]</Default><FieldRefs><FieldRef Name=\"fAllDayEvent\" RefType=\"AllDayEvent\"/></FieldRefs></Field><Field Type=\"Note\" ID=\"{9da97a8a-1da5-4a77-98d3-4bc10456e700}\" Name=\"Description\" RichText=\"TRUE\" RichTextMode=\"FullHtml\" DisplayName=\"Description\" Sortable=\"FALSE\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Description\" ColName=\"ntext2\"/><Field ID=\"{6df9bd52-550e-4a30-bc31-a4366832a87d}\" Name=\"Category\" DisplayName=\"Category\" Type=\"Choice\" Format=\"Dropdown\" FillInChoice=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Category\" ColName=\"nvarchar7\"><CHOICES><CHOICE>Meeting</CHOICE><CHOICE>Work hours</CHOICE><CHOICE>Business</CHOICE><CHOICE>Holiday</CHOICE><CHOICE>Get-together</CHOICE><CHOICE>Gifts</CHOICE><CHOICE>Birthday</CHOICE><CHOICE>Anniversary</CHOICE></CHOICES></Field><Field ID=\"{7d95d1f4-f5fd-4a70-90cd-b35abc9b5bc8}\" Type=\"AllDayEvent\" Name=\"fAllDayEvent\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"All Day Event\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"fAllDayEvent\" ColName=\"bit1\"><FieldRefs><FieldRef Name=\"EventDate\" RefType=\"StartDate\"/><FieldRef Name=\"EndDate\" RefType=\"EndDate\"/><FieldRef Name=\"TimeZone\" RefType=\"TimeZone\"/><FieldRef Name=\"XMLTZone\" RefType=\"XMLTZone\"/></FieldRefs></Field><Field ID=\"{f2e63656-135e-4f1c-8fc2-ccbe74071901}\" Type=\"Recurrence\" Name=\"fRecurrence\" DisplayName=\"Recurrence\" DisplayImage=\"recur.gif\" ExceptionImage=\"recurEx.gif\" HeaderImage=\"recurrence.gif\" ClassInfo=\"Icon\" Title=\"Recurrence\" Sealed=\"TRUE\" NoEditFormBreak=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"fRecurrence\" ColName=\"bit2\"><Default>FALSE</Default><FieldRefs><FieldRef Name=\"RecurrenceData\" RefType=\"RecurData\"/><FieldRef Name=\"EventType\" RefType=\"EventType\"/><FieldRef Name=\"UID\" RefType=\"UID\"/><FieldRef Name=\"RecurrenceID\" RefType=\"RecurrenceId\"/><FieldRef Name=\"EventCanceled\" RefType=\"EventCancel\"/><FieldRef Name=\"EventDate\" RefType=\"StartDate\"/><FieldRef Name=\"EndDate\" RefType=\"EndDate\"/><FieldRef Name=\"Duration\" RefType=\"Duration\"/><FieldRef Name=\"TimeZone\" RefType=\"TimeZone\"/><FieldRef Name=\"XMLTZone\" RefType=\"XMLTZone\"/><FieldRef Name=\"MasterSeriesItemID\" RefType=\"MasterSeriesItemID\"/><FieldRef Name=\"WorkspaceLink\" RefType=\"CPLink\"/><FieldRef Name=\"Workspace\" RefType=\"LinkURL\"/></FieldRefs></Field><Field ID=\"{08fc65f9-48eb-4e99-bd61-5946c439e691}\" Type=\"CrossProjectLink\" Name=\"WorkspaceLink\" Format=\"EventList\" DisplayName=\"Workspace\" DisplayImage=\"mtgicon.gif\" HeaderImage=\"mtgicnhd.gif\" ClassInfo=\"Icon\" Title=\"Meeting Workspace\" Filterable=\"TRUE\" Sealed=\"TRUE\" Hidden=\"TRUE\" ShowInViewForm=\"FALSE\" ShowInEditForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"WorkspaceLink\" ColName=\"bit3\"><FieldRefs><FieldRef Name=\"Workspace\" RefType=\"LinkURL\" CreateURL=\"newMWS.aspx\">Use a Meeting Workspace to organize attendees, agendas, documents, minutes, and other details for this event.</FieldRef><FieldRef Name=\"RecurrenceID\" RefType=\"RecurrenceId\" DisplayName=\"Instance ID\"/><FieldRef Name=\"EventType\" RefType=\"EventType\"/><FieldRef Name=\"UID\" RefType=\"UID\"/></FieldRefs></Field><Field ID=\"{5d1d4e76-091a-4e03-ae83-6a59847731c0}\" Type=\"Integer\" Name=\"EventType\" DisplayName=\"Event Type\" Sealed=\"TRUE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"EventType\" ColName=\"int1\"/><Field ID=\"{63055d04-01b5-48f3-9e1e-e564e7c6b23b}\" Type=\"Guid\" Name=\"UID\" DisplayName=\"UID\" Sealed=\"TRUE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"UID\" ColName=\"uniqueidentifier1\"/><Field ID=\"{dfcc8fff-7c4c-45d6-94ed-14ce0719efef}\" Type=\"DateTime\" Name=\"RecurrenceID\" DisplayName=\"Recurrence ID\" CalType=\"1\" Format=\"ISO8601Gregorian\" Sealed=\"TRUE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"RecurrenceID\" ColName=\"datetime3\"/><Field ID=\"{b8bbe503-bb22-4237-8d9e-0587756a2176}\" Type=\"Boolean\" Name=\"EventCanceled\" DisplayName=\"Event Cancelled\" Sealed=\"TRUE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"EventCanceled\" ColName=\"bit4\"/><Field ID=\"{4d54445d-1c84-4a6d-b8db-a51ded4e1acc}\" Type=\"Integer\" Name=\"Duration\" DisplayName=\"Duration\" Hidden=\"TRUE\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Duration\" ColName=\"int2\"/><Field ID=\"{d12572d0-0a1e-4438-89b5-4d0430be7603}\" Type=\"Note\" Name=\"RecurrenceData\" DisplayName=\"RecurrenceData\" Hidden=\"TRUE\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"RecurrenceData\" ColName=\"ntext3\"/><Field ID=\"{6cc1c612-748a-48d8-88f2-944f477f301b}\" Type=\"Integer\" Name=\"TimeZone\" DisplayName=\"TimeZone\" Sealed=\"TRUE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"TimeZone\" ColName=\"int3\"/><Field ID=\"{c4b72ed6-45aa-4422-bff1-2b6750d30819}\" Type=\"Note\" Name=\"XMLTZone\" DisplayName=\"XMLTZone\" Hidden=\"TRUE\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"XMLTZone\" ColName=\"ntext4\"/><Field ID=\"{9b2bed84-7769-40e3-9b1d-7954a4053834}\" Type=\"Integer\" Name=\"MasterSeriesItemID\" DisplayName=\"MasterSeriesItemID\" Sealed=\"TRUE\" Hidden=\"TRUE\" Indexed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"MasterSeriesItemID\" ColName=\"int4\"/><Field ID=\"{881eac4a-55a5-48b6-a28e-8329d7486120}\" Type=\"URL\" Name=\"Workspace\" DisplayName=\"WorkspaceUrl\" Hidden=\"TRUE\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Workspace\" ColName=\"nvarchar5\" ColName2=\"nvarchar6\"/></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments><Folder TargetName=\"Event\"/></ContentType>","Scope":"/sites/portal/Lists/Events","Sealed":false,"StringId":"0x010200973548ACFFDA0948BE80AF607C4E28F9"});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '0x010200973548ACFFDA0948BE80AF607C4E28F9', listTitle: 'Events' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({"Description":"Create a new meeting, deadline or other event.","DisplayFormTemplateName":"ListForm","DisplayFormUrl":"","DocumentTemplate":"","DocumentTemplateUrl":"","EditFormTemplateName":"ListForm","EditFormUrl":"","Group":"List Content Types","Hidden":false,"Id":{"StringValue":"0x010200973548ACFFDA0948BE80AF607C4E28F9"},"JSLink":"","MobileDisplayFormUrl":"","MobileEditFormUrl":"","MobileNewFormUrl":"","Name":"Event","NewFormTemplateName":"ListForm","NewFormUrl":"","ReadOnly":false,"SchemaXml":"<ContentType ID=\"0x010200973548ACFFDA0948BE80AF607C4E28F9\" Name=\"Event\" Group=\"List Content Types\" V2ListTemplateName=\"events\" Description=\"Create a new meeting, deadline or other event.\" Version=\"0\" FeatureId=\"{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}\"><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Type=\"Computed\" DisplayName=\"Content Type\" Name=\"ContentType\" DisplaceOnUpgrade=\"TRUE\" RenderXMLUsingPattern=\"TRUE\" Sortable=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ContentTypeId\"/></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\"/></MapToContentType></DisplayPattern></Field><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Type=\"Text\" Name=\"Title\" DisplayName=\"Title\" Required=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" FromBaseType=\"TRUE\" ColName=\"nvarchar1\" ShowInNewForm=\"TRUE\" ShowInEditForm=\"TRUE\"/><Field ID=\"{288f5f32-8462-4175-8f09-dd7ba29359a9}\" Type=\"Text\" Name=\"Location\" DisplayName=\"Location\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Location\" ColName=\"nvarchar4\"/><Field Type=\"DateTime\" ID=\"{64cd368d-2f95-4bfc-a1f9-8d4324ecb007}\" Name=\"EventDate\" DisplayName=\"Start Time\" Format=\"DateTime\" Sealed=\"TRUE\" Required=\"TRUE\" FromBaseType=\"TRUE\" Filterable=\"FALSE\" FilterableNoRecurrence=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"EventDate\" ColName=\"datetime1\"><Default>[today]</Default><FieldRefs><FieldRef Name=\"fAllDayEvent\" RefType=\"AllDayEvent\"/></FieldRefs></Field><Field ID=\"{2684f9f2-54be-429f-ba06-76754fc056bf}\" Type=\"DateTime\" Name=\"EndDate\" DisplayName=\"End Time\" Format=\"DateTime\" Sealed=\"TRUE\" Required=\"TRUE\" Filterable=\"FALSE\" FilterableNoRecurrence=\"TRUE\" Indexed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"EndDate\" ColName=\"datetime2\"><Default>[today]</Default><FieldRefs><FieldRef Name=\"fAllDayEvent\" RefType=\"AllDayEvent\"/></FieldRefs></Field><Field Type=\"Note\" ID=\"{9da97a8a-1da5-4a77-98d3-4bc10456e700}\" Name=\"Description\" RichText=\"TRUE\" RichTextMode=\"FullHtml\" DisplayName=\"Description\" Sortable=\"FALSE\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Description\" ColName=\"ntext2\"/><Field ID=\"{6df9bd52-550e-4a30-bc31-a4366832a87d}\" Name=\"Category\" DisplayName=\"Category\" Type=\"Choice\" Format=\"Dropdown\" FillInChoice=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Category\" ColName=\"nvarchar7\"><CHOICES><CHOICE>Meeting</CHOICE><CHOICE>Work hours</CHOICE><CHOICE>Business</CHOICE><CHOICE>Holiday</CHOICE><CHOICE>Get-together</CHOICE><CHOICE>Gifts</CHOICE><CHOICE>Birthday</CHOICE><CHOICE>Anniversary</CHOICE></CHOICES></Field><Field ID=\"{7d95d1f4-f5fd-4a70-90cd-b35abc9b5bc8}\" Type=\"AllDayEvent\" Name=\"fAllDayEvent\" DisplaceOnUpgrade=\"TRUE\" DisplayName=\"All Day Event\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"fAllDayEvent\" ColName=\"bit1\"><FieldRefs><FieldRef Name=\"EventDate\" RefType=\"StartDate\"/><FieldRef Name=\"EndDate\" RefType=\"EndDate\"/><FieldRef Name=\"TimeZone\" RefType=\"TimeZone\"/><FieldRef Name=\"XMLTZone\" RefType=\"XMLTZone\"/></FieldRefs></Field><Field ID=\"{f2e63656-135e-4f1c-8fc2-ccbe74071901}\" Type=\"Recurrence\" Name=\"fRecurrence\" DisplayName=\"Recurrence\" DisplayImage=\"recur.gif\" ExceptionImage=\"recurEx.gif\" HeaderImage=\"recurrence.gif\" ClassInfo=\"Icon\" Title=\"Recurrence\" Sealed=\"TRUE\" NoEditFormBreak=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"fRecurrence\" ColName=\"bit2\"><Default>FALSE</Default><FieldRefs><FieldRef Name=\"RecurrenceData\" RefType=\"RecurData\"/><FieldRef Name=\"EventType\" RefType=\"EventType\"/><FieldRef Name=\"UID\" RefType=\"UID\"/><FieldRef Name=\"RecurrenceID\" RefType=\"RecurrenceId\"/><FieldRef Name=\"EventCanceled\" RefType=\"EventCancel\"/><FieldRef Name=\"EventDate\" RefType=\"StartDate\"/><FieldRef Name=\"EndDate\" RefType=\"EndDate\"/><FieldRef Name=\"Duration\" RefType=\"Duration\"/><FieldRef Name=\"TimeZone\" RefType=\"TimeZone\"/><FieldRef Name=\"XMLTZone\" RefType=\"XMLTZone\"/><FieldRef Name=\"MasterSeriesItemID\" RefType=\"MasterSeriesItemID\"/><FieldRef Name=\"WorkspaceLink\" RefType=\"CPLink\"/><FieldRef Name=\"Workspace\" RefType=\"LinkURL\"/></FieldRefs></Field><Field ID=\"{08fc65f9-48eb-4e99-bd61-5946c439e691}\" Type=\"CrossProjectLink\" Name=\"WorkspaceLink\" Format=\"EventList\" DisplayName=\"Workspace\" DisplayImage=\"mtgicon.gif\" HeaderImage=\"mtgicnhd.gif\" ClassInfo=\"Icon\" Title=\"Meeting Workspace\" Filterable=\"TRUE\" Sealed=\"TRUE\" Hidden=\"TRUE\" ShowInViewForm=\"FALSE\" ShowInEditForm=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"WorkspaceLink\" ColName=\"bit3\"><FieldRefs><FieldRef Name=\"Workspace\" RefType=\"LinkURL\" CreateURL=\"newMWS.aspx\">Use a Meeting Workspace to organize attendees, agendas, documents, minutes, and other details for this event.</FieldRef><FieldRef Name=\"RecurrenceID\" RefType=\"RecurrenceId\" DisplayName=\"Instance ID\"/><FieldRef Name=\"EventType\" RefType=\"EventType\"/><FieldRef Name=\"UID\" RefType=\"UID\"/></FieldRefs></Field><Field ID=\"{5d1d4e76-091a-4e03-ae83-6a59847731c0}\" Type=\"Integer\" Name=\"EventType\" DisplayName=\"Event Type\" Sealed=\"TRUE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"EventType\" ColName=\"int1\"/><Field ID=\"{63055d04-01b5-48f3-9e1e-e564e7c6b23b}\" Type=\"Guid\" Name=\"UID\" DisplayName=\"UID\" Sealed=\"TRUE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"UID\" ColName=\"uniqueidentifier1\"/><Field ID=\"{dfcc8fff-7c4c-45d6-94ed-14ce0719efef}\" Type=\"DateTime\" Name=\"RecurrenceID\" DisplayName=\"Recurrence ID\" CalType=\"1\" Format=\"ISO8601Gregorian\" Sealed=\"TRUE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"RecurrenceID\" ColName=\"datetime3\"/><Field ID=\"{b8bbe503-bb22-4237-8d9e-0587756a2176}\" Type=\"Boolean\" Name=\"EventCanceled\" DisplayName=\"Event Cancelled\" Sealed=\"TRUE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"EventCanceled\" ColName=\"bit4\"/><Field ID=\"{4d54445d-1c84-4a6d-b8db-a51ded4e1acc}\" Type=\"Integer\" Name=\"Duration\" DisplayName=\"Duration\" Hidden=\"TRUE\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Duration\" ColName=\"int2\"/><Field ID=\"{d12572d0-0a1e-4438-89b5-4d0430be7603}\" Type=\"Note\" Name=\"RecurrenceData\" DisplayName=\"RecurrenceData\" Hidden=\"TRUE\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"RecurrenceData\" ColName=\"ntext3\"/><Field ID=\"{6cc1c612-748a-48d8-88f2-944f477f301b}\" Type=\"Integer\" Name=\"TimeZone\" DisplayName=\"TimeZone\" Sealed=\"TRUE\" Hidden=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"TimeZone\" ColName=\"int3\"/><Field ID=\"{c4b72ed6-45aa-4422-bff1-2b6750d30819}\" Type=\"Note\" Name=\"XMLTZone\" DisplayName=\"XMLTZone\" Hidden=\"TRUE\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"XMLTZone\" ColName=\"ntext4\"/><Field ID=\"{9b2bed84-7769-40e3-9b1d-7954a4053834}\" Type=\"Integer\" Name=\"MasterSeriesItemID\" DisplayName=\"MasterSeriesItemID\" Sealed=\"TRUE\" Hidden=\"TRUE\" Indexed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"MasterSeriesItemID\" ColName=\"int4\"/><Field ID=\"{881eac4a-55a5-48b6-a28e-8329d7486120}\" Type=\"URL\" Name=\"Workspace\" DisplayName=\"WorkspaceUrl\" Hidden=\"TRUE\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Workspace\" ColName=\"nvarchar5\" ColName2=\"nvarchar6\"/></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>ListForm</Display><Edit>ListForm</Edit><New>ListForm</New></FormTemplates></XmlDocument></XmlDocuments><Folder TargetName=\"Event\"/></ContentType>","Scope":"/sites/portal/Lists/Events","Sealed":false,"StringId":"0x010200973548ACFFDA0948BE80AF607C4E28F9"}));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly escapes special characters in the content type id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0%3D0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '0=0100558D85B7216F6A489A499DB361E1AE2F' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Content type with ID 0=0100558D85B7216F6A489A499DB361E1AE2F not found`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles site content type not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '0x0100558D85B7216F6A489A499DB361E1AE2F' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Content type with ID 0x0100558D85B7216F6A489A499DB361E1AE2F not found`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles list content type not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.resolve({
          "odata.null": true
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '0x0100558D85B7216F6A489A499DB361E1AE2F', listTitle: 'Documents' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Content type with ID 0x0100558D85B7216F6A489A499DB361E1AE2F not found`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles list not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/contenttypes('0x0100558D85B7216F6A489A499DB361E1AE2F')`) > -1) {
        return Promise.reject({
          error: {
            "odata.error": {
              "code": "-1, System.ArgumentException",
              "message": {
                "lang": "en-US",
                "value": "List 'Documents' does not exist at site with URL 'https://contoso.sharepoint.com/sites/portal'."
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', id: '0x0100558D85B7216F6A489A499DB361E1AE2F', listTitle: 'Documents' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("List 'Documents' does not exist at site with URL 'https://contoso.sharepoint.com/sites/portal'.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notStrictEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('configures id as string option', () => {
    const types = (command.types() as CommandTypes);
    ['i', 'id'].forEach(o => {
      assert.notStrictEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
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

  it('fails validation if the specified site URL is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'site.com', id: '0x0100558D85B7216F6A489A499DB361E1AE2F' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', id: '0x0100558D85B7216F6A489A499DB361E1AE2F' } });
    assert.strictEqual(actual, true);
  });
});