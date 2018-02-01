import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { ListInstance } from "./ListInstance";
import Auth from '../../../../Auth';
import { ListTemplateType } from './ListTemplateType';
import { DraftVisibilityType } from './DraftVisibilityType';
import { ListExperience } from './ListExperience';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title: string;
  baseTemplate: string;
  webUrl: string;
  description?: string;
  templateFeatureId?: string;
  schemaXml?: string;
  allowDeletion?: string;
  allowEveryoneViewItems?: string;
  allowMultiResponses?: string;
  contentTypesEnabled?: string;
  crawlNonDefaultViews?: string;
  defaultContentApprovalWorkflowId?: string;
  defaultDisplayFormUrl?: string;
  defaultEditFormUrl?: string;
  direction?: string;
  disableGridEditing?: string;
  draftVersionVisibility?: string;
  emailAlias?: string;
  enableAssignToEmail?: string;
  enableAttachments?: string;
  enableDeployWithDependentList?: string;
  enableFolderCreation?: string;
  enableMinorVersions?: string;
  enableModeration?: string;
  enablePeopleSelector?: string;
  enableResourceSelector?: string;
  enableSchemaCaching?: string;
  enableSyndication?: string;
  enableThrottling?: string;
  enableVersioning?: string;
  enforceDataValidation?: string;
  excludeFromOfflineClient?: string;
  fetchPropertyBagForListView?: string;
  followable?: string;
  forceCheckout?: string;
  forceDefaultContentType?: string;
  hidden?: string;
  includedInMyFilesScope?: string;
  irmEnabled?: string;
  irmExpire?: string;
  irmReject?: string;
  isApplicationList?: string;
  listExperienceOptions?: string;
  majorVersionLimit?: number;
  majorWithMinorVersionsLimit?: number;
  multipleDataList?: string;
  navigateForFormsPages?: string;
  needUpdateSiteClientTag?: string;
  noCrawl?: string;
  onQuickLaunch?: string;
  ordered?: string;
  parserDisabled?: string;
  readOnlyUI?: string;
  readSecurity?: number;
  requestAccessEnabled?: string;
  restrictUserUpdates?: string;
  sendToLocationName?: string;
  sendToLocationUrl?: string;
  showUser?: string;
  useFormsForDisplay?: string;
  validationFormula?: string;
  validationMessage?: string;
  writeSecurity?: number;
}

class ListAddCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_ADD;
  }

  public get description(): string {
    return 'Gets information about the specific list';
  }

  /**
   * Maps the base ListTemplateType enum to string array so it can 
   * more easily be used in validation or descriptions.
   */
  protected get listTemplateTypeMap(): string[] {
    const result: string[] = [];

    for (let template in ListTemplateType) {
      if (typeof ListTemplateType[template] === 'number') {
        result.push(template);
      }
    }
    return result;
  }

  /**
   * Maps the base DraftVisibilityType enum to string array so it can 
   * more easily be used in validation or descriptions.
   */
  protected get draftVisibilityTypeMap(): string[] {
    const result: string[] = [];

    for (let draftType in DraftVisibilityType) {
      if (typeof DraftVisibilityType[draftType] === 'number') {
        result.push(draftType);
      }
    }
    return result;
  }

  /**
   * Maps the base ListExperience enum to string array so it can 
   * more easily be used in validation or descriptions.
   */
  protected get listExperienceMap(): string[] {
    const result: string[] = [];

    for (let experience in ListExperience) {
      if (typeof ListExperience[experience] === 'number') {
        result.push(experience);
      }
    }
    return result;
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.description = typeof args.options.description === 'string';
    telemetryProps.templateFeatureId = typeof args.options.templateFeatureId === 'string';
    telemetryProps.schemaXml = typeof args.options.schemaXml === 'string';
    telemetryProps.allowDeletion = typeof args.options.allowDeletion === 'string';
    telemetryProps.allowEveryoneViewItems = typeof args.options.allowEveryoneViewItems === 'string';
    telemetryProps.allowMultiResponses = typeof args.options.allowMultiResponses === 'string';
    telemetryProps.contentTypesEnabled = typeof args.options.contentTypesEnabled === 'string';
    telemetryProps.crawlNonDefaultViews = typeof args.options.crawlNonDefaultViews === 'string';
    telemetryProps.defaultContentApprovalWorkflowId = typeof args.options.defaultContentApprovalWorkflowId === 'string';
    telemetryProps.defaultDisplayFormUrl = typeof args.options.defaultDisplayFormUrl === 'string';
    telemetryProps.defaultEditFormUrl = typeof args.options.defaultEditFormUrl === 'string';
    telemetryProps.direction = typeof args.options.direction === 'string';
    telemetryProps.disableGridEditing = typeof args.options.disableGridEditing === 'string';
    telemetryProps.draftVersionVisibility = typeof args.options.draftVersionVisibility === 'string';
    telemetryProps.emailAlias = typeof args.options.emailAlias === 'string';
    telemetryProps.enableAssignToEmail = typeof args.options.enableAssignToEmail === 'string';
    telemetryProps.enableAttachments = typeof args.options.enableAttachments === 'string';
    telemetryProps.enableDeployWithDependentList = typeof args.options.enableDeployWithDependentList === 'string';
    telemetryProps.enableFolderCreation = typeof args.options.enableFolderCreation === 'string';
    telemetryProps.enableMinorVersions = typeof args.options.enableMinorVersions === 'string';
    telemetryProps.enableModeration = typeof args.options.enableModeration === 'string';
    telemetryProps.enablePeopleSelector = typeof args.options.enablePeopleSelector === 'string';
    telemetryProps.enableResourceSelector = typeof args.options.enableResourceSelector === 'string';
    telemetryProps.enableSchemaCaching = typeof args.options.enableSchemaCaching === 'string';
    telemetryProps.enableSyndication = typeof args.options.enableSyndication === 'string';
    telemetryProps.enableThrottling = typeof args.options.enableThrottling === 'string';
    telemetryProps.enableVersioning = typeof args.options.enableVersioning === 'string';
    telemetryProps.enforceDataValidation = typeof args.options.enforceDataValidation === 'string';
    telemetryProps.excludeFromOfflineClient = typeof args.options.excludeFromOfflineClient === 'string';
    telemetryProps.fetchPropertyBagForListView = typeof args.options.fetchPropertyBagForListView === 'string';
    telemetryProps.followable = typeof args.options.followable === 'string';
    telemetryProps.forceCheckout = typeof args.options.forceCheckout === 'string';
    telemetryProps.forceDefaultContentType = typeof args.options.forceDefaultContentType === 'string';
    telemetryProps.hidden = typeof args.options.hidden === 'string';
    telemetryProps.includedInMyFilesScope = typeof args.options.includedInMyFilesScope === 'string';
    telemetryProps.irmEnabled = typeof args.options.irmEnabled === 'string';
    telemetryProps.irmExpire = typeof args.options.irmExpire === 'string';
    telemetryProps.irmReject = typeof args.options.irmReject === 'string';
    telemetryProps.isApplicationList = typeof args.options.isApplicationList === 'string';
    telemetryProps.listExperienceOptions = typeof args.options.listExperienceOptions === 'string';
    telemetryProps.majorVersionLimit = typeof args.options.majorVersionLimit === 'string';
    telemetryProps.majorWithMinorVersionsLimit = typeof args.options.majorWithMinorVersionsLimit === 'string';
    telemetryProps.multipleDataList = typeof args.options.multipleDataList === 'string';
    telemetryProps.navigateForFormsPages = typeof args.options.navigateForFormsPages === 'string';
    telemetryProps.needUpdateSiteClientTag = typeof args.options.needUpdateSiteClientTag === 'string';
    telemetryProps.noCrawl = typeof args.options.noCrawl === 'string';
    telemetryProps.onQuickLaunch = typeof args.options.onQuickLaunch === 'string';
    telemetryProps.ordered = typeof args.options.ordered === 'string';
    telemetryProps.parserDisabled = typeof args.options.parserDisabled === 'string';
    telemetryProps.readOnlyUI = typeof args.options.readOnlyUI === 'string';
    telemetryProps.readSecurity = typeof args.options.readSecurity === 'string';
    telemetryProps.requestAccessEnabled = typeof args.options.requestAccessEnabled === 'string';
    telemetryProps.restrictUserUpdates = typeof args.options.readOnlyUI === 'string';
    telemetryProps.sendToLocationName = typeof args.options.sendToLocationName === 'string';
    telemetryProps.sendToLocationUrl = typeof args.options.sendToLocationUrl === 'string';
    telemetryProps.showUser = typeof args.options.showUser === 'string';
    telemetryProps.useFormsForDisplay = typeof args.options.useFormsForDisplay === 'string';
    telemetryProps.validationFormula = typeof args.options.validationFormula === 'string';
    telemetryProps.validationMessage = typeof args.options.validationMessage === 'string';
    telemetryProps.writeSecurity = typeof args.options.writeSecurity === 'string';

    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): Promise<ListInstance> => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving information for list in site at ${args.options.webUrl}...`);
        }

        const requestBody: any = this.mapRequestBody(args.options);

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/lists`,
          method: 'POST',
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          body: requestBody, 
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((listInstance: ListInstance): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(listInstance);
          cmd.log('');
        }

        cmd.log(listInstance);

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-t, --title <title>',
        description: 'The displayed title for the list'
      },
      {
        option: '--baseTemplate <baseTemplate>',
        description: `The list definition type on which the list is based. Allowed values ${this.listTemplateTypeMap.join('|')}. Default ${this.listTemplateTypeMap[0]}`,
        autocomplete: this.listTemplateTypeMap
      },
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list should be added'
      },
      {
        option: '--description [description]',
        description: 'The description for the list'
      },
      {
        option: '--templateFeatureId [templateFeatureId]',
        description: 'The globally unique identifier (GUID) of a template feature that is associated with the list'
      },
      {
        option: '--schemaXml [schemaXml]',
        description: 'The schema in Collaborative Application Markup Language (CAML) schemas that defines the list'
      },
      {
        option: '--allowDeletion [allowDeletion]',
        description: 'Boolean value specifying whether the list can be deleted',
        autocomplete: ['true', 'false']
      },   
      {
        option: '--allowEveryoneViewItems [allowEveryoneViewItems]',
        description: 'Boolean value specifying whether everyone can view documents in the document library or attachments to items in the list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowMultiResponses [allowMultiResponses]',
        description: 'Boolean value specifying whether users are allowed to give multiple responses to the survey',
        autocomplete: ['true', 'false']
      },
      {
        option: '--contentTypesEnabled [contentTypesEnabled]',
        description: 'Boolean value specifying whether content types are enabled for the list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--crawlNonDefaultViews [crawlNonDefaultViews]',
        description: 'Boolean value specifying whether to crawl non default views',
        autocomplete: ['true', 'false']
      },
      {
        option: '--defaultContentApprovalWorkflowId [defaultContentApprovalWorkflowId]',
        description: 'Value that specifies the default workflow identifier for content approval on the list (GUID)'
      },
      {
        option: '--defaultDisplayFormUrl [defaultDisplayFormUrl]',
        description: 'Value that specifies the location of the default display form for the list'
      },
      {
        option: '--defaultEditFormUrl [defaultEditFormUrl]',
        description: 'Value that specifies the URL of the edit form to use for list items in the list'
      },
      {
        option: '--direction [direction]',
        description: 'Value that specifies the reading order of the list. Valid values are NONE|LTR|RTL',
        autocomplete: ['NONE', 'LTR', 'RTL']
      },
      {
        option: '--disableGridEditing [disableGridEditing]',
        description: 'Property for assigning or retrieving grid editing on the list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--draftVersionVisibility [draftVersionVisibility]',
        description: `Value that specifies the minimum permission required to view minor versions and drafts within the list. Allowed values ${this.draftVisibilityTypeMap.join('|')}. Default ${this.draftVisibilityTypeMap[0]}`,
        autocomplete: this.draftVisibilityTypeMap
      },
      {
        option: '--emailAlias [emailAlias]',
        description: 'If e-mail notification is enabled, gets or sets the e-mail address to use to notify to the owner of an item when an assignment has changed or the item has been updated.'
      },
      {
        option: '--enableAssignToEmail [enableAssignToEmail]',
        description: 'Boolean value specifying whether e-mail notification is enabled for the list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableAttachments [enableAttachments]',
        description: 'Boolean value that specifies whether attachments can be added to items in the list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableDeployWithDependentList [enableDeployWithDependentList]',
        description: ' Boolean value that specifies whether the list can be deployed with a dependent list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableFolderCreation [enableFolderCreation]',
        description: 'Boolean value that specifies whether folders can be created for the list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableMinorVersions [enableMinorVersions]',
        description: 'Boolean value that specifies whether minor versions are enabled when versioning is enabled for the document library',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableModeration [enableModeration]',
        description: 'Boolean value that specifies whether Content Approval is enabled for the list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enablePeopleSelector [enablePeopleSelector]',
        description: 'Enable user selector on event list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableResourceSelector [enableResourceSelector]',
        description: 'Enables resource selector on an event list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableSchemaCaching [enableSchemaCaching]',
        description: 'Boolean value specifying whether schema caching is enabled for the list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableSyndication [enableSyndication]',
        description: 'Boolean value that specifies whether RSS syndication is enabled for the list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableThrottling [enableThrottling]',
        description: 'Indicates whether throttling for this list is enabled or not',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableVersioning [enableVersioning]',
        description: 'Boolean value that specifies whether versioning is enabled for the document library.',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enforceDataValidation [enforceDataValidation]',
        description: 'Value that indicates whether certain field properties are enforced when an item is added or updated',
        autocomplete: ['true', 'false']
      },
      {
        option: '--excludeFromOfflineClient [excludeFromOfflineClient]',
        description: 'Value that indicates whether the list should be downloaded to the client during offline synchronization',
        autocomplete: ['true', 'false']
      },
      {
        option: '--fetchPropertyBagForListView [fetchPropertyBagForListView]',
        description: 'Specifies whether property bag information, as part of the list schema JSON, is retrieved when the list is being rendered on the client',
        autocomplete: ['true', 'false']
      },
      {
        option: '--followable [followable]',
        description: 'Can a list be followed in an activity feed?',
        autocomplete: ['true', 'false']
      },
      {
        option: '--forceCheckout [forceCheckout]',
        description: 'Boolean value that specifies whether forced checkout is enabled for the document library',
        autocomplete: ['true', 'false']
      },
      {
        option: '--forceDefaultContentType [forceDefaultContentType]',
        description: 'Specifies whether we want to return the default Document root content type',
        autocomplete: ['true', 'false']
      },
      {
        option: '--hidden [hidden]',
        description: 'Boolean value that specifies whether the list is hidden',
        autocomplete: ['true', 'false']
      },
      {
        option: '--includedInMyFilesScope [includedInMyFilesScope]',
        description: 'Specifies whether this list is accessible to an app principal that has been granted an OAuth scope that contains the string “myfiles” by a case-insensitive comparison when the current user is a site collection administrator of the personal site that contains the list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--irmEnabled [irmEnabled]',
        description: 'Gets or sets a Boolean value that specifies whether Information Rights Management (IRM) is enabled for the list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--irmExpire [irmExpire]',
        description: 'Gets or sets a Boolean value that specifies whether Information Rights Management (IRM) expiration is enabled for the list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--irmReject [irmReject]',
        description: 'Gets or sets a Boolean value that specifies whether Information Rights Management (IRM) rejection is enabled for the list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--isApplicationList [isApplicationList]',
        description: 'Indicates whether this list should be treated as a top level navigation object or not',
        autocomplete: ['true', 'false']
      },
      {
        option: '--listExperienceOptions [listExperienceOptions]',
        description: `Gets or sets the list experience for the list. Allowed values ${this.listExperienceMap.join('|')}. Default ${this.listExperienceMap[0]}`,
        autocomplete: this.listExperienceMap
      },
      {
        option: '--majorVersionLimit [majorVersionLimit]',
        description: 'Gets or sets the maximum number of major versions allowed for an item in a document library that uses version control with major versions only.'
      },
      {
        option: '--majorWithMinorVersionsLimit [majorWithMinorVersionsLimit]',
        description: 'Gets or sets the maximum number of major versions that are allowed for an item in a document library that uses version control with both major and minor versions.'
      },
      {
        option: '--multipleDataList [multipleDataList]',
        description: 'Gets or sets a Boolean value that specifies whether the list in a Meeting Workspace site contains data for multiple meeting instances within the site',
        autocomplete: ['true', 'false']
      },
      {
        option: '--navigateForFormsPages [navigateForFormsPages]',
        description: 'Indicates whether to navigate for forms pages or use a modal dialog',
        autocomplete: ['true', 'false']
      },
      {
        option: '--needUpdateSiteClientTag [needUpdateSiteClientTag]',
        description: 'A boolean value that determines whether to editing documents in this list should increment the ClientTag for the site. The tag is used to allow clients to cache JS/CSS/resources that are retrieved from the Content DB, including custom CSR templates.',
        autocomplete: ['true', 'false']
      },
      {
        option: '--noCrawl [noCrawl]',
        description: 'Gets or sets a Boolean value specifying whether crawling is enabled for the list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--onQuickLaunch [onQuickLaunch]',
        description: 'Gets or sets a Boolean value that specifies whether the list appears on the Quick Launch area of the home page',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ordered [ordered]',
        description: 'Gets or sets a Boolean value that specifies whether the option to allow users to reorder items in the list is available on the Edit View page for the list',
        autocomplete: ['true', 'false']
      },
      {
        option: '-parserDisabled [parserDisabled]',
        description: 'Gets or sets a Boolean value that specifies whether the parser should be disabled',
        autocomplete: ['true', 'false']
      },
      {
        option: '--readOnlyUI [readOnlyUI]',
        description: 'A boolean value that indicates whether the UI for this list should be presented in a read-only fashion. This will not affect security nor will it actually prevent changes to the list from occurring - it only affects the way the UI is displayed',
        autocomplete: ['true', 'false']
      },
      {
        option: '--readSecurity [readSecurity]',
        description: 'Gets or sets the Read security setting for the list. Valid values are 1 (All users have Read access to all items)|2 (Users have Read access only to items that they create)',
        autocomplete: ['1', '2']
      },
      {
        option: '--requestAccessEnabled [requestAccessEnabled]',
        description: 'Gets or sets a Boolean value that specifies whether the option to allow users to request access to the list is available',
        autocomplete: ['true', 'false']
      },
      {
        option: '--restrictUserUpdates [restrictUserUpdates]',
        description: 'A boolean value that indicates whether the this list is a restricted one or not The value can\'t be changed if there are existing items in the list',
        autocomplete: ['true', 'false']
      },
      {
        option: '--sendToLocationName [sendToLocationName]',
        description: 'Gets or sets a file name to use when copying an item in the list to another document library.'
      },
      {
        option: '--sendToLocationUrl [sendToLocationUrl]',
        description: 'Gets or sets a URL to use when copying an item in the list to another document library'
      },
      {
        option: '--showUser [showUser]',
        description: 'Gets or sets a Boolean value that specifies whether names of users are shown in the results of the survey',
        autocomplete: ['true', 'false']
      },
      {
        option: '--useFormsForDisplay [useFormsForDisplay]',
        description: 'Indicates whether forms should be considered for display context or not',
        autocomplete: ['true', 'false']
      },
      {
        option: '--validationFormula [validationFormula]',
        description: 'Gets or sets a formula that is evaluated each time that a list item is added or updated.'
      },
      {
        option: '--validationMessage [validationMessage]',
        description: 'Gets or sets the message that is displayed when validation fails for a list item.'
      },
      {
        option: '--writeSecurity [writeSecurity]',
        description: 'Gets or sets the Write security setting for the list. Valid values are 1 (All users can modify all items)|2 (Users can modify only items that they create)|4 (Users cannot modify any list item)',
        autocomplete: ['1', '2', '4']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.title) {
        return 'Required parameter title missing';
      }

      if (!args.options.baseTemplate) {
        return 'Required parameter baseTemplate is missing';
      }

      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      const baseTemplateErrorMessage = `BaseTemplate option ${args.options.baseTemplate} is not recognized as valid choice. Please note it is case sensitive`;
      if (typeof args.options.baseTemplate === 'string') {
        const template: ListTemplateType = ListTemplateType[(args.options.baseTemplate.trim() as keyof typeof ListTemplateType)];
        if (!template) {
          return baseTemplateErrorMessage;
        }
      }
      else {
        return baseTemplateErrorMessage;
      }

      if (args.options.allowDeletion) {
        if (!Utils.isValidBoolean(args.options.allowDeletion)) {
          return `${args.options.allowDeletion} is not a valid Boolean`;
        }
      }

      if (args.options.allowEveryoneViewItems) {
        if (!Utils.isValidBoolean(args.options.allowEveryoneViewItems)) {
          return `${args.options.allowEveryoneViewItems} is not a valid Boolean`;
        }
      }

      if (args.options.allowMultiResponses) {
        if (!Utils.isValidBoolean(args.options.allowMultiResponses)) {
          return `${args.options.allowMultiResponses} is not a valid Boolean`;
        }
      }

      if (args.options.contentTypesEnabled) {
        if (!Utils.isValidBoolean(args.options.contentTypesEnabled)) {
          return `${args.options.contentTypesEnabled} is not a valid Boolean`;
        }
      }

      if (args.options.crawlNonDefaultViews) {
        if (!Utils.isValidBoolean(args.options.crawlNonDefaultViews)) {
          return `${args.options.crawlNonDefaultViews} is not a valid Boolean`;
        }
      }

      if (args.options.disableGridEditing) {
        if (!Utils.isValidBoolean(args.options.disableGridEditing)) {
          return `${args.options.disableGridEditing} is not a valid Boolean`;
        }
      }

      if (args.options.enableAssignToEmail) {
        if (!Utils.isValidBoolean(args.options.enableAssignToEmail)) {
          return `${args.options.enableAssignToEmail} is not a valid Boolean`;
        }
      }

      if (args.options.enableAttachments) {
        if (!Utils.isValidBoolean(args.options.enableAttachments)) {
          return `${args.options.enableAttachments} is not a valid Boolean`;
        }
      }

      if (args.options.enableDeployWithDependentList) {
        if (!Utils.isValidBoolean(args.options.enableDeployWithDependentList)) {
          return `${args.options.enableDeployWithDependentList} is not a valid Boolean`;
        }
      }

      if (args.options.enableFolderCreation) {
        if (!Utils.isValidBoolean(args.options.enableFolderCreation)) {
          return `${args.options.enableFolderCreation} is not a valid Boolean`;
        }
      }

      if (args.options.enableMinorVersions) {
        if (!Utils.isValidBoolean(args.options.enableMinorVersions)) {
          return `${args.options.enableMinorVersions} is not a valid Boolean`;
        }
      }

      if (args.options.enableModeration) {
        if (!Utils.isValidBoolean(args.options.enableModeration)) {
          return `${args.options.enableModeration} is not a valid Boolean`;
        }
      }

      if (args.options.enablePeopleSelector) {
        if (!Utils.isValidBoolean(args.options.enablePeopleSelector)) {
          return `${args.options.enablePeopleSelector} is not a valid Boolean`;
        }
      }

      if (args.options.enableResourceSelector) {
        if (!Utils.isValidBoolean(args.options.enableResourceSelector)) {
          return `${args.options.enableResourceSelector} is not a valid Boolean`;
        }
      }

      if (args.options.enableSchemaCaching) {
        if (!Utils.isValidBoolean(args.options.enableSchemaCaching)) {
          return `${args.options.enableSchemaCaching} is not a valid Boolean`;
        }
      }

      if (args.options.enableSyndication) {
        if (!Utils.isValidBoolean(args.options.enableSyndication)) {
          return `${args.options.enableSyndication} is not a valid Boolean`;
        }
      }

      if (args.options.enableThrottling) {
        if (!Utils.isValidBoolean(args.options.enableThrottling)) {
          return `${args.options.enableThrottling} is not a valid Boolean`;
        }
      }

      if (args.options.enableVersioning) {
        if (!Utils.isValidBoolean(args.options.enableVersioning)) {
          return `${args.options.enableVersioning} is not a valid Boolean`;
        }
      }

      if (args.options.enforceDataValidation) {
        if (!Utils.isValidBoolean(args.options.enforceDataValidation)) {
          return `${args.options.enforceDataValidation} is not a valid Boolean`;
        }
      }

      if (args.options.excludeFromOfflineClient) {
        if (!Utils.isValidBoolean(args.options.excludeFromOfflineClient)) {
          return `${args.options.excludeFromOfflineClient} is not a valid Boolean`;
        }
      }

      if (args.options.fetchPropertyBagForListView) {
        if (!Utils.isValidBoolean(args.options.fetchPropertyBagForListView)) {
          return `${args.options.fetchPropertyBagForListView} is not a valid Boolean`;
        }
      }

      if (args.options.followable) {
        if (!Utils.isValidBoolean(args.options.followable)) {
          return `${args.options.followable} is not a valid Boolean`;
        }
      }

      if (args.options.forceCheckout) {
        if (!Utils.isValidBoolean(args.options.forceCheckout)) {
          return `${args.options.forceCheckout} is not a valid Boolean`;
        }
      }

      if (args.options.forceDefaultContentType) {
        if (!Utils.isValidBoolean(args.options.forceDefaultContentType)) {
          return `${args.options.forceDefaultContentType} is not a valid Boolean`;
        }
      }

      if (args.options.hidden) {
        if (!Utils.isValidBoolean(args.options.hidden)) {
          return `${args.options.hidden} is not a valid Boolean`;
        }
      }

      if (args.options.includedInMyFilesScope) {
        if (!Utils.isValidBoolean(args.options.includedInMyFilesScope)) {
          return `${args.options.includedInMyFilesScope} is not a valid Boolean`;
        }
      }

      if (args.options.irmEnabled) {
        if (!Utils.isValidBoolean(args.options.irmEnabled)) {
          return `${args.options.irmEnabled} is not a valid Boolean`;
        }
      }

      if (args.options.irmExpire) {
        if (!Utils.isValidBoolean(args.options.irmExpire)) {
          return `${args.options.irmExpire} is not a valid Boolean`;
        }
      }

      if (args.options.irmReject) {
        if (!Utils.isValidBoolean(args.options.irmReject)) {
          return `${args.options.irmReject} is not a valid Boolean`;
        }
      }

      if (args.options.isApplicationList) {
        if (!Utils.isValidBoolean(args.options.isApplicationList)) {
          return `${args.options.isApplicationList} is not a valid Boolean`;
        }
      }

      if (args.options.multipleDataList) {
        if (!Utils.isValidBoolean(args.options.multipleDataList)) {
          return `${args.options.multipleDataList} is not a valid Boolean`;
        }
      }

      if (args.options.navigateForFormsPages) {
        if (!Utils.isValidBoolean(args.options.navigateForFormsPages)) {
          return `${args.options.navigateForFormsPages} is not a valid Boolean`;
        }
      }

      if (args.options.needUpdateSiteClientTag) {
        if (!Utils.isValidBoolean(args.options.needUpdateSiteClientTag)) {
          return `${args.options.needUpdateSiteClientTag} is not a valid Boolean`;
        }
      }

      if (args.options.noCrawl) {
        if (!Utils.isValidBoolean(args.options.noCrawl)) {
          return `${args.options.noCrawl} is not a valid Boolean`;
        }
      }

      if (args.options.onQuickLaunch) {
        if (!Utils.isValidBoolean(args.options.onQuickLaunch)) {
          return `${args.options.onQuickLaunch} is not a valid Boolean`;
        }
      }

      if (args.options.ordered) {
        if (!Utils.isValidBoolean(args.options.ordered)) {
          return `${args.options.ordered} is not a valid Boolean`;
        }
      }

      if (args.options.parserDisabled) {
        if (!Utils.isValidBoolean(args.options.parserDisabled)) {
          return `${args.options.parserDisabled} is not a valid Boolean`;
        }
      }

      if (args.options.readOnlyUI) {
        if (!Utils.isValidBoolean(args.options.readOnlyUI)) {
          return `${args.options.readOnlyUI} is not a valid Boolean`;
        }
      }

      if (args.options.requestAccessEnabled) {
        if (!Utils.isValidBoolean(args.options.requestAccessEnabled)) {
          return `${args.options.requestAccessEnabled} is not a valid Boolean`;
        }
      }

      if (args.options.restrictUserUpdates) {
        if (!Utils.isValidBoolean(args.options.restrictUserUpdates)) {
          return `${args.options.restrictUserUpdates} is not a valid Boolean`;
        }
      }

      if (args.options.showUser) {
        if (!Utils.isValidBoolean(args.options.showUser)) {
          return `${args.options.showUser} is not a valid Boolean`;
        }
      }

      if (args.options.useFormsForDisplay) {
        if (!Utils.isValidBoolean(args.options.useFormsForDisplay)) {
          return `${args.options.useFormsForDisplay} is not a valid Boolean`;
        }
      }

      if (args.options.templateFeatureId) {
        if (!Utils.isValidGuid(args.options.templateFeatureId)) {
          return `${args.options.templateFeatureId} is not a valid GUID`;
        }
      }

      if (args.options.defaultContentApprovalWorkflowId) {
        if (!Utils.isValidGuid(args.options.defaultContentApprovalWorkflowId)) {
          return `defaultContentApprovalWorkflowId value ${args.options.defaultContentApprovalWorkflowId} is not a valid GUID`;
        }
      }

      if (args.options.direction) {
        if (args.options.direction.toLowerCase() !== 'none' &&
          args.options.direction.toLowerCase() !== 'ltr' &&
          args.options.direction.toLowerCase() !== 'rtl') {
          return `direction value ${args.options.direction} is not a valid value. Allowed values are NONE|LTR|RTL`;
        }
      }

      if (args.options.draftVersionVisibility) {
        const draftVersionVisibilityErrorMessage = `draftVisibilityType option '${args.options.draftVersionVisibility}' is not recognized as valid choice. Please note it is case sensitive`;
        if (typeof args.options.draftVersionVisibility === 'string') {
          const draftType: DraftVisibilityType = DraftVisibilityType[(args.options.draftVersionVisibility.trim() as keyof typeof DraftVisibilityType)];

          if (!draftType) {
            return draftVersionVisibilityErrorMessage;
          }
        }
        else {
          return draftVersionVisibilityErrorMessage;
        }
      }

      if (args.options.emailAlias && !args.options.enableAssignToEmail) {
        return `emailAlias could not be set if enableAssignToEmail is not set to true. Please set enableAssignToEmail.`;
      }

      if (args.options.listExperienceOptions) {
        const listExperienceOptionsErrorMessage = `listExperienceOptions option '${args.options.listExperienceOptions}' is not recognized as valid choice. Please note it is case sensitive`;
        if (typeof args.options.listExperienceOptions === 'string') {
          const experience: ListExperience = ListExperience[(args.options.listExperienceOptions.trim() as keyof typeof ListExperience)];

          if (!experience) {
            return listExperienceOptionsErrorMessage;
          }
        }
        else {
          return listExperienceOptionsErrorMessage;
        }
      }

      if (args.options.majorVersionLimit && !args.options.enableVersioning) {
          return `majorVersionLimit option is only valid in combination with enableVersioning.`;
      }

      if ((args.options.majorWithMinorVersionsLimit && !args.options.enableMinorVersions) && (args.options.majorWithMinorVersionsLimit && !args.options.enableModeration)) {
        return `majorWithMinorVersionsLimit option is only valid in combination with enableMinorVersions or enableModeration.`;
      }
      
      if (args.options.readSecurity) {
        if (args.options.readSecurity !== 1 &&
          args.options.readSecurity !== 2) {
          return `readSecurity value ${args.options.readSecurity} is not a valid value. Allowed values are 1|2`;
        }
      }

      if (args.options.writeSecurity) {
        if (args.options.writeSecurity !== 1 &&
          args.options.writeSecurity !== 2 &&
          args.options.writeSecurity !== 4) {
          return `writeSecurity value ${args.options.writeSecurity} is not a valid value. Allowed values are 1|2|4`;
        }
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
      using the ${chalk.blue(commands.CONNECT)} command.
  
  Remarks:
  
    To add a list, you have to first connect to SharePoint using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    Add a list with title ${chalk.grey('Announcements')}, baseTemplate ${chalk.grey('Announcements')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.LIST_ADD} --title 'DemoList' --baseTemplate Announcements --webUrl https://contoso.sharepoint.com/sites/project-x

    Add a list with title ${chalk.grey('Announcements')}, baseTemplate ${chalk.grey('Announcements')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')} 
    with schemaXml ${chalk.grey('<List DocTemplateUrl="" DefaultViewUrl="" MobileDefaultViewUrl="" ID="{92FF93AB-920E-4D33-AE42-58B5E245BEFF}" Title="Announcements" Description="" ImageUrl="/_layouts/15/images/itann.png?rev=44" Name="{92FF93AB-920E-4D33-AE42-58B5E245BEFF}" BaseType="0" FeatureId="{00BFEA71-D1CE-42DE-9C63-A44004CE0104}" ServerTemplate="104" Created="20161221 20:02:12" Modified="20180110 19:35:15" LastDeleted="20161221 20:02:12" Version="0" Direction="none" ThumbnailSize="0" WebImageWidth="0" WebImageHeight="0" Flags="536875008" ItemCount="1" AnonymousPermMask="0" RootFolder="/sites/project-x/Lists/Announcements"      ReadSecurity="1" WriteSecurity="1" Author="3" EventSinkAssembly="" EventSinkClass="" EventSinkData="" EmailAlias="" WebFullUrl="/sites/project-x" WebId="7694137e-7038-4831-a1bd-218b28fe5d34" SendToLocation="" ScopeId="92facaf9-8d7a-40eb-9e69-362c91513cbd" MajorVersionLimit="0" MajorWithMinorVersionsLimit="0" WorkFlowId="00000000-0000-0000-0000-000000000000" HasUniqueScopes="False" NoThrottleListOperations="False" HasRelatedLists="False" Followable="False" Acl="" Flags2="0" RootFolderId="d4d67cc1-ad6e-4293-b039-ea49263d195f" ComplianceTag="" ComplianceFlags="0" UserModified="20161221 20:03:00" ListSchemaVersion="3" AclVersion="" AllowDeletion="True" AllowMultiResponses="False" EnableAttachments="True" EnableModeration="False" EnableVersioning="False" HasExternalDataSource="False" Hidden="False" MultipleDataList="False" Ordered="False" ShowUser="True" EnablePeopleSelector="False" EnableResourceSelector="False" EnableMinorVersion="False" RequireCheckout="False" ThrottleListOperations="False" ExcludeFromOfflineClient="False" CanOpenFileAsync="True" EnableFolderCreation="False" IrmEnabled="False" IrmSyncable="False" IsApplicationList="False" PreserveEmptyValues="False" StrictTypeCoercion="False" EnforceDataValidation="False" MaxItemsPerThrottledOperation="5000"></List>')}
      ${chalk.grey(config.delimiter)} ${commands.LIST_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --title Announcements --baseTemplate Announcements --schemaXml '<List DocTemplateUrl="" DefaultViewUrl="" MobileDefaultViewUrl="" ID="{92FF93AB-920E-4D33-AE42-58B5E245BEFF}" Title="Announcements" Description="" ImageUrl="/_layouts/15/images/itann.png?rev=44" Name="{92FF93AB-920E-4D33-AE42-58B5E245BEFF}" BaseType="0" FeatureId="{00BFEA71-D1CE-42DE-9C63-A44004CE0104}" ServerTemplate="104" Created="20161221 20:02:12" Modified="20180110 19:35:15" LastDeleted="20161221 20:02:12" Version="0" Direction="none" ThumbnailSize="0" WebImageWidth="0" WebImageHeight="0" Flags="536875008" ItemCount="1" AnonymousPermMask="0" RootFolder="/sites/project-x/Lists/Announcements"      ReadSecurity="1" WriteSecurity="1" Author="3" EventSinkAssembly="" EventSinkClass="" EventSinkData="" EmailAlias="" WebFullUrl="/sites/project-x" WebId="7694137e-7038-4831-a1bd-218b28fe5d34" SendToLocation="" ScopeId="92facaf9-8d7a-40eb-9e69-362c91513cbd" MajorVersionLimit="0" MajorWithMinorVersionsLimit="0" WorkFlowId="00000000-0000-0000-0000-000000000000" HasUniqueScopes="False" NoThrottleListOperations="False" HasRelatedLists="False" Followable="False" Acl="" Flags2="0" RootFolderId="d4d67cc1-ad6e-4293-b039-ea49263d195f" ComplianceTag="" ComplianceFlags="0" UserModified="20161221 20:03:00" ListSchemaVersion="3" AclVersion="" AllowDeletion="True" AllowMultiResponses="False" EnableAttachments="True" EnableModeration="False" EnableVersioning="False" HasExternalDataSource="False" Hidden="False" MultipleDataList="False" Ordered="False" ShowUser="True" EnablePeopleSelector="False" EnableResourceSelector="False" EnableMinorVersion="False" RequireCheckout="False" ThrottleListOperations="False" ExcludeFromOfflineClient="False" CanOpenFileAsync="True" EnableFolderCreation="False" IrmEnabled="False" IrmSyncable="False" IsApplicationList="False" PreserveEmptyValues="False" StrictTypeCoercion="False" EnforceDataValidation="False" MaxItemsPerThrottledOperation="5000"></List>'
    
    Add a list with title ${chalk.grey('Announcements')}, baseTemplate ${chalk.grey('Announcements')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
    with additional options --contentTypesEnabled --enableVersioning --majorVersionLimit 50
    ${chalk.grey(config.delimiter)} ${commands.LIST_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --title Announcements --baseTemplate Announcements --contentTypesEnabled --enableVersioning --majorVersionLimit 50
      `);
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {
      Title: options.title,
      BaseTemplate: ListTemplateType[(options.baseTemplate.trim() as keyof typeof ListTemplateType)].valueOf()
    };

    if (options.description) {
      requestBody.Description = options.description;
    }

    if (options.templateFeatureId) {
      requestBody.TemplateFeatureId = options.templateFeatureId;
    }

    if (options.schemaXml) {
      requestBody.SchemaXml = options.schemaXml.replace('\\', '\\\\').replace('"', '\\"');
    }

    if (options.allowDeletion) {
      requestBody.AllowDeletion = options.allowDeletion;
    }

    if (options.allowEveryoneViewItems) {
      requestBody.AllowEveryoneViewItems = options.allowEveryoneViewItems;
    }

    if (options.allowMultiResponses) {
      requestBody.AllowMultiResponses = options.allowMultiResponses;
    }

    if (options.contentTypesEnabled) {
      requestBody.ContentTypesEnabled = options.contentTypesEnabled;
    }

    if (options.crawlNonDefaultViews) {
      requestBody.CrawlNonDefaultViews = options.crawlNonDefaultViews;
    }

    if (options.defaultContentApprovalWorkflowId) {
      requestBody.DefaultContentApprovalWorkflowId = options.defaultContentApprovalWorkflowId;
    }

    if (options.defaultDisplayFormUrl) {
      requestBody.DefaultDisplayFormUrl = options.defaultDisplayFormUrl;
    }

    if (options.defaultEditFormUrl) {
      requestBody.DefaultEditFormUrl = options.defaultEditFormUrl;
    }

    if (options.direction) {
      requestBody.Direction = options.direction;
    }

    if (options.disableGridEditing) {
      requestBody.DisableGridEditing = options.disableGridEditing;
    }

    if (options.draftVersionVisibility) {
      requestBody.DraftVersionVisibility = options.draftVersionVisibility;
    }

    if (options.emailAlias) {
      requestBody.EmailAlias = options.emailAlias;
    }

    if (options.enableAssignToEmail) {
      requestBody.EnableAssignToEmail = options.enableAssignToEmail;
    }

    if (options.enableAttachments) {
      requestBody.EnableAttachments = options.enableAttachments;
    }

    if (options.enableDeployWithDependentList) {
      requestBody.EnableDeployWithDependentList = options.enableDeployWithDependentList;
    }

    if (options.enableFolderCreation) {
      requestBody.EnableFolderCreation = options.enableFolderCreation;
    }

    if (options.enableMinorVersions) {
      requestBody.EnableMinorVersions = options.enableMinorVersions;
    }

    if (options.enableModeration) {
      requestBody.EnableModeration = options.enableModeration;
    }

    if (options.enablePeopleSelector) {
      requestBody.EnablePeopleSelector = options.enablePeopleSelector;
    }

    if (options.enableResourceSelector) {
      requestBody.EnableResourceSelector = options.enableResourceSelector;
    }

    if (options.enableSchemaCaching) {
      requestBody.EnableSchemaCaching = options.enableSchemaCaching;
    }

    if (options.enableSyndication) {
      requestBody.EnableSyndication = options.enableSyndication;
    }

    if (options.enableThrottling) {
      requestBody.EnableThrottling = options.enableThrottling;
    }

    if (options.enableVersioning) {
      requestBody.EnableVersioning = options.enableVersioning;
    }

    if (options.enforceDataValidation) {
      requestBody.EnforceDataValidation = options.enforceDataValidation;
    }

    if (options.excludeFromOfflineClient) {
      requestBody.ExcludeFromOfflineClient = options.excludeFromOfflineClient;
    }

    if (options.fetchPropertyBagForListView) {
      requestBody.FetchPropertyBagForListView = options.fetchPropertyBagForListView;
    }

    if (options.followable) {
      requestBody.Followable = options.followable;
    }

    if (options.forceCheckout) {
      requestBody.ForceCheckout = options.forceCheckout;
    }

    if (options.forceDefaultContentType) {
      requestBody.ForceDefaultContentType = options.forceDefaultContentType;
    }

    if (options.hidden) {
      requestBody.Hidden = options.hidden;
    }

    if (options.includedInMyFilesScope) {
      requestBody.IncludedInMyFilesScope = options.includedInMyFilesScope;
    }

    if (options.irmEnabled) {
      requestBody.IrmEnabled = options.irmEnabled;
    }

    if (options.irmExpire) {
      requestBody.IrmExpire = options.irmExpire;
    }

    if (options.irmReject) {
      requestBody.IrmReject = options.irmReject;
    }

    if (options.isApplicationList) {
      requestBody.IsApplicationList = options.isApplicationList;
    }

    if (options.listExperienceOptions) {
      requestBody.ListExperienceOptions = options.listExperienceOptions;
    }

    if (options.majorVersionLimit) {
      requestBody.MajorVersionLimit = options.majorVersionLimit;
    }

    if (options.majorWithMinorVersionsLimit) {
      requestBody.MajorWithMinorVersionsLimit = options.majorWithMinorVersionsLimit;
    }

    if (options.multipleDataList) {
      requestBody.MultipleDataList = options.multipleDataList;
    }

    if (options.navigateForFormsPages) {
      requestBody.NavigateForFormsPages = options.navigateForFormsPages;
    }

    if (options.needUpdateSiteClientTag) {
      requestBody.NeedUpdateSiteClientTag = options.needUpdateSiteClientTag;
    }

    if (options.noCrawl) {
      requestBody.NoCrawl = options.noCrawl;
    }

    if (options.onQuickLaunch) {
      requestBody.OnQuickLaunch = options.onQuickLaunch;
    }

    if (options.ordered) {
      requestBody.Ordered = options.ordered;
    }

    if (options.parserDisabled) {
      requestBody.ParserDisabled = options.parserDisabled;
    }

    if (options.readOnlyUI) {
      requestBody.ReadOnlyUI = options.readOnlyUI;
    }

    if (options.readSecurity) {
      requestBody.ReadSecurity = options.readSecurity;
    }

    if (options.requestAccessEnabled) {
      requestBody.RequestAccessEnabled = options.requestAccessEnabled;
    }

    if (options.restrictUserUpdates) {
      requestBody.RestrictUserUpdates = options.restrictUserUpdates;
    }

    if (options.sendToLocationName) {
      requestBody.SendToLocationName = options.sendToLocationName;
    }

    if (options.sendToLocationUrl) {
      requestBody.SendToLocationUrl = options.sendToLocationUrl;
    }

    if (options.showUser) {
      requestBody.ShowUser = options.showUser;
    }

    if (options.useFormsForDisplay) {
      requestBody.UseFormsForDisplay = options.useFormsForDisplay;
    }

    if (options.validationFormula) {
      requestBody.ValidationFormula = options.validationFormula;
    }

    if (options.validationMessage) {
      requestBody.ValidationMessage = options.validationMessage;
    }

    if (options.writeSecurity) {
      requestBody.WriteSecurity = options.writeSecurity;
    }

    return requestBody;
  }
}

module.exports = new ListAddCommand();