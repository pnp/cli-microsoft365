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
  allowDeletion?: boolean;
  allowEveryoneViewItems?: boolean;
  allowMultiResponses?: boolean;
  contentTypesEnabled?: boolean;
  crawlNonDefaultViews?: boolean;
  defaultContentApprovalWorkflowId?: string;
  defaultDisplayFormUrl?: string;
  defaultEditFormUrl?: string;
  direction?: string;
  disableGridEditing?: boolean;
  draftVersionVisibility?: string;
  emailAlias?: string;
  enableAssignToEmail?: boolean;
  enableAttachments?: boolean;
  enableDeployWithDependentList?: boolean;
  enableFolderCreation?: boolean;
  enableMinorVersions?: boolean;
  enableModeration?: boolean;
  enablePeopleSelector?: boolean;
  enableResourceSelector?: boolean;
  enableSchemaCaching?: boolean;
  enableSyndication?: boolean;
  enableThrottling?: boolean;
  enableVersioning?: boolean;
  enforceDataValidation?: boolean;
  excludeFromOfflineClient?: boolean;
  fetchPropertyBagForListView?: boolean;
  followable?: boolean;
  forceCheckout?: boolean;
  forceDefaultContentType?: boolean;
  hidden?: boolean;
  includedInMyFilesScope?: boolean;
  irmEnabled?: boolean;
  irmExpire?: boolean;
  irmReject?: boolean;
  isApplicationList?: boolean;
  listExperienceOptions?: string;
  majorVersionLimit?: number;
  majorWithMinorVersionsLimit?: number;
  multipleDataList?: boolean;
  navigateForFormsPages?: boolean;
  needUpdateSiteClientTag?: boolean;
  noCrawl?: boolean;
  onQuickLaunch?: boolean;
  ordered?: boolean;
  parserDisabled?: boolean;
  readOnlyUI?: boolean;
  readSecurity?: number;
  requestAccessEnabled?: boolean;
  restrictUserUpdates?: boolean;
  sendToLocationName?: string;
  sendToLocationUrl?: string;
  showUser?: boolean;
  useFormsForDisplay?: boolean;
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
    telemetryProps.description = (!(!args.options.description)).toString();
    telemetryProps.templateFeatureId = (!(!args.options.templateFeatureId)).toString();
    telemetryProps.schemaXml = (!(!args.options.schemaXml)).toString();
    telemetryProps.allowDeletion = args.options.allowDeletion || false;
    telemetryProps.allowEveryoneViewItems = args.options.allowEveryoneViewItems || false;
    telemetryProps.allowMultiResponses = args.options.allowMultiResponses || false;
    telemetryProps.contentTypesEnabled = args.options.contentTypesEnabled || false;
    telemetryProps.crawlNonDefaultViews = args.options.crawlNonDefaultViews || false;
    telemetryProps.defaultContentApprovalWorkflowId = (!(!args.options.defaultContentApprovalWorkflowId)).toString();
    telemetryProps.defaultDisplayFormUrl = (!(!args.options.defaultDisplayFormUrl)).toString();
    telemetryProps.defaultEditFormUrl = (!(!args.options.defaultEditFormUrl)).toString();
    telemetryProps.direction = (!(!args.options.direction)).toString();
    telemetryProps.disableGridEditing = args.options.disableGridEditing || false;
    telemetryProps.draftVersionVisibility = (!(!args.options.draftVersionVisibility)).toString();
    telemetryProps.emailAlias = (!(!args.options.emailAlias)).toString();
    telemetryProps.enableAssignToEmail = args.options.enableAssignToEmail || false;
    telemetryProps.enableAttachments = args.options.enableAttachments || false;
    telemetryProps.enableDeployWithDependentList = args.options.enableDeployWithDependentList || false;
    telemetryProps.enableFolderCreation = args.options.enableFolderCreation || false;
    telemetryProps.enableMinorVersions = args.options.enableMinorVersions || false;
    telemetryProps.enableModeration = args.options.enableModeration || false;
    telemetryProps.enablePeopleSelector = args.options.enablePeopleSelector || false;
    telemetryProps.enableResourceSelector = args.options.enableResourceSelector || false;
    telemetryProps.enableSchemaCaching = args.options.enableSchemaCaching || false;
    telemetryProps.enableSyndication = args.options.enableSyndication || false;
    telemetryProps.enableThrottling = args.options.enableThrottling || false;
    telemetryProps.enableVersioning = args.options.enableVersioning || false;
    telemetryProps.enforceDataValidation = args.options.enforceDataValidation || false;
    telemetryProps.excludeFromOfflineClient = args.options.excludeFromOfflineClient || false;
    telemetryProps.fetchPropertyBagForListView = args.options.fetchPropertyBagForListView || false;
    telemetryProps.followable = args.options.followable || false;
    telemetryProps.forceCheckout = args.options.forceCheckout || false;
    telemetryProps.forceDefaultContentType = args.options.forceDefaultContentType || false;
    telemetryProps.hidden = args.options.hidden || false;
    telemetryProps.includedInMyFilesScope = args.options.includedInMyFilesScope || false;
    telemetryProps.irmEnabled = args.options.irmEnabled || false;
    telemetryProps.irmExpire = args.options.irmExpire || false;
    telemetryProps.irmReject = args.options.irmReject || false;
    telemetryProps.isApplicationList = args.options.isApplicationList || false;
    telemetryProps.listExperienceOptions = (!(!args.options.listExperienceOptions)).toString();
    telemetryProps.majorVersionLimit = (!(!args.options.majorVersionLimit)).toString();
    telemetryProps.majorWithMinorVersionsLimit = (!(!args.options.majorWithMinorVersionsLimit)).toString();
    telemetryProps.multipleDataList = args.options.multipleDataList || false;
    telemetryProps.navigateForFormsPages = args.options.navigateForFormsPages || false;
    telemetryProps.needUpdateSiteClientTag = args.options.needUpdateSiteClientTag || false;
    telemetryProps.noCrawl = args.options.noCrawl || false;
    telemetryProps.onQuickLaunch = args.options.onQuickLaunch || false;
    telemetryProps.ordered = args.options.ordered || false;
    telemetryProps.parserDisabled = args.options.parserDisabled || false;
    telemetryProps.readOnlyUI = args.options.readOnlyUI || false;
    telemetryProps.readSecurity = (!(!args.options.readSecurity)).toString();
    telemetryProps.requestAccessEnabled = args.options.requestAccessEnabled || false;
    telemetryProps.restrictUserUpdates = args.options.restrictUserUpdates || false;
    telemetryProps.sendToLocationName = (!(!args.options.sendToLocationName)).toString();
    telemetryProps.sendToLocationUrl = (!(!args.options.sendToLocationUrl)).toString();
    telemetryProps.showUser = args.options.showUser || false;
    telemetryProps.useFormsForDisplay = args.options.useFormsForDisplay || false;    
    telemetryProps.validationFormula = (!(!args.options.validationFormula)).toString();
    telemetryProps.validationMessage = (!(!args.options.validationMessage)).toString();
    telemetryProps.writeSecurity = (!(!args.options.writeSecurity)).toString();

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

        //return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
      //})
      //.then((res: ContextInfo): Promise<ListInstance> => {
        // if (this.debug) {
        //   cmd.log('Response:');
        //   cmd.log(res);
        //   cmd.log('');
        // }

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
        description: 'Boolean value specifying whether the list can be deleted'
      },   
      {
        option: '--allowEveryoneViewItems [allowEveryoneViewItems]',
        description: 'Boolean value specifying whether everyone can view documents in the document library or attachments to items in the list'
      },
      {
        option: '--allowMultiResponses [allowMultiResponses]',
        description: 'Boolean value specifying whether users are allowed to give multiple responses to the survey'
      },
      {
        option: '--contentTypesEnabled [contentTypesEnabled]',
        description: 'Boolean value specifying whether content types are enabled for the list'
      },
      {
        option: '--crawlNonDefaultViews [crawlNonDefaultViews]',
        description: 'Boolean value specifying whether to crawl non default views'
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
        description: 'Property for assigning or retrieving grid editing on the list'
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
        description: 'Boolean value specifying whether e-mail notification is enabled for the list'
      },
      {
        option: '--enableAttachments [enableAttachments]',
        description: 'Boolean value that specifies whether attachments can be added to items in the list'
      },
      {
        option: '--enableDeployWithDependentList [enableDeployWithDependentList]',
        description: ' Boolean value that specifies whether the list can be deployed with a dependent list'
      },
      {
        option: '--enableFolderCreation [enableFolderCreation]',
        description: 'Boolean value that specifies whether folders can be created for the list'
      },
      {
        option: '--enableMinorVersions [enableMinorVersions]',
        description: 'Boolean value that specifies whether minor versions are enabled when versioning is enabled for the document library'
      },
      {
        option: '--enableModeration [enableModeration]',
        description: 'Boolean value that specifies whether Content Approval is enabled for the list'
      },
      {
        option: '--enablePeopleSelector [enablePeopleSelector]',
        description: 'Enable user selector on event list'
      },
      {
        option: '--enableResourceSelector [enableResourceSelector]',
        description: 'Enables resource selector on an event list'
      },
      {
        option: '--enableSchemaCaching [enableSchemaCaching]',
        description: 'Boolean value specifying whether schema caching is enabled for the list'
      },
      {
        option: '--enableSyndication [enableSyndication]',
        description: 'Boolean value that specifies whether RSS syndication is enabled for the list'
      },
      {
        option: '--enableThrottling [enableThrottling]',
        description: 'Indicates whether throttling for this list is enabled or not'
      },
      {
        option: '--enableVersioning [enableVersioning]',
        description: 'Boolean value that specifies whether versioning is enabled for the document library.'
      },
      {
        option: '--enforceDataValidation [enforceDataValidation]',
        description: 'Value that indicates whether certain field properties are enforced when an item is added or updated'
      },
      {
        option: '--excludeFromOfflineClient [excludeFromOfflineClient]',
        description: 'Value that indicates whether the list should be downloaded to the client during offline synchronization'
      },
      {
        option: '--fetchPropertyBagForListView [fetchPropertyBagForListView]',
        description: 'Specifies whether property bag information, as part of the list schema JSON, is retrieved when the list is being rendered on the client'
      },
      {
        option: '--followable [followable]',
        description: 'Can a list be followed in an activity feed?'
      },
      {
        option: '--forceCheckout [forceCheckout]',
        description: 'Boolean value that specifies whether forced checkout is enabled for the document library'
      },
      {
        option: '--forceDefaultContentType [forceDefaultContentType]',
        description: 'Specifies whether we want to return the default Document root content type'
      },
      {
        option: '--hidden [hidden]',
        description: 'Boolean value that specifies whether the list is hidden'
      },
      {
        option: '--includedInMyFilesScope [includedInMyFilesScope]',
        description: 'Specifies whether this list is accessible to an app principal that has been granted an OAuth scope that contains the string “myfiles” by a case-insensitive comparison when the current user is a site collection administrator of the personal site that contains the list'
      },
      {
        option: '--irmEnabled [irmEnabled]',
        description: 'Gets or sets a Boolean value that specifies whether Information Rights Management (IRM) is enabled for the list'
      },
      {
        option: '--irmExpire [irmExpire]',
        description: 'Gets or sets a Boolean value that specifies whether Information Rights Management (IRM) expiration is enabled for the list'
      },
      {
        option: '--irmReject [irmReject]',
        description: 'Gets or sets a Boolean value that specifies whether Information Rights Management (IRM) rejection is enabled for the list'
      },
      {
        option: '--isApplicationList [isApplicationList]',
        description: 'Indicates whether this list should be treated as a top level navigation object or not'
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
        description: 'Gets or sets a Boolean value that specifies whether the list in a Meeting Workspace site contains data for multiple meeting instances within the site'
      },
      {
        option: '--navigateForFormsPages [navigateForFormsPages]',
        description: 'Indicates whether to navigate for forms pages or use a modal dialog'
      },
      {
        option: '--needUpdateSiteClientTag [needUpdateSiteClientTag]',
        description: 'A boolean value that determines whether to editing documents in this list should increment the ClientTag for the site. The tag is used to allow clients to cache JS/CSS/resources that are retrieved from the Content DB, including custom CSR templates.'
      },
      {
        option: '--noCrawl [noCrawl]',
        description: 'Gets or sets a Boolean value specifying whether crawling is enabled for the list'
      },
      {
        option: '--onQuickLaunch [onQuickLaunch]',
        description: 'Gets or sets a Boolean value that specifies whether the list appears on the Quick Launch area of the home page'
      },
      {
        option: '--ordered [ordered]',
        description: 'Gets or sets a Boolean value that specifies whether the option to allow users to reorder items in the list is available on the Edit View page for the list'
      },
      {
        option: '-parserDisabled [parserDisabled]',
        description: 'Gets or sets a Boolean value that specifies whether the parser should be disabled'
      },
      {
        option: '--readOnlyUI [readOnlyUI]',
        description: 'A boolean value that indicates whether the UI for this list should be presented in a read-only fashion. This will not affect security nor will it actually prevent changes to the list from occurring - it only affects the way the UI is displayed'
      },
      {
        option: '--readSecurity [readSecurity]',
        description: 'Gets or sets the Read security setting for the list. Valid values are 1 (All users have Read access to all items)|2 (Users have Read access only to items that they create)',
        autocomplete: ['1', '2']
      },
      {
        option: '--requestAccessEnabled [requestAccessEnabled]',
        description: 'Gets or sets a Boolean value that specifies whether the option to allow users to request access to the list is available'
      },
      {
        option: '--restrictUserUpdates [restrictUserUpdates]',
        description: 'A boolean value that indicates whether the this list is a restricted one or not The value can\'t be changed if there are existing items in the list'
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
        description: 'Gets or sets a Boolean value that specifies whether names of users are shown in the results of the survey'
      },
      {
        option: '--useFormsForDisplay [useFormsForDisplay]',
        description: 'Indicates whether forms should be considered for display context or not'
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
  
    Add a list with title ${chalk.grey('Announcements')}, baseTemplate ${chalk.grey('107')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.LIST_ADD} --title 'DemoList' --baseTemplate 107 --webUrl https://contoso.sharepoint.com/sites/project-x

    Add a list with title ${chalk.grey('Announcements')}, baseTemplate ${chalk.grey('107')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')} 
    with schemaXml ${chalk.grey('<List DocTemplateUrl="" DefaultViewUrl="" MobileDefaultViewUrl="" ID="{92FF93AB-920E-4D33-AE42-58B5E245BEFF}" Title="Announcements" Description="" ImageUrl="/_layouts/15/images/itann.png?rev=44" Name="{92FF93AB-920E-4D33-AE42-58B5E245BEFF}" BaseType="0" FeatureId="{00BFEA71-D1CE-42DE-9C63-A44004CE0104}" ServerTemplate="104" Created="20161221 20:02:12" Modified="20180110 19:35:15" LastDeleted="20161221 20:02:12" Version="0" Direction="none" ThumbnailSize="0" WebImageWidth="0" WebImageHeight="0" Flags="536875008" ItemCount="1" AnonymousPermMask="0" RootFolder="/sites/project-x/Lists/Announcements"      ReadSecurity="1" WriteSecurity="1" Author="3" EventSinkAssembly="" EventSinkClass="" EventSinkData="" EmailAlias="" WebFullUrl="/sites/project-x" WebId="7694137e-7038-4831-a1bd-218b28fe5d34" SendToLocation="" ScopeId="92facaf9-8d7a-40eb-9e69-362c91513cbd" MajorVersionLimit="0" MajorWithMinorVersionsLimit="0" WorkFlowId="00000000-0000-0000-0000-000000000000" HasUniqueScopes="False" NoThrottleListOperations="False" HasRelatedLists="False" Followable="False" Acl="" Flags2="0" RootFolderId="d4d67cc1-ad6e-4293-b039-ea49263d195f" ComplianceTag="" ComplianceFlags="0" UserModified="20161221 20:03:00" ListSchemaVersion="3" AclVersion="" AllowDeletion="True" AllowMultiResponses="False" EnableAttachments="True" EnableModeration="False" EnableVersioning="False" HasExternalDataSource="False" Hidden="False" MultipleDataList="False" Ordered="False" ShowUser="True" EnablePeopleSelector="False" EnableResourceSelector="False" EnableMinorVersion="False" RequireCheckout="False" ThrottleListOperations="False" ExcludeFromOfflineClient="False" CanOpenFileAsync="True" EnableFolderCreation="False" IrmEnabled="False" IrmSyncable="False" IsApplicationList="False" PreserveEmptyValues="False" StrictTypeCoercion="False" EnforceDataValidation="False" MaxItemsPerThrottledOperation="5000"></List>')}
      ${chalk.grey(config.delimiter)} ${commands.LIST_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --title Announcements --baseTemplate Announcements --schemaXml '<List DocTemplateUrl="" DefaultViewUrl="" MobileDefaultViewUrl="" ID="{92FF93AB-920E-4D33-AE42-58B5E245BEFF}" Title="Announcements" Description="" ImageUrl="/_layouts/15/images/itann.png?rev=44" Name="{92FF93AB-920E-4D33-AE42-58B5E245BEFF}" BaseType="0" FeatureId="{00BFEA71-D1CE-42DE-9C63-A44004CE0104}" ServerTemplate="104" Created="20161221 20:02:12" Modified="20180110 19:35:15" LastDeleted="20161221 20:02:12" Version="0" Direction="none" ThumbnailSize="0" WebImageWidth="0" WebImageHeight="0" Flags="536875008" ItemCount="1" AnonymousPermMask="0" RootFolder="/sites/project-x/Lists/Announcements"      ReadSecurity="1" WriteSecurity="1" Author="3" EventSinkAssembly="" EventSinkClass="" EventSinkData="" EmailAlias="" WebFullUrl="/sites/project-x" WebId="7694137e-7038-4831-a1bd-218b28fe5d34" SendToLocation="" ScopeId="92facaf9-8d7a-40eb-9e69-362c91513cbd" MajorVersionLimit="0" MajorWithMinorVersionsLimit="0" WorkFlowId="00000000-0000-0000-0000-000000000000" HasUniqueScopes="False" NoThrottleListOperations="False" HasRelatedLists="False" Followable="False" Acl="" Flags2="0" RootFolderId="d4d67cc1-ad6e-4293-b039-ea49263d195f" ComplianceTag="" ComplianceFlags="0" UserModified="20161221 20:03:00" ListSchemaVersion="3" AclVersion="" AllowDeletion="True" AllowMultiResponses="False" EnableAttachments="True" EnableModeration="False" EnableVersioning="False" HasExternalDataSource="False" Hidden="False" MultipleDataList="False" Ordered="False" ShowUser="True" EnablePeopleSelector="False" EnableResourceSelector="False" EnableMinorVersion="False" RequireCheckout="False" ThrottleListOperations="False" ExcludeFromOfflineClient="False" CanOpenFileAsync="True" EnableFolderCreation="False" IrmEnabled="False" IrmSyncable="False" IsApplicationList="False" PreserveEmptyValues="False" StrictTypeCoercion="False" EnforceDataValidation="False" MaxItemsPerThrottledOperation="5000"></List>'
    
    Add a list with title ${chalk.grey('Announcements')}, baseTemplate ${chalk.grey('107')} in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
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