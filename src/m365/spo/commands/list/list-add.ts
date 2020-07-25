import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate,
  CommandTypes
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ListInstance } from "./ListInstance";
import { ListTemplateType } from './ListTemplateType';
import { DraftVisibilityType } from './DraftVisibilityType';
import { ListExperience } from './ListExperience';
import { CommandInstance } from '../../../../cli';

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

class SpoListAddCommand extends SpoCommand {
  private static booleanOptions: string[] = [
    'allowDeletion',
    'allowEveryoneViewItems',
    'allowMultiResponses',
    'contentTypesEnabled',
    'crawlNonDefaultViews',
    'disableGridEditing',
    'enableAssignToEmail',
    'enableAttachments',
    'enableDeployWithDependentList',
    'enableFolderCreation',
    'enableMinorVersions',
    'enableModeration',
    'enablePeopleSelector',
    'enableResourceSelector',
    'enableSchemaCaching',
    'enableSyndication',
    'enableThrottling',
    'enableVersioning',
    'enforceDataValidation',
    'excludeFromOfflineClient',
    'fetchPropertyBagForListView',
    'followable',
    'forceCheckout',
    'forceDefaultContentType',
    'hidden',
    'includedInMyFilesScope',
    'irmEnabled',
    'irmExpire',
    'irmReject',
    'isApplicationList',
    'multipleDataList',
    'navigateForFormsPages',
    'needUpdateSiteClientTag',
    'noCrawl',
    'onQuickLaunch',
    'ordered',
    'parserDisabled',
    'readOnlyUI',
    'requestAccessEnabled',
    'restrictUserUpdates',
    'showUser',
    'useFormsForDisplay'
  ];

  public get name(): string {
    return commands.LIST_ADD;
  }

  public get description(): string {
    return 'Creates list in the specified site';
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

    // add properties with identifiable data
    [
      'description',
      'templateFeatureId',
      'schemaXml',
      'defaultContentApprovalWorkflowId',
      'defaultDisplayFormUrl',
      'defaultEditFormUrl',
      'emailAlias',
      'sendToLocationName',
      'sendToLocationUrl',
      'validationFormula',
      'validationMessage'
    ].forEach(o => {
      const value: any = (args.options as any)[o];
      if (value) {
        telemetryProps[o] = (typeof value !== 'undefined').toString();
      }
    });

    // add boolean values
    SpoListAddCommand.booleanOptions.forEach(o => {
      const value: any = (args.options as any)[o];
      if (value) {
        telemetryProps[o] = (value === 'true').toString();
      }
    });

    // add properties with non-identifiable data
    [
      'baseTemplate',
      'direction',
      'draftVersionVisibility',
      'listExperienceOptions',
      'majorVersionLimit',
      'majorWithMinorVersionsLimit',
      'readSecurity',
      'writeSecurity'
    ].forEach(o => {
      const value: any = (args.options as any)[o];
      if (value) {
        telemetryProps[o] = value.toString();
      }
    });

    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Creating list in site at ${args.options.webUrl}...`);
    }

    const requestBody: any = this.mapRequestBody(args.options);

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/lists`,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      body: requestBody,
      json: true
    };

    request
      .post<ListInstance>(requestOptions)
      .then((listInstance: ListInstance): void => {
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

  public types(): CommandTypes {
    return {
      string: SpoListAddCommand.booleanOptions.concat([
        'baseTemplate',
        'webUrl',
        'templateFeatureId',
        'defaultContentApprovalWorkflowId',
        'draftVersionVisibility',
        'listExperienceOptions'
      ])
    };
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      const template: ListTemplateType = ListTemplateType[(args.options.baseTemplate.trim() as keyof typeof ListTemplateType)];
      if (!template) {
        return `${args.options.baseTemplate} is not a valid baseTemplate value`;
      }

      for (let i = 0; i < SpoListAddCommand.booleanOptions.length; i++) {
        const option: string = SpoListAddCommand.booleanOptions[i];
        const value: string | undefined = (args.options as any)[option];
        if (value && !Utils.isValidBoolean(value)) {
          return `${value} in option ${option} is not a valid boolean value`
        }
      }

      if (args.options.templateFeatureId &&
        !Utils.isValidGuid(args.options.templateFeatureId)) {
        return `${args.options.templateFeatureId} in option templateFeatureId is not a valid GUID`;
      }

      if (args.options.defaultContentApprovalWorkflowId &&
        !Utils.isValidGuid(args.options.defaultContentApprovalWorkflowId)) {
        return `${args.options.defaultContentApprovalWorkflowId} in option defaultContentApprovalWorkflowId is not a valid GUID`;
      }

      if (args.options.direction &&
        ['NONE', 'LTR', 'RTL'].indexOf(args.options.direction) === -1) {
        return `${args.options.direction} is not a valid direction value. Allowed values are NONE|LTR|RTL`;
      }

      if (args.options.draftVersionVisibility) {
        const draftType: DraftVisibilityType = DraftVisibilityType[(args.options.draftVersionVisibility.trim() as keyof typeof DraftVisibilityType)];

        if (!draftType) {
          return `${args.options.draftVersionVisibility} is not a valid draftVisibilityType value`;
        }
      }

      if (args.options.emailAlias && args.options.enableAssignToEmail !== 'true') {
        return `emailAlias could not be set if enableAssignToEmail is not set to true. Please set enableAssignToEmail.`;
      }

      if (args.options.listExperienceOptions) {
        const experience: ListExperience = ListExperience[(args.options.listExperienceOptions.trim() as keyof typeof ListExperience)];

        if (!experience) {
          return `${args.options.listExperienceOptions} is not a valid listExperienceOptions value`;
        }
      }

      if (args.options.majorVersionLimit && args.options.enableVersioning !== 'true') {
        return `majorVersionLimit option is only valid in combination with enableVersioning.`;
      }

      if (args.options.majorWithMinorVersionsLimit &&
        args.options.enableMinorVersions !== 'true' &&
        args.options.enableModeration !== 'true') {
        return `majorWithMinorVersionsLimit option is only valid in combination with enableMinorVersions or enableModeration.`;
      }

      if (args.options.readSecurity &&
        args.options.readSecurity !== 1 &&
        args.options.readSecurity !== 2) {
        return `${args.options.readSecurity} is not a valid readSecurity value. Allowed values are 1|2`;
      }

      if (args.options.writeSecurity &&
        args.options.writeSecurity !== 1 &&
        args.options.writeSecurity !== 2 &&
        args.options.writeSecurity !== 4) {
        return `${args.options.writeSecurity} is not a valid writeSecurity value. Allowed values are 1|2|4`;
      }

      return true;
    };
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

module.exports = new SpoListAddCommand();