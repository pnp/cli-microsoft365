import { Logger } from '../../../../cli';
import {
  CommandOption,
  CommandTypes
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { DraftVisibilityType } from './DraftVisibilityType';
import { ListExperience } from './ListExperience';
import { ListInstance } from "./ListInstance";
import { ListTemplateType } from './ListTemplateType';

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

    for (const template in ListTemplateType) {
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

    for (const draftType in DraftVisibilityType) {
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

    for (const experience in ListExperience) {
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Creating list in site at ${args.options.webUrl}...`);
    }

    const requestBody: any = this.mapRequestBody(args.options);

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/lists`,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    request
      .post<ListInstance>(requestOptions)
      .then((listInstance: ListInstance): void => {
        logger.log(listInstance);

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-t, --title <title>'
      },
      {
        option: '--baseTemplate <baseTemplate>',
        autocomplete: this.listTemplateTypeMap
      },
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--description [description]'
      },
      {
        option: '--templateFeatureId [templateFeatureId]'
      },
      {
        option: '--schemaXml [schemaXml]'
      },
      {
        option: '--allowDeletion [allowDeletion]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowEveryoneViewItems [allowEveryoneViewItems]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowMultiResponses [allowMultiResponses]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--contentTypesEnabled [contentTypesEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--crawlNonDefaultViews [crawlNonDefaultViews]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--defaultContentApprovalWorkflowId [defaultContentApprovalWorkflowId]'
      },
      {
        option: '--defaultDisplayFormUrl [defaultDisplayFormUrl]'
      },
      {
        option: '--defaultEditFormUrl [defaultEditFormUrl]'
      },
      {
        option: '--direction [direction]',
        autocomplete: ['NONE', 'LTR', 'RTL']
      },
      {
        option: '--disableGridEditing [disableGridEditing]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--draftVersionVisibility [draftVersionVisibility]',
        autocomplete: this.draftVisibilityTypeMap
      },
      {
        option: '--emailAlias [emailAlias]'
      },
      {
        option: '--enableAssignToEmail [enableAssignToEmail]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableAttachments [enableAttachments]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableDeployWithDependentList [enableDeployWithDependentList]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableFolderCreation [enableFolderCreation]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableMinorVersions [enableMinorVersions]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableModeration [enableModeration]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enablePeopleSelector [enablePeopleSelector]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableResourceSelector [enableResourceSelector]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableSchemaCaching [enableSchemaCaching]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableSyndication [enableSyndication]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableThrottling [enableThrottling]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableVersioning [enableVersioning]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enforceDataValidation [enforceDataValidation]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--excludeFromOfflineClient [excludeFromOfflineClient]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--fetchPropertyBagForListView [fetchPropertyBagForListView]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--followable [followable]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--forceCheckout [forceCheckout]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--forceDefaultContentType [forceDefaultContentType]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--hidden [hidden]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--includedInMyFilesScope [includedInMyFilesScope]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--irmEnabled [irmEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--irmExpire [irmExpire]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--irmReject [irmReject]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--isApplicationList [isApplicationList]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--listExperienceOptions [listExperienceOptions]',
        autocomplete: this.listExperienceMap
      },
      {
        option: '--majorVersionLimit [majorVersionLimit]'
      },
      {
        option: '--majorWithMinorVersionsLimit [majorWithMinorVersionsLimit]'
      },
      {
        option: '--multipleDataList [multipleDataList]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--navigateForFormsPages [navigateForFormsPages]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--needUpdateSiteClientTag [needUpdateSiteClientTag]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--noCrawl [noCrawl]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--onQuickLaunch [onQuickLaunch]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--ordered [ordered]',
        autocomplete: ['true', 'false']
      },
      {
        option: '-parserDisabled [parserDisabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--readOnlyUI [readOnlyUI]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--readSecurity [readSecurity]',
        autocomplete: ['1', '2']
      },
      {
        option: '--requestAccessEnabled [requestAccessEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--restrictUserUpdates [restrictUserUpdates]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--sendToLocationName [sendToLocationName]'
      },
      {
        option: '--sendToLocationUrl [sendToLocationUrl]'
      },
      {
        option: '--showUser [showUser]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--useFormsForDisplay [useFormsForDisplay]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--validationFormula [validationFormula]'
      },
      {
        option: '--validationMessage [validationMessage]'
      },
      {
        option: '--writeSecurity [writeSecurity]',
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

  public validate(args: CommandArgs): boolean | string {
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
        return `${value} in option ${option} is not a valid boolean value`;
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