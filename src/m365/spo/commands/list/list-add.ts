import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { DraftVisibilityType } from './DraftVisibilityType.js';
import { ListExperience } from './ListExperience.js';
import { ListInstance } from "./ListInstance.js";
import { ListTemplateType } from './ListTemplateType.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  title: z.string().alias('t'),
  baseTemplate: z.enum([
    'GenericList',
    'DocumentLibrary',
    'Survey',
    'Links',
    'Announcements',
    'Contacts',
    'Events',
    'Tasks',
    'DiscussionBoard',
    'PictureLibrary',
    'DataSources',
    'XmlForm',
    'NoCodeWorkflows',
    'WorkflowProcess',
    'WebPageLibrary',
    'CustomGrid',
    'WorkflowHistory',
    'GanttTasks',
    'IssuesTracking'
  ]).optional(),
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `${e.input} is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  description: z.string().optional(),
  templateFeatureId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} in option templateFeatureId is not a valid GUID`
    })
    .optional(),
  allowDeletion: z.boolean().optional().autocomplete(['true', 'false']),
  allowEveryoneViewItems: z.boolean().optional().autocomplete(['true', 'false']),
  allowMultiResponses: z.boolean().optional().autocomplete(['true', 'false']),
  contentTypesEnabled: z.boolean().optional().autocomplete(['true', 'false']),
  crawlNonDefaultViews: z.boolean().optional().autocomplete(['true', 'false']),
  defaultContentApprovalWorkflowId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} in option defaultContentApprovalWorkflowId is not a valid GUID`
    })
    .optional(),
  defaultDisplayFormUrl: z.string().optional(),
  defaultEditFormUrl: z.string().optional(),
  direction: z.enum(['NONE', 'LTR', 'RTL']).optional(),
  disableCommenting: z.boolean().optional().autocomplete(['true', 'false']),
  disableGridEditing: z.boolean().optional().autocomplete(['true', 'false']),
  draftVersionVisibility: z.enum(['Reader', 'Author', 'Approver']).optional(),
  emailAlias: z.string().optional(),
  enableAssignToEmail: z.boolean().optional().autocomplete(['true', 'false']),
  enableAttachments: z.boolean().optional().autocomplete(['true', 'false']),
  enableDeployWithDependentList: z.boolean().optional().autocomplete(['true', 'false']),
  enableFolderCreation: z.boolean().optional().autocomplete(['true', 'false']),
  enableMinorVersions: z.boolean().optional().autocomplete(['true', 'false']),
  enableModeration: z.boolean().optional().autocomplete(['true', 'false']),
  enablePeopleSelector: z.boolean().optional().autocomplete(['true', 'false']),
  enableResourceSelector: z.boolean().optional().autocomplete(['true', 'false']),
  enableSchemaCaching: z.boolean().optional().autocomplete(['true', 'false']),
  enableSyndication: z.boolean().optional().autocomplete(['true', 'false']),
  enableThrottling: z.boolean().optional().autocomplete(['true', 'false']),
  enableVersioning: z.boolean().optional().autocomplete(['true', 'false']),
  enforceDataValidation: z.boolean().optional().autocomplete(['true', 'false']),
  excludeFromOfflineClient: z.boolean().optional().autocomplete(['true', 'false']),
  fetchPropertyBagForListView: z.boolean().optional().autocomplete(['true', 'false']),
  followable: z.boolean().optional().autocomplete(['true', 'false']),
  forceCheckout: z.boolean().optional().autocomplete(['true', 'false']),
  forceDefaultContentType: z.boolean().optional().autocomplete(['true', 'false']),
  hidden: z.boolean().optional().autocomplete(['true', 'false']),
  includedInMyFilesScope: z.boolean().optional().autocomplete(['true', 'false']),
  irmEnabled: z.boolean().optional().autocomplete(['true', 'false']),
  irmExpire: z.boolean().optional().autocomplete(['true', 'false']),
  irmReject: z.boolean().optional().autocomplete(['true', 'false']),
  isApplicationList: z.boolean().optional().autocomplete(['true', 'false']),
  listExperienceOptions: z.enum(['Auto', 'NewExperience', 'ClassicExperience']).optional(),
  majorVersionLimit: z.number().optional(),
  majorWithMinorVersionsLimit: z.number().optional(),
  multipleDataList: z.boolean().optional().autocomplete(['true', 'false']),
  navigateForFormsPages: z.boolean().optional().autocomplete(['true', 'false']),
  needUpdateSiteClientTag: z.boolean().optional().autocomplete(['true', 'false']),
  noCrawl: z.boolean().optional().autocomplete(['true', 'false']),
  onQuickLaunch: z.boolean().optional().autocomplete(['true', 'false']),
  ordered: z.boolean().optional().autocomplete(['true', 'false']),
  parserDisabled: z.boolean().optional().autocomplete(['true', 'false']),
  readOnlyUI: z.boolean().optional().autocomplete(['true', 'false']),
  readSecurity: z.number()
    .refine(v => v === 1 || v === 2, {
      error: e => `${e.input} is not a valid readSecurity value. Allowed values are 1|2`
    })
    .optional(),
  requestAccessEnabled: z.boolean().optional().autocomplete(['true', 'false']),
  restrictUserUpdates: z.boolean().optional().autocomplete(['true', 'false']),
  sendToLocationName: z.string().optional(),
  sendToLocationUrl: z.string().optional(),
  showUser: z.boolean().optional().autocomplete(['true', 'false']),
  useFormsForDisplay: z.boolean().optional().autocomplete(['true', 'false']),
  validationFormula: z.string().optional(),
  validationMessage: z.string().optional(),
  writeSecurity: z.number()
    .refine(v => v === 1 || v === 2 || v === 4, {
      error: e => `${e.input} is not a valid writeSecurity value. Allowed values are 1|2|4`
    })
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListAddCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_ADD;
  }

  public get description(): string {
    return 'Creates list in the specified site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => !opts.emailAlias || opts.enableAssignToEmail === true, {
        error: 'emailAlias could not be set if enableAssignToEmail is not set to true. Please set enableAssignToEmail.'
      })
      .refine(opts => opts.majorVersionLimit === undefined || opts.enableVersioning === true, {
        error: 'majorVersionLimit option is only valid in combination with enableVersioning.'
      })
      .refine(opts => opts.majorWithMinorVersionsLimit === undefined || opts.enableMinorVersions === true || opts.enableModeration === true, {
        error: 'majorWithMinorVersionsLimit option is only valid in combination with enableMinorVersions or enableModeration.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Creating list in site at ${args.options.webUrl}...`);
    }

    const requestBody: any = this.mapRequestBody(args.options);

    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/web/lists`,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    try {
      const listInstance = await request.post<ListInstance>(requestOptions);
      await logger.log(listInstance);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {
      Title: options.title,
      BaseTemplate: options.baseTemplate ? ListTemplateType[(options.baseTemplate.trim() as keyof typeof ListTemplateType)].valueOf() : ListTemplateType.GenericList
    };

    if (options.description) {
      requestBody.Description = options.description;
    }

    if (options.templateFeatureId) {
      requestBody.TemplateFeatureId = options.templateFeatureId;
    }

    if (options.allowDeletion !== undefined) {
      requestBody.AllowDeletion = options.allowDeletion;
    }

    if (options.allowEveryoneViewItems !== undefined) {
      requestBody.AllowEveryoneViewItems = options.allowEveryoneViewItems;
    }

    if (options.allowMultiResponses !== undefined) {
      requestBody.AllowMultiResponses = options.allowMultiResponses;
    }

    if (options.contentTypesEnabled !== undefined) {
      requestBody.ContentTypesEnabled = options.contentTypesEnabled;
    }

    if (options.crawlNonDefaultViews !== undefined) {
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

    if (options.disableCommenting !== undefined) {
      requestBody.DisableCommenting = options.disableCommenting;
    }

    if (options.disableGridEditing !== undefined) {
      requestBody.DisableGridEditing = options.disableGridEditing;
    }

    if (options.draftVersionVisibility) {
      requestBody.DraftVersionVisibility = DraftVisibilityType[(options.draftVersionVisibility.trim() as keyof typeof DraftVisibilityType)];
    }

    if (options.emailAlias) {
      requestBody.EmailAlias = options.emailAlias;
    }

    if (options.enableAssignToEmail !== undefined) {
      requestBody.EnableAssignToEmail = options.enableAssignToEmail;
    }

    if (options.enableAttachments !== undefined) {
      requestBody.EnableAttachments = options.enableAttachments;
    }

    if (options.enableDeployWithDependentList !== undefined) {
      requestBody.EnableDeployWithDependentList = options.enableDeployWithDependentList;
    }

    if (options.enableFolderCreation !== undefined) {
      requestBody.EnableFolderCreation = options.enableFolderCreation;
    }

    if (options.enableMinorVersions !== undefined) {
      requestBody.EnableMinorVersions = options.enableMinorVersions;
    }

    if (options.enableModeration !== undefined) {
      requestBody.EnableModeration = options.enableModeration;
    }

    if (options.enablePeopleSelector !== undefined) {
      requestBody.EnablePeopleSelector = options.enablePeopleSelector;
    }

    if (options.enableResourceSelector !== undefined) {
      requestBody.EnableResourceSelector = options.enableResourceSelector;
    }

    if (options.enableSchemaCaching !== undefined) {
      requestBody.EnableSchemaCaching = options.enableSchemaCaching;
    }

    if (options.enableSyndication !== undefined) {
      requestBody.EnableSyndication = options.enableSyndication;
    }

    if (options.enableThrottling !== undefined) {
      requestBody.EnableThrottling = options.enableThrottling;
    }

    if (options.enableVersioning !== undefined) {
      requestBody.EnableVersioning = options.enableVersioning;
    }

    if (options.enforceDataValidation !== undefined) {
      requestBody.EnforceDataValidation = options.enforceDataValidation;
    }

    if (options.excludeFromOfflineClient !== undefined) {
      requestBody.ExcludeFromOfflineClient = options.excludeFromOfflineClient;
    }

    if (options.fetchPropertyBagForListView !== undefined) {
      requestBody.FetchPropertyBagForListView = options.fetchPropertyBagForListView;
    }

    if (options.followable !== undefined) {
      requestBody.Followable = options.followable;
    }

    if (options.forceCheckout !== undefined) {
      requestBody.ForceCheckout = options.forceCheckout;
    }

    if (options.forceDefaultContentType !== undefined) {
      requestBody.ForceDefaultContentType = options.forceDefaultContentType;
    }

    if (options.hidden !== undefined) {
      requestBody.Hidden = options.hidden;
    }

    if (options.includedInMyFilesScope !== undefined) {
      requestBody.IncludedInMyFilesScope = options.includedInMyFilesScope;
    }

    if (options.irmEnabled !== undefined) {
      requestBody.IrmEnabled = options.irmEnabled;
    }

    if (options.irmExpire !== undefined) {
      requestBody.IrmExpire = options.irmExpire;
    }

    if (options.irmReject !== undefined) {
      requestBody.IrmReject = options.irmReject;
    }

    if (options.isApplicationList !== undefined) {
      requestBody.IsApplicationList = options.isApplicationList;
    }

    if (options.listExperienceOptions) {
      requestBody.ListExperienceOptions = ListExperience[(options.listExperienceOptions.trim() as keyof typeof ListExperience)];
    }

    if (options.majorVersionLimit) {
      requestBody.MajorVersionLimit = options.majorVersionLimit;
    }

    if (options.majorWithMinorVersionsLimit) {
      requestBody.MajorWithMinorVersionsLimit = options.majorWithMinorVersionsLimit;
    }

    if (options.multipleDataList !== undefined) {
      requestBody.MultipleDataList = options.multipleDataList;
    }

    if (options.navigateForFormsPages !== undefined) {
      requestBody.NavigateForFormsPages = options.navigateForFormsPages;
    }

    if (options.needUpdateSiteClientTag !== undefined) {
      requestBody.NeedUpdateSiteClientTag = options.needUpdateSiteClientTag;
    }

    if (options.noCrawl !== undefined) {
      requestBody.NoCrawl = options.noCrawl;
    }

    if (options.onQuickLaunch !== undefined) {
      requestBody.OnQuickLaunch = options.onQuickLaunch;
    }

    if (options.ordered !== undefined) {
      requestBody.Ordered = options.ordered;
    }

    if (options.parserDisabled !== undefined) {
      requestBody.ParserDisabled = options.parserDisabled;
    }

    if (options.readOnlyUI !== undefined) {
      requestBody.ReadOnlyUI = options.readOnlyUI;
    }

    if (options.readSecurity) {
      requestBody.ReadSecurity = options.readSecurity;
    }

    if (options.requestAccessEnabled !== undefined) {
      requestBody.RequestAccessEnabled = options.requestAccessEnabled;
    }

    if (options.restrictUserUpdates !== undefined) {
      requestBody.RestrictUserUpdates = options.restrictUserUpdates;
    }

    if (options.sendToLocationName) {
      requestBody.SendToLocationName = options.sendToLocationName;
    }

    if (options.sendToLocationUrl) {
      requestBody.SendToLocationUrl = options.sendToLocationUrl;
    }

    if (options.showUser !== undefined) {
      requestBody.ShowUser = options.showUser;
    }

    if (options.useFormsForDisplay !== undefined) {
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

export default new SpoListAddCommand();