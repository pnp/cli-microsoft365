import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import config from '../../../../config.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { DraftVisibilityType } from './DraftVisibilityType.js';
import { ListExperience } from './ListExperience.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `${e.input} is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  id: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .optional(),
  title: z.string().optional(),
  url: z.string().optional(),
  newTitle: z.string().optional(),
  allowDeletion: z.boolean().optional(),
  allowEveryoneViewItems: z.boolean().optional(),
  allowMultiResponses: z.boolean().optional(),
  contentTypesEnabled: z.boolean().optional(),
  crawlNonDefaultViews: z.boolean().optional(),
  defaultContentApprovalWorkflowId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} in option defaultContentApprovalWorkflowId is not a valid GUID`
    })
    .optional(),
  defaultDisplayFormUrl: z.string().optional(),
  defaultEditFormUrl: z.string().optional(),
  description: z.string().optional(),
  direction: z.enum(['NONE', 'LTR', 'RTL']).optional(),
  disableCommenting: z.boolean().optional(),
  disableGridEditing: z.boolean().optional(),
  draftVersionVisibility: z.enum(['Reader', 'Author', 'Approver']).optional(),
  emailAlias: z.string().optional(),
  enableAssignToEmail: z.boolean().optional(),
  enableAttachments: z.boolean().optional(),
  enableDeployWithDependentList: z.boolean().optional(),
  enableFolderCreation: z.boolean().optional(),
  enableMinorVersions: z.boolean().optional(),
  enableModeration: z.boolean().optional(),
  enablePeopleSelector: z.boolean().optional(),
  enableResourceSelector: z.boolean().optional(),
  enableSchemaCaching: z.boolean().optional(),
  enableSyndication: z.boolean().optional(),
  enableThrottling: z.boolean().optional(),
  enableVersioning: z.boolean().optional(),
  enforceDataValidation: z.boolean().optional(),
  excludeFromOfflineClient: z.boolean().optional(),
  fetchPropertyBagForListView: z.boolean().optional(),
  followable: z.boolean().optional(),
  forceCheckout: z.boolean().optional(),
  forceDefaultContentType: z.boolean().optional(),
  hidden: z.boolean().optional(),
  includedInMyFilesScope: z.boolean().optional(),
  irmEnabled: z.boolean().optional(),
  irmExpire: z.boolean().optional(),
  irmReject: z.boolean().optional(),
  isApplicationList: z.boolean().optional(),
  listExperienceOptions: z.enum(['Auto', 'NewExperience', 'ClassicExperience']).optional(),
  majorVersionLimit: z.number().int().positive().optional(),
  majorWithMinorVersionsLimit: z.number().int().positive().optional(),
  multipleDataList: z.boolean().optional(),
  navigateForFormsPages: z.boolean().optional(),
  needUpdateSiteClientTag: z.boolean().optional(),
  noCrawl: z.boolean().optional(),
  onQuickLaunch: z.boolean().optional(),
  ordered: z.boolean().optional(),
  parserDisabled: z.boolean().optional(),
  readOnlyUI: z.boolean().optional(),
  readSecurity: z.number().refine(v => v === 1 || v === 2, {
    error: e => `${e.input} is not a valid readSecurity value. Allowed values are 1|2`
  }).optional(),
  requestAccessEnabled: z.boolean().optional(),
  restrictUserUpdates: z.boolean().optional(),
  sendToLocationName: z.string().optional(),
  sendToLocationUrl: z.string().optional(),
  showUser: z.boolean().optional(),
  templateFeatureId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} in option templateFeatureId is not a valid GUID`
    })
    .optional(),
  useFormsForDisplay: z.boolean().optional(),
  validationFormula: z.string().optional(),
  validationMessage: z.string().optional(),
  versionAutoExpireTrim: z.boolean().optional(),
  versionExpireAfterDays: z.number().int().positive().optional(),
  writeSecurity: z.number().refine(v => v === 1 || v === 2 || v === 4, {
    error: e => `${e.input} is not a valid writeSecurity value. Allowed values are 1|2|4`
  }).optional()
});

type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListSetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_SET;
  }

  public get description(): string {
    return 'Updates the settings of the specified list';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(
        (opts: Options) => [opts.id, opts.title, opts.url].filter(o => o !== undefined).length === 1,
        {
          error: 'Use one of the following options: id, title, or url.'
        }
      )
      .refine(
        (opts: Options) => !opts.emailAlias || opts.enableAssignToEmail === true,
        {
          error: 'emailAlias could not be set if enableAssignToEmail is not set to true. Please set enableAssignToEmail.'
        }
      )
      .refine(
        (opts: Options) => opts.majorWithMinorVersionsLimit === undefined || opts.enableMinorVersions === true || opts.enableModeration === true,
        {
          error: 'majorWithMinorVersionsLimit option is only valid in combination with enableMinorVersions or enableModeration.'
        }
      )
      .refine(
        (opts: Options) => opts.versionExpireAfterDays === undefined || opts.versionAutoExpireTrim !== true,
        {
          error: 'versionExpireAfterDays cannot be used together with versionAutoExpireTrim set to true.'
        }
      )
      .refine(
        (opts: Options) => {
          const identifierAndGlobalKeys = new Set(['webUrl', 'id', 'title', 'url', 'output', 'query', 'debug', 'verbose']);
          return Object.entries(opts).some(([key, value]) => !identifierAndGlobalKeys.has(key) && value !== undefined);
        },
        {
          error: 'Specify at least one option to update.'
        }
      );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Updating list in site at ${args.options.webUrl}...`);
    }

    const requestBody: any = this.mapRequestBody(args.options);

    let requestUrl = `${args.options.webUrl}/_api/web/`;
    if (args.options.id) {
      requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.id)}')/`;
    }
    else if (args.options.title) {
      requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(args.options.title)}')/`;
    }
    else if (args.options.url) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
      requestUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      method: 'POST',
      headers: {
        'X-HTTP-Method': 'MERGE',
        'If-Match': '*',
        'accept': 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    try {
      if (Object.keys(requestBody).length > 0) {
        await request.post(requestOptions);
      }

      if (args.options.versionExpireAfterDays !== undefined || args.options.versionAutoExpireTrim !== undefined) {
        await this.setVersionPolicies(args.options);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async setVersionPolicies(options: Options): Promise<void> {
    const digest = await spo.getRequestDigest(options.webUrl);

    let objectPaths = '';
    let actions = '';

    // SPContext.Current
    objectPaths += `<StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />`;
    actions += `<ObjectPath Id="1" ObjectPathId="0" />`;

    // Web
    objectPaths += `<Property Id="2" ParentId="0" Name="Web" />`;
    actions += `<ObjectPath Id="3" ObjectPathId="2" />`;

    let listObjectPathId: number;

    if (options.url) {
      const listServerRelativeUrl = urlUtil.getServerRelativePath(options.webUrl, options.url);
      objectPaths += `<Method Id="4" ParentId="2" Name="GetList"><Parameters><Parameter Type="String">${formatting.escapeXml(listServerRelativeUrl)}</Parameter></Parameters></Method>`;
      listObjectPathId = 4;
      actions += `<ObjectPath Id="5" ObjectPathId="4" />`;
    }
    else if (options.id) {
      objectPaths += `<Property Id="4" ParentId="2" Name="Lists" />`;
      actions += `<ObjectPath Id="5" ObjectPathId="4" />`;
      objectPaths += `<Method Id="6" ParentId="4" Name="GetById"><Parameters><Parameter Type="Guid">{${formatting.escapeXml(options.id)}}</Parameter></Parameters></Method>`;
      listObjectPathId = 6;
      actions += `<ObjectPath Id="7" ObjectPathId="6" />`;
    }
    else {
      const titleForLookup = options.newTitle ?? options.title!;
      objectPaths += `<Property Id="4" ParentId="2" Name="Lists" />`;
      actions += `<ObjectPath Id="5" ObjectPathId="4" />`;
      objectPaths += `<Method Id="6" ParentId="4" Name="GetByTitle"><Parameters><Parameter Type="String">${formatting.escapeXml(titleForLookup)}</Parameter></Parameters></Method>`;
      listObjectPathId = 6;
      actions += `<ObjectPath Id="7" ObjectPathId="6" />`;
    }

    const versionPoliciesId = listObjectPathId + 10;
    objectPaths += `<Property Id="${versionPoliciesId}" ParentId="${listObjectPathId}" Name="VersionPolicies" />`;
    actions += `<ObjectPath Id="${versionPoliciesId + 1}" ObjectPathId="${versionPoliciesId}" />`;

    let nextActionId = versionPoliciesId + 2;

    if (options.versionExpireAfterDays !== undefined) {
      actions += `<SetProperty Id="${nextActionId++}" ObjectPathId="${versionPoliciesId}" Name="DefaultTrimMode"><Parameter Type="Int32">1</Parameter></SetProperty>`;
      actions += `<SetProperty Id="${nextActionId++}" ObjectPathId="${versionPoliciesId}" Name="DefaultExpireAfterDays"><Parameter Type="Int32">${options.versionExpireAfterDays}</Parameter></SetProperty>`;
    }
    else if (options.versionAutoExpireTrim === true) {
      actions += `<SetProperty Id="${nextActionId++}" ObjectPathId="${versionPoliciesId}" Name="DefaultTrimMode"><Parameter Type="Int32">2</Parameter></SetProperty>`;
    }
    else if (options.versionAutoExpireTrim === false) {
      actions += `<SetProperty Id="${nextActionId++}" ObjectPathId="${versionPoliciesId}" Name="DefaultTrimMode"><Parameter Type="Int32">0</Parameter></SetProperty>`;
    }

    actions += `<Method Name="Update" Id="${nextActionId++}" ObjectPathId="${listObjectPathId}" />`;

    const csomRequestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': digest.FormDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${actions}</Actions><ObjectPaths>${objectPaths}</ObjectPaths></Request>`
    };

    const res = await request.post<string>(csomRequestOptions);
    const json: ClientSvcResponse = JSON.parse(res);
    const response: ClientSvcResponseContents = json[0];

    if (response.ErrorInfo) {
      throw response.ErrorInfo.ErrorMessage;
    }
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};

    if (options.newTitle) {
      requestBody.Title = options.newTitle;
    }

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

      if (options.enableVersioning === undefined) {
        requestBody.EnableVersioning = true;
      }
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

export default new SpoListSetCommand();