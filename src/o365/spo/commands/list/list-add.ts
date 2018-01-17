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
import { ContextInfo } from '../../spo';
import Utils from '../../../../Utils';
import { ListInstance } from "./ListInstance";
import Auth from '../../../../Auth';
import { ListTemplateType } from './ListTemplateType';

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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    //telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.title = (!(!args.options.title)).toString();
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
      .then((accessToken: string): Promise<ContextInfo> => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
      })
      .then((res: ContextInfo): Promise<ListInstance> => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
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
          cmd.log(args.options.schemaXml);
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
        description: 'Boolean value specifying whether the list can be deleted. Valid values are true|false',
        autocomplete: ['true', 'false']
      },   
      {
        option: '--allowEveryoneViewItems [allowEveryoneViewItems]',
        description: 'Boolean value specifying whether everyone can view documents in the document library or attachments to items in the list. Valid values are true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--allowMultiResponses [allowMultiResponses]',
        description: 'Boolean value specifying whether users are allowed to give multiple responses to the survey. Valid values are true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--contentTypesEnabled [contentTypesEnabled]',
        description: 'Boolean value specifying whether content types are enabled for the list. Valid values are true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--crawlNonDefaultViews [crawlNonDefaultViews]',
        description: 'Boolean value specifying whether to crawl non default views. Valid values are true|false',
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
        description: 'Property for assigning or retrieving grid editing on the list. Valid values are true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--draftVersionVisibility [draftVersionVisibility]',
        description: 'Value that specifies the minimum permission required to view minor versions and drafts within the list'
      },
      {
        option: '--emailAlias [emailAlias]',
        description: 'If e-mail notification is enabled, gets or sets the e-mail address to use to notify to the owner of an item when an assignment has changed or the item has been updated'
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
        description: 'Boolean value that specifies whether versioning is enabled for the document library'
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
        description: ''
      },
      {
        option: '--indexedRootFolderPropertyKeys [indexedRootFolderPropertyKeys]',
        description: ''
      },
      {
        option: '--irmEnabled [irmEnabled]',
        description: ''
      },
      {
        option: '--irmExpire [irmExpire]',
        description: ''
      },
      {
        option: '--irmReject [irmReject]',
        description: ''
      },
      {
        option: '--isApplicationList [isApplicationList]',
        description: ''
      },
      {
        option: '--listExperienceOptions [listExperienceOptions]',
        description: ''
      },
      {
        option: '--majorVersionLimit [majorVersionLimit]',
        description: ''
      },
      {
        option: '--majorWithMinorVersionsLimit [majorWithMinorVersionsLimit]',
        description: ''
      },
      {
        option: '--multipleDataList [multipleDataList]',
        description: ''
      },
      {
        option: '--navigateForFormsPages [navigateForFormsPages]',
        description: ''
      },
      {
        option: '--needUpdateSiteClientTag [needUpdateSiteClientTag]',
        description: ''
      },
      {
        option: '--noCrawl [noCrawl]',
        description: ''
      },
      {
        option: '--onQuickLaunch [onQuickLaunch]',
        description: ''
      },
      {
        option: '--ordered [ordered]',
        description: ''
      },
      {
        option: '-parserDisabled [parserDisabled]',
        description: ''
      },
      {
        option: '--readOnlyUI [readOnlyUI]',
        description: ''
      },
      {
        option: '--readSecurity [readSecurity]',
        description: ''
      },
      {
        option: '--requestAccessEnabled [requestAccessEnabled]',
        description: ''
      },
      {
        option: '--restrictUserUpdates [restrictUserUpdates]',
        description: ''
      },
      {
        option: '--sendToLocationName [sendToLocationName]',
        description: ''
      },
      {
        option: '--sendToLocationUrl [sendToLocationUrl]',
        description: ''
      },
      {
        option: '--showUser [showUser]',
        description: ''
      },
      {
        option: '--smsAlertTemplate [smsAlertTemplate]',
        description: ''
      },
      {
        option: '--useFormsForDisplay [useFormsForDisplay]',
        description: ''
      },
      {
        option: '--validationFormula [validationFormula]',
        description: ''
      },
      {
        option: '--validationMessage [validationMessage]',
        description: ''
      },
      {
        option: '--writeSecurity [writeSecurity]',
        description: ''
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
      
      if (args.options.baseTemplate) {
        const template: ListTemplateType = ListTemplateType[(args.options.baseTemplate.trim() as keyof typeof ListTemplateType)];

        if (!template) {
          return `BaseTemplate option '${args.options.baseTemplate}' is not recognized as valid choice. Please note it is case sensitive`;
        }
      }

      // if (args.options.baseTemplate) {
      //   if (typeof args.options.baseTemplate !== 'number'){
      //     return `${args.options.baseTemplate} is not a number`;
      //   }
      // }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (args.options.templateFeatureId) {
        if (!Utils.isValidGuid(args.options.templateFeatureId)) {
          return `${args.options.templateFeatureId} is not a valid GUID`;
        }
      }

      if (args.options.allowDeletion) {
        if (args.options.allowDeletion.toLowerCase() !== 'true' &&
          args.options.allowDeletion.toLowerCase() !== 'false') {
          return `allowDeletion value ${args.options.allowDeletion} is not a valid boolean value. Allowed values are true|false`;
        }
      }

      if (args.options.allowEveryoneViewItems) {
        if (args.options.allowEveryoneViewItems.toLowerCase() !== 'true' &&
          args.options.allowEveryoneViewItems.toLowerCase() !== 'false') {
          return `allowEveryoneViewItems value ${args.options.allowEveryoneViewItems} is not a valid boolean value. Allowed values are true|false`;
        }
      }

      if (args.options.allowMultiResponses) {
        if (args.options.allowMultiResponses.toLowerCase() !== 'true' &&
          args.options.allowMultiResponses.toLowerCase() !== 'false') {
          return `allowMultiResponses value ${args.options.allowMultiResponses} is not a valid boolean value. Allowed values are true|false`;
        }
      }

      if (args.options.contentTypesEnabled) {
        if (args.options.contentTypesEnabled.toLowerCase() !== 'true' &&
          args.options.contentTypesEnabled.toLowerCase() !== 'false') {
          return `contentTypesEnabled value ${args.options.contentTypesEnabled} is not a valid boolean value. Allowed values are true|false`;
        }
      }

      if (args.options.crawlNonDefaultViews) {
        if (args.options.crawlNonDefaultViews.toLowerCase() !== 'true' &&
          args.options.crawlNonDefaultViews.toLowerCase() !== 'false') {
          return `crawlNonDefaultViews value ${args.options.crawlNonDefaultViews} is not a valid boolean value. Allowed values are true|false`;
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
          return `direction value ${args.options.direction} is not a valid boolean value. Allowed values are NONE|LTR|RTL`;
        }
      }

      if (args.options.disableGridEditing) {
        if (args.options.disableGridEditing.toLowerCase() !== 'true' &&
          args.options.disableGridEditing.toLowerCase() !== 'false') {
          return `disableGridEditing value ${args.options.disableGridEditing} is not a valid boolean value. Allowed values are true|false`;
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
  
    Add a list with title ${chalk.grey('Announcements')}, baseTemplate ${chalk.grey('107')}
    in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.LIST_ADD} --title 'DemoList' --baseTemplate 107 --webUrl https://contoso.sharepoint.com/sites/project-x

    Add a list with title ${chalk.grey('Announcements')}, baseTemplate ${chalk.grey('107')}
    in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')} 
    with schemaXml ${chalk.grey('<List DocTemplateUrl="" DefaultViewUrl="" MobileDefaultViewUrl="" ID="{92FF93AB-920E-4D33-AE42-58B5E245BEFF}" Title="Announcements" Description="" ImageUrl="/_layouts/15/images/itann.png?rev=44" Name="{92FF93AB-920E-4D33-AE42-58B5E245BEFF}" BaseType="0" FeatureId="{00BFEA71-D1CE-42DE-9C63-A44004CE0104}" ServerTemplate="104" Created="20161221 20:02:12" Modified="20180110 19:35:15" LastDeleted="20161221 20:02:12" Version="0" Direction="none" ThumbnailSize="0" WebImageWidth="0" WebImageHeight="0" Flags="536875008" ItemCount="1" AnonymousPermMask="0" RootFolder="/sites/project-x/Lists/Announcements"      ReadSecurity="1" WriteSecurity="1" Author="3" EventSinkAssembly="" EventSinkClass="" EventSinkData="" EmailAlias="" WebFullUrl="/sites/project-x" WebId="7694137e-7038-4831-a1bd-218b28fe5d34" SendToLocation="" ScopeId="92facaf9-8d7a-40eb-9e69-362c91513cbd" MajorVersionLimit="0" MajorWithMinorVersionsLimit="0" WorkFlowId="00000000-0000-0000-0000-000000000000" HasUniqueScopes="False" NoThrottleListOperations="False" HasRelatedLists="False" Followable="False" Acl="" Flags2="0" RootFolderId="d4d67cc1-ad6e-4293-b039-ea49263d195f" ComplianceTag="" ComplianceFlags="0" UserModified="20161221 20:03:00" ListSchemaVersion="3" AclVersion="" AllowDeletion="True" AllowMultiResponses="False" EnableAttachments="True" EnableModeration="False" EnableVersioning="False" HasExternalDataSource="False" Hidden="False" MultipleDataList="False" Ordered="False" ShowUser="True" EnablePeopleSelector="False" EnableResourceSelector="False" EnableMinorVersion="False" RequireCheckout="False" ThrottleListOperations="False" ExcludeFromOfflineClient="False" CanOpenFileAsync="True" EnableFolderCreation="False" IrmEnabled="False" IrmSyncable="False" IsApplicationList="False" PreserveEmptyValues="False" StrictTypeCoercion="False" EnforceDataValidation="False" MaxItemsPerThrottledOperation="5000"></List>')}
      ${chalk.grey(config.delimiter)} ${commands.LIST_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --title Documents
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
      requestBody.SchemaXml = options.schemaXml;
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

    return requestBody;
  }
}

module.exports = new ListAddCommand();