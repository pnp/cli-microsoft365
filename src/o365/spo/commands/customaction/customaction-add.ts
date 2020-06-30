import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { CustomAction } from './customaction';
import { BasePermissions, PermissionKind } from '../../base-permissions';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  name: string;
  title: string;
  location: string;
  group?: string;
  description?: string;
  sequence?: number;
  actionUrl?: string;
  imageUrl?: string;
  commandUIExtension?: string;
  registrationId?: string;
  registrationType?: string;
  rights?: string;
  scriptSrc?: string;
  scriptBlock?: string;
  scope?: string;
  clientSideComponentId?: string;
  clientSideComponentProperties?: string;
}

class SpoCustomActionAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.CUSTOMACTION_ADD}`;
  }

  public get description(): string {
    return 'Adds a user custom action for site or site collection';
  }

  /**
   * Maps the base PermissionsKind enum to string array so it can 
   * more easily be used in validation or descriptions.
   */
  protected get permissionsKindMap(): string[] {
    const result: string[] = [];

    for (let kind in PermissionKind) {
      if (typeof PermissionKind[kind] === 'number') {
        result.push(kind);
      }
    }
    return result;
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.location = args.options.location;
    telemetryProps.scope = args.options.scope || 'Web';
    telemetryProps.group = (!(!args.options.group)).toString();
    telemetryProps.description = (!(!args.options.description)).toString();
    telemetryProps.sequence = (!(!args.options.sequence)).toString();
    telemetryProps.actionUrl = (!(!args.options.actionUrl)).toString();
    telemetryProps.imageUrl = (!(!args.options.imageUrl)).toString();
    telemetryProps.commandUIExtension = (!(!args.options.commandUIExtension)).toString();
    telemetryProps.registrationId = args.options.registrationId;
    telemetryProps.registrationType = args.options.registrationType;
    telemetryProps.rights = args.options.rights;
    telemetryProps.scriptSrc = (!(!args.options.scriptSrc)).toString();
    telemetryProps.scriptBlock = (!(!args.options.scriptBlock)).toString();
    telemetryProps.clientSideComponentId = (!(!args.options.clientSideComponentId)).toString();
    telemetryProps.clientSideComponentProperties = (!(!args.options.clientSideComponentProperties)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (!args.options.scope) {
      args.options.scope = 'Web';
    }

    const requestBody: any = this.mapRequestBody(args.options);

    const requestOptions: any = {
      url: `${args.options.url}/_api/${args.options.scope}/UserCustomActions`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      body: requestBody,
      json: true
    };

    request
      .post<CustomAction>(requestOptions)
      .then((customAction: CustomAction): void => {
        if (this.verbose) {
          cmd.log({
            ClientSideComponentId: customAction.ClientSideComponentId,
            ClientSideComponentProperties: customAction.ClientSideComponentProperties,
            CommandUIExtension: customAction.CommandUIExtension,
            Description: customAction.Description,
            Group: customAction.Group,
            Id: customAction.Id,
            ImageUrl: customAction.ImageUrl,
            Location: customAction.Location,
            Name: customAction.Name,
            RegistrationId: customAction.RegistrationId,
            RegistrationType: customAction.RegistrationType,
            Rights: JSON.stringify(customAction.Rights),
            Scope: args.options.scope, // because it is more human readable
            ScriptBlock: customAction.ScriptBlock,
            ScriptSrc: customAction.ScriptSrc,
            Sequence: customAction.Sequence,
            Title: customAction.Title,
            Url: customAction.Url,
            VersionOfUserCustomAction: customAction.VersionOfUserCustomAction
          });
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'Url of the site or site collection to add the custom action'
      },
      {
        option: '-n, --name <name>',
        description: 'The name of the custom action'
      },
      {
        option: '-t, --title <title>',
        description: 'The title of the custom action'
      },
      {
        option: '-l, --location <location>',
        description: 'The actual location where this custom action need to be added like "CommandUI.Ribbon"'
      },
      {
        option: '-g, --group [group]',
        description: 'The group where this custom action needs to be added like "SiteActions"'
      },
      {
        option: '-d, --description [description]',
        description: 'The description of the custom action'
      },
      {
        option: '--sequence [sequence]',
        description: 'Sequence of this CustomAction being injected. Use when you have a specific sequence with which to have multiple CustomActions being added to the page'
      },
      {
        option: '--actionUrl [actionUrl]',
        description: 'The URL, URI or JavaScript function associated with the action. URL example ~site/_layouts/sampleurl.aspx or ~sitecollection/_layouts/sampleurl.aspx'
      },
      {
        option: '--imageUrl [imageUrl]',
        description: 'The URL of the image associated with the custom action'
      },
      {
        option: '-e, --commandUIExtension [commandUIExtension]',
        description: 'XML fragment that determines user interface properties of the custom action'
      },
      {
        option: '--registrationId [registrationId]',
        description: 'Specifies the identifier of the list or item content type that this action is associated with, or the file type or programmatic identifier'
      },
      {
        option: '--registrationType [registrationType]',
        description: 'Specifies the type of object associated with the custom action. Allowed values None|List|ContentType|ProgId|FileType. Default None',
        autocomplete: ['None', 'List', 'ContentType', 'ProgId', 'FileType']
      },
      {
        option: '--rights [rights]',
        description: `A case sensitive string array that contain the permissions needed for the custom action. Allowed values ${this.permissionsKindMap.join('|')}. Default ${this.permissionsKindMap[0]}`,
        autocomplete: this.permissionsKindMap
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the custom action. Allowed values Site|Web. Default Web',
        autocomplete: ['Site', 'Web']
      },
      {
        option: '--scriptBlock [scriptBlock]',
        description: 'Specifies a block of script to be executed. This attribute is only applicable when the Location attribute is set to ScriptLink'
      },
      {
        option: '--scriptSrc [scriptSrc]',
        description: 'Specifies a file that contains script to be executed. This attribute is only applicable when the Location attribute is set to ScriptLink'
      },
      {
        option: '-c, --clientSideComponentId [clientSideComponentId]',
        description: 'The Client Side Component Id (GUID) of the custom action'
      },
      {
        option: '-p, --clientSideComponentProperties [clientSideComponentProperties]',
        description: 'The Client Side Component Properties of the custom action. Specify values as a JSON string : "{Property1 : "Value1", Property2: "Value2"}"'
      }
    ];
    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.title) {
        return 'Missing required option title';
      }

      if (!args.options.name) {
        return 'Missing required option name';
      }

      if (!args.options.location) {
        return 'Missing required option location';
      }

      if (!args.options.url || SpoCommand.isValidSharePointUrl(args.options.url) !== true) {
        return 'Missing required option url';
      }

      if (args.options.registrationId && !args.options.registrationType) {
        return 'Option registrationId is specified, but registrationType is missing';
      }

      if (args.options.registrationType && !args.options.registrationId) {
        return 'Option registrationType is specified, but registrationId is missing';
      }

      let location: string = args.options.location.toLowerCase();
      const locationsRequireGroup: string[] = [
        'microsoft.sharepoint.standardmenu', 'microsoft.sharepoint.contenttypesettings',
        'microsoft.sharepoint.contenttypetemplatesettings', 'microsoft.sharepoint.create',
        'microsoft.sharepoint.groupspage', 'microsoft.sharepoint.listedit',
        'microsoft.sharepoint.listedit.documentlibrary', 'microsoft.sharepoint.peoplepage',
        'microsoft.sharepoint.sitesettings'
      ];

      if (locationsRequireGroup.indexOf(location) > -1 && !args.options.group) {
        return `The location specified requires the group option to be specified as well`;
      }

      if (location === 'scriptlink' &&
        !args.options.scriptSrc &&
        !args.options.scriptBlock
      ) {
        return 'Option scriptSrc or scriptBlock is required when the location is set to ScriptLink';
      }

      if ((args.options.scriptSrc || args.options.scriptBlock) && location !== 'scriptlink') {
        return 'Option scriptSrc or scriptBlock is specified, but the location option is different than ScriptLink. Please use --actionUrl, if the location should be different than ScriptLink';
      }

      if (args.options.scriptSrc && args.options.scriptBlock) {
        return 'Either option scriptSrc or scriptBlock can be specified, but not both';
      }

      if (args.options.sequence && (args.options.sequence < 0 || args.options.sequence > 65536)) {
        return 'Invalid option sequence. Expected value in range from 0 to 65536';
      }

      if (args.options.clientSideComponentId && Utils.isValidGuid(args.options.clientSideComponentId) === false) {
        return `ClientSideComponentId ${args.options.clientSideComponentId} is not a valid GUID`;
      }

      if (args.options.clientSideComponentProperties && !args.options.clientSideComponentId) {
        return `Option clientSideComponentProperties is specified, but the clientSideComponentId option is missing`;
      }

      if (args.options.scope &&
        args.options.scope !== 'Site' &&
        args.options.scope !== 'Web'
      ) {
        return `${args.options.scope} is not a valid custom action scope. Allowed values are Site|Web`;
      }

      if (args.options.rights) {
        const rights = args.options.rights.split(',');

        for (let item of rights) {
          const kind: PermissionKind = PermissionKind[(item.trim() as keyof typeof PermissionKind)];

          if (!kind) {
            return `Rights option '${item}' is not recognized as valid PermissionKind choice. Please note it is case sensitive`;
          }
        }
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.CUSTOMACTION_ADD).helpInformation());
    log(
      `  Remarks:
          
    Running this command from the Windows Command Shell (cmd.exe) or PowerShell for Windows OS XP,
    7, 8, 8.1 without bash installed might require additional formatting for command options that have
    JSON, XML or JavaScript values, because the command shell treat quotes differently. For example,
    this is how ApplicationCustomizer user custom action can be created from the Windows cmd.exe:

      o365 ${commands.CUSTOMACTION_ADD} -u https://contoso.sharepoint.com/sites/test -t "YourAppCustomizer" -n "YourName" -l "ClientSideExtension.ApplicationCustomizer" -c b41916e7-e69d-467f-b37f-ff8ecf8f99f2 -p '{\"testMessage\":\"Test message\"}'
    
    Note, how the clientSideComponentProperties option (-p) has escaped double quotes
    ${chalk.grey(`'{\"testMessage\":\"Test message\"}'`)} compared to execution from bash:
    ${chalk.grey(`'{"testMessage":"Test message"}'`)}.

    The ${chalk.grey(`--rights`)} option accepts case-sensitive values.

  Examples:
    
    Adds tenant-wide SharePoint Framework Application Customizer extension in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.CUSTOMACTION_ADD} -u https://contoso.sharepoint.com/sites/test -t "YourAppCustomizer" -n "YourName" -l "ClientSideExtension.ApplicationCustomizer" -c b41916e7-e69d-467f-b37f-ff8ecf8f99f2 -p '{"testMessage":"Test message"}'
    
    Adds tenant-wide SharePoint Framework ${chalk.blue('modern list view')} Command Set extension in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.CUSTOMACTION_ADD} -u https://contoso.sharepoint.com/sites/test -t "YourCommandSet" -n "YourName" -l "ClientSideExtension.ListViewCommandSet" -c db3e6e35-363c-42b9-a254-ca661e437848 -p '{"sampleTextOne":"One item is selected in the list.", "sampleTextTwo":"This command is always visible."}' --registrationId 100 --registrationType List
    
    Creates url custom action in the SiteActions menu in site ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.CUSTOMACTION_ADD} -u https://contoso.sharepoint.com/sites/test -t "YourTitle" -n "YourName" -l "Microsoft.SharePoint.StandardMenu" -g "SiteActions" --actionUrl "~site/SitePages/Home.aspx" --sequence 100
    
    Creates custom action in ${chalk.blue('classic')} Document Library edit context menu in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.CUSTOMACTION_ADD} -u https://contoso.sharepoint.com/sites/test -t "YourTitle" -n "YourName" -l "EditControlBlock" --actionUrl "javascript:(function(){ return console.log('office365-cli rocks!'); })();" --registrationId 101 --registrationType List
    
    Creates ScriptLink custom action with script source in ${chalk.blue('classic pages')} in
    site collection ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.CUSTOMACTION_ADD} -u https://contoso.sharepoint.com/sites/test -t "YourTitle" -n "YourName" -l "ScriptLink" --scriptSrc "~sitecollection/SiteAssets/YourScript.js" --sequence 101 -s Site
    
    Creates ScriptLink custom action with script block in ${chalk.blue('classic pages')} in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.CUSTOMACTION_ADD} -u https://contoso.sharepoint.com/sites/test -t "YourTitle" -n "YourName" -l "ScriptLink" --scriptBlock "(function(){ return console.log('Hello office365-cli!'); })();" --sequence 102
    
    Creates ${chalk.blue('classic List View')} custom action located in the Ribbon in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.CUSTOMACTION_ADD} -u https://contoso.sharepoint.com/sites/test -t "YourTitle" -n "YourName" -l "CommandUI.Ribbon" --commandUIExtension '<CommandUIExtension><CommandUIDefinitions><CommandUIDefinition Location="Ribbon.List.Share.Controls._children"><Button Id="Ribbon.List.Share.GetItemsCountButton" Alt="Get list items count" Sequence="11" Command="Invoke_GetItemsCountButtonRequest" LabelText="Get Items Count" TemplateAlias="o1" Image32by32="_layouts/15/images/placeholder32x32.png" Image16by16="_layouts/15/images/placeholder16x16.png" /></CommandUIDefinition></CommandUIDefinitions><CommandUIHandlers><CommandUIHandler Command="Invoke_GetItemsCountButtonRequest" CommandAction="javascript: alert(ctx.TotalListItems);" EnabledScript="javascript: function checkEnable() { return (true);} checkEnable();"/></CommandUIHandlers></CommandUIExtension>'
    
    Creates custom action with delegated rights in the SiteActions menu in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${commands.CUSTOMACTION_ADD} -u https://contoso.sharepoint.com/sites/test -t "YourTitle" -n "YourName" -l "Microsoft.SharePoint.StandardMenu" -g "SiteActions" --actionUrl "~site/SitePages/Home.aspx" --rights "AddListItems,DeleteListItems,ManageLists"
  
  More information:

    UserCustomAction REST API resources:
      https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomAction
      
    UserCustomAction Locations and Group IDs:
      https://msdn.microsoft.com/en-us/library/office/bb802730.aspx

    UserCustomAction Element:
      https://msdn.microsoft.com/en-us/library/office/ms460194.aspx

    UserCustomAction Rights:
      https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.spbasepermissions.aspx

      `);
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {
      Title: options.title,
      Name: options.name,
      Location: options.location
    };

    if (options.group) {
      requestBody.Group = options.group;
    }

    if (options.description) {
      requestBody.Description = options.description;
    }

    if (options.sequence) {
      requestBody.Sequence = options.sequence;
    }

    if (options.registrationType) {
      requestBody.RegistrationType = this.getRegistrationType(options.registrationType);
    }

    if (options.registrationId) {
      requestBody.RegistrationId = options.registrationId.toString();
    }

    if (options.actionUrl) {
      requestBody.Url = options.actionUrl;
    }

    if (options.imageUrl) {
      requestBody.ImageUrl = options.imageUrl;
    }

    if (options.clientSideComponentId) {
      requestBody.ClientSideComponentId = options.clientSideComponentId;
    }

    if (options.clientSideComponentProperties) {
      requestBody.ClientSideComponentProperties = options.clientSideComponentProperties;
    }

    if (options.scriptBlock) {
      requestBody.ScriptBlock = options.scriptBlock;
    }

    if (options.scriptSrc) {
      requestBody.ScriptSrc = options.scriptSrc;
    }

    if (options.commandUIExtension) {
      requestBody.CommandUIExtension = `${options.commandUIExtension}`;
    }

    if (options.rights) {
      const permissions: BasePermissions = new BasePermissions();
      const rights = options.rights.split(',');

      for (let item of rights) {
        const kind: PermissionKind = PermissionKind[(item.trim() as keyof typeof PermissionKind)];

        permissions.set(kind);
      }
      requestBody.Rights = {
        High: permissions.high.toString(),
        Low: permissions.low.toString()
      }
    }

    return requestBody;
  }

  private getRegistrationType(registrationType: string): number {
    switch (registrationType.toLowerCase()) {
      case 'list':
        return 1;
      case 'contenttype':
        return 2;
      case 'progid':
        return 3;
      case 'filetype':
        return 4;
    }
    return 0; // None
  }
}

module.exports = new SpoCustomActionAddCommand();