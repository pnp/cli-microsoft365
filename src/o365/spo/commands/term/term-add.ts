import auth from '../../SpoAuth';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';
import * as request from 'request-promise-native';
const uuidv4 = require('uuid/v4');
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandError,
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { Term } from './Term';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  customProperties?: string;
  description?: string;
  id?: string;
  localCustomProperties?: string;
  name: string;
  termGroupId?: string;
  termGroupName?: string;
  termSetId?: string;
  termSetName?: string;
}

class SpoTermAddCommand extends SpoCommand {
  public get name(): string {
    return commands.TERM_ADD;
  }

  public get description(): string {
    return 'Adds taxonomy term';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.customProperties = typeof args.options.customProperties !== 'undefined';
    telemetryProps.description = typeof args.options.description !== 'undefined';
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.localCustomProperties = typeof args.options.localCustomProperties !== 'undefined';
    telemetryProps.termGroupId = typeof args.options.termGroupId !== 'undefined';
    telemetryProps.termGroupName = typeof args.options.termGroupName !== 'undefined';
    telemetryProps.termSetId = typeof args.options.termSetId !== 'undefined';
    telemetryProps.termSetName = typeof args.options.termSetName !== 'undefined';
    return telemetryProps;
  }

  protected requiresTenantAdmin(): boolean {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let term: Term;
    let formDigest: string;

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        return this.getRequestDigest(cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        formDigest = res.FormDigestValue;

        if (this.verbose) {
          cmd.log(`Adding taxonomy term...`);
        }

        const termGroupQuery: string = args.options.termGroupId ? `<Method Id="11" ParentId="9" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termGroupId}}</Parameter></Parameters></Method>` : `<Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.termGroupName)}</Parameter></Parameters></Method>`;
        const termSetQuery: string = args.options.termSetId ? `<Method Id="16" ParentId="14" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termSetId}}</Parameter></Parameters></Method>` : `<Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.termSetName)}</Parameter></Parameters></Method>`;
        const termId: string = args.options.id || uuidv4();
        const body: string = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" />${termGroupQuery}<Property Id="14" ParentId="11" Name="TermSets" />${termSetQuery}<Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.name)}</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{${termId}}</Parameter></Parameters></Method></ObjectPaths></Request>`;

        const requestOptions: any = {
          url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'X-RequestDigest': res.FormDigestValue
          }),
          body: body
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: string): request.RequestPromise | Promise<void> => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(JSON.parse(res)));
          cmd.log('');
        }

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          return Promise.reject(response.ErrorInfo.ErrorMessage);
        }

        term = json[json.length - 1];

        if (!args.options.description &&
          !args.options.customProperties &&
          !args.options.localCustomProperties) {
          return Promise.resolve();
        }

        if (this.verbose) {
          cmd.log(`Setting term properties...`);
        }

        const properties: string[] = [];
        let i: number = 127;
        if (args.options.description) {
          properties.push(`<Method Name="SetDescription" Id="${i++}" ObjectPathId="117"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.description)}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method>`);
          term.Description = args.options.description;
        }
        if (args.options.customProperties) {
          const customProperties: any = JSON.parse(args.options.customProperties);
          Object.keys(customProperties).forEach(k => {
            properties.push(`<Method Name="SetCustomProperty" Id="${i++}" ObjectPathId="117"><Parameters><Parameter Type="String">${Utils.escapeXml(k)}</Parameter><Parameter Type="String">${Utils.escapeXml(customProperties[k])}</Parameter></Parameters></Method>`);
          });
          term.CustomProperties = customProperties;
        }
        if (args.options.localCustomProperties) {
          const localCustomProperties: any = JSON.parse(args.options.localCustomProperties);
          Object.keys(localCustomProperties).forEach(k => {
            properties.push(`<Method Name="SetLocalCustomProperty" Id="${i++}" ObjectPathId="117"><Parameters><Parameter Type="String">${Utils.escapeXml(k)}</Parameter><Parameter Type="String">${Utils.escapeXml(localCustomProperties[k])}</Parameter></Parameters></Method>`);
          });
          term.LocalCustomProperties = localCustomProperties;
        }

        let termStoreObjectIdentity: string = '';
        // get term store object identity
        for (let i: number = 0; i < json.length; i++) {
          if (json[i] !== 8) {
            continue;
          }

          termStoreObjectIdentity = json[i + 1]._ObjectIdentity_;
          break;
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'X-RequestDigest': formDigest
          }),
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${properties.join('')}<Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="${term._ObjectIdentity_}" /><Identity Id="109" Name="${termStoreObjectIdentity}" /></ObjectPaths></Request>`
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res?: string): void => {
        if (res) {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            cb(new CommandError(response.ErrorInfo.ErrorMessage));
            return;
          }
        }

        delete term._ObjectIdentity_;
        delete term._ObjectType_;
        term.CreatedDate = new Date(Number(term.CreatedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
        term.Id = term.Id.replace('/Guid(', '').replace(')/', '');
        term.LastModifiedDate = new Date(Number(term.LastModifiedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
        cmd.log(term);
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'Name of the term to add'
      },
      {
        option: '--termSetId [termSetId]',
        description: 'ID of the term set in which to create the term. Specify termSetId or termSetName but not both'
      },
      {
        option: '--termSetName [termSetName]',
        description: 'Name of the term set in which to create the term. Specify termSetId or termSetName but not both'
      },
      {
        option: '--termGroupId [termGroupId]',
        description: 'ID of the term group to which the term set belongs. Specify termGroupId or termGroupName but not both'
      },
      {
        option: '--termGroupName [termGroupName]',
        description: 'Name of the term group to which the term set belongs. Specify termGroupId or termGroupName but not both'
      },
      {
        option: '-i, --id [id]',
        description: 'ID of the term to add'
      },
      {
        option: '-d, --description [description]',
        description: 'Description of the term to add'
      },
      {
        option: '--customProperties [customProperties]',
        description: 'JSON string with key-value pairs representing custom properties to set on the term'
      },
      {
        option: '--localCustomProperties [localCustomProperties]',
        description: 'JSON string with key-value pairs representing local custom properties to set on the term'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.name) {
        return 'Required option name missing';
      }

      if (args.options.id) {
        if (!Utils.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }
      }

      if (!args.options.termGroupId && !args.options.termGroupName) {
        return 'Specify termGroupId or termGroupName';
      }

      if (args.options.termGroupId && args.options.termGroupName) {
        return 'Specify termGroupId or termGroupName but not both';
      }

      if (args.options.termGroupId) {
        if (!Utils.isValidGuid(args.options.termGroupId)) {
          return `${args.options.termGroupId} is not a valid GUID`;
        }
      }

      if (!args.options.termSetId && !args.options.termSetName) {
        return 'Specify termSetId or termSetName';
      }

      if (args.options.termSetId && args.options.termSetName) {
        return 'Specify termSetId or termSetName but not both';
      }

      if (args.options.termSetId) {
        if (!Utils.isValidGuid(args.options.termSetId)) {
          return `${args.options.termSetId} is not a valid GUID`;
        }
      }

      if (args.options.customProperties) {
        try {
          JSON.parse(args.options.customProperties);
        }
        catch (e) {
          return `An error has occurred while parsing customProperties: ${e}`;
        }
      }

      if (args.options.localCustomProperties) {
        try {
          JSON.parse(args.options.localCustomProperties);
        }
        catch (e) {
          return `An error has occurred while parsing localCustomProperties: ${e}`;
        }
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.TERM_ADD).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to SharePoint Online tenant admin site,
    using the ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    To add a taxonomy term, you have to first log in
    to a tenant admin site using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso-admin.sharepoint.com`)}.

  Examples:
  
    Add taxonomy term with the specified name to the term group and term set
    specified by their names
      ${chalk.grey(config.delimiter)} ${commands.TERM_ADD} --name IT --termSetName Department --termGroupName People

    Add taxonomy term with the specified name to the term group and term set
    specified by their IDs
      ${chalk.grey(config.delimiter)} ${commands.TERM_ADD} --name IT --termSetId 8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f --termGroupId 5c928151-c140-4d48-aab9-54da901c7fef

    Add taxonomy term with the specified name and ID
      ${chalk.grey(config.delimiter)} ${commands.TERM_ADD} --name IT --id 5c928151-c140-4d48-aab9-54da901c7fef --termSetName Department --termGroupName People

    Add taxonomy term with custom properties
      ${chalk.grey(config.delimiter)} ${commands.TERM_ADD} --name IT --termSetName Department --termGroupName People --customProperties '{"Property": "Value"}'
`);
  }
}

module.exports = new SpoTermAddCommand();