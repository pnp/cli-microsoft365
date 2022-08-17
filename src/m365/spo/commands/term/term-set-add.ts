import { v4 } from 'uuid';
import { Logger } from '../../../../cli';
import {
  CommandError
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, formatting, spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { TermSet } from './TermSet';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  customProperties?: string;
  description?: string;
  id?: string;
  name: string;
  termGroupId?: string;
  termGroupName?: string;
}

class SpoTermSetAddCommand extends SpoCommand {
  public get name(): string {
    return commands.TERM_SET_ADD;
  }

  public get description(): string {
    return 'Adds taxonomy term set';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        customProperties: typeof args.options.customProperties !== 'undefined',
        description: typeof args.options.description !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        termGroupId: typeof args.options.termGroupId !== 'undefined',
        termGroupName: typeof args.options.termGroupName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '--termGroupId [termGroupId]'
      },
      {
        option: '--termGroupName [termGroupName]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '--customProperties [customProperties]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id) {
          if (!validation.isValidGuid(args.options.id)) {
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
          if (!validation.isValidGuid(args.options.termGroupId)) {
            return `${args.options.termGroupId} is not a valid GUID`;
          }
        }

        if (args.options.customProperties) {
          try {
            JSON.parse(args.options.customProperties);
          }
          catch (e) {
            return `Error when parsing customProperties JSON: ${e}`;
          }
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let formDigest: string = '';
    let termSet: TermSet;
    let spoAdminUrl: string = '';

    spo
      .getSpoAdminUrl(logger, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;
        return spo.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        formDigest = res.FormDigestValue;

        if (this.verbose) {
          logger.logToStderr(`Adding taxonomy term set...`);
        }

        const termGroupQuery: string = args.options.termGroupName ?
          `<Method Id="42" ParentId="40" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termGroupName)}</Parameter></Parameters></Method>` :
          `<Method Id="42" ParentId="40" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termGroupId}}</Parameter></Parameters></Method>`;
        const termSetId: string = args.options.id || v4();

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" />${termGroupQuery}<Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.name)}</Parameter><Parameter Type="Guid">{${termSetId}}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): Promise<string> => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          return Promise.reject(response.ErrorInfo.ErrorMessage);
        }

        termSet = json[json.length - 1];

        if (!args.options.description &&
          !args.options.customProperties) {
          return Promise.resolve(undefined as any);
        }

        let termStoreObjectIdentity: string = '';
        // get term store object identity
        for (let i: number = 0; i < json.length; i++) {
          if (json[i] !== 39) {
            continue;
          }

          termStoreObjectIdentity = json[i + 1]._ObjectIdentity_;
          break;
        }

        if (this.verbose) {
          logger.logToStderr(`Setting term set properties...`);
        }

        const properties: string[] = [];
        let i: number = 127;
        if (args.options.description) {
          properties.push(`<SetProperty Id="${i++}" ObjectPathId="117" Name="Description"><Parameter Type="String">${formatting.escapeXml(args.options.description)}</Parameter></SetProperty>`);
          termSet.Description = args.options.description;
        }
        if (args.options.customProperties) {
          const customProperties: any = JSON.parse(args.options.customProperties);
          Object.keys(customProperties).forEach(k => {
            properties.push(`<Method Name="SetCustomProperty" Id="${i++}" ObjectPathId="117"><Parameters><Parameter Type="String">${formatting.escapeXml(k)}</Parameter><Parameter Type="String">${formatting.escapeXml(customProperties[k])}</Parameter></Parameters></Method>`);
          });
          termSet.CustomProperties = customProperties;
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': formDigest
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${properties.join('')}<Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="${termSet._ObjectIdentity_}" /><Identity Id="109" Name="${termStoreObjectIdentity}" /></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res?: string): void => {
        if (res) {
          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            cb(new CommandError(response.ErrorInfo.ErrorMessage));
            return;
          }
        }

        delete termSet._ObjectIdentity_;
        delete termSet._ObjectType_;
        termSet.CreatedDate = new Date(Number(termSet.CreatedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
        termSet.Id = termSet.Id.replace('/Guid(', '').replace(')/', '');
        termSet.LastModifiedDate = new Date(Number(termSet.LastModifiedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
        logger.log(termSet);
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }
}

module.exports = new SpoTermSetAddCommand();