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
import { Term } from './Term';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  termGroupId?: string;
  termGroupName?: string;
  termSetId?: string;
  termSetName?: string;
}

class SpoTermGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TERM_GET;
  }

  public get description(): string {
    return 'Gets information about the specified taxonomy term';
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
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        termGroupId: typeof args.options.termGroupId !== 'undefined',
        termGroupName: typeof args.options.termGroupName !== 'undefined',
        termSetId: typeof args.options.termSetId !== 'undefined',
        termSetName: typeof args.options.termSetName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '--termGroupId [termGroupId]'
      },
      {
        option: '--termGroupName [termGroupName]'
      },
      {
        option: '--termSetId [termSetId]'
      },
      {
        option: '--termSetName [termSetName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!args.options.id && !args.options.name) {
          return 'Specify either id or name';
        }

        if (args.options.id && args.options.name) {
          return 'Specify either id or name but not both';
        }

        if (args.options.id) {
          if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
          }
        }

        if (args.options.name) {
          if (!args.options.termGroupId && !args.options.termGroupName) {
            return 'Specify termGroupId or termGroupName';
          }

          if (!args.options.termSetId && !args.options.termSetName) {
            return 'Specify termSetId or termSetName';
          }
        }

        if (args.options.termGroupId && args.options.termGroupName) {
          return 'Specify termGroupId or termGroupName but not both';
        }

        if (args.options.termGroupId) {
          if (!validation.isValidGuid(args.options.termGroupId)) {
            return `${args.options.termGroupId} is not a valid GUID`;
          }
        }

        if (args.options.termSetId && args.options.termSetName) {
          return 'Specify termSetId or termSetName but not both';
        }

        if (args.options.termSetId) {
          if (!validation.isValidGuid(args.options.termSetId)) {
            return `${args.options.termSetId} is not a valid GUID`;
          }
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    spo
      .getSpoAdminUrl(logger, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;
        return spo.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          logger.logToStderr(`Retrieving taxonomy term...`);
        }

        let data: string = '';

        if (args.options.id) {
          data = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="6" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="7" ParentId="6" Name="GetDefaultSiteCollectionTermStore" /><Method Id="13" ParentId="7" Name="GetTerm"><Parameters><Parameter Type="Guid">{${args.options.id}}</Parameter></Parameters></Method></ObjectPaths></Request>`;
        }
        else {
          const termGroupQuery: string = args.options.termGroupId ? `<Method Id="98" ParentId="96" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termGroupId}}</Parameter></Parameters></Method>` : `<Method Id="98" ParentId="96" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termGroupName)}</Parameter></Parameters></Method>`;
          const termSetQuery: string = args.options.termSetId ? `<Method Id="103" ParentId="101" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termSetId}}</Parameter></Parameters></Method>` : `<Method Id="103" ParentId="101" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termSetName)}</Parameter></Parameters></Method>`;
          data = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="91" ObjectPathId="90" /><ObjectIdentityQuery Id="92" ObjectPathId="90" /><ObjectPath Id="94" ObjectPathId="93" /><ObjectIdentityQuery Id="95" ObjectPathId="93" /><ObjectPath Id="97" ObjectPathId="96" /><ObjectPath Id="99" ObjectPathId="98" /><ObjectIdentityQuery Id="100" ObjectPathId="98" /><ObjectPath Id="102" ObjectPathId="101" /><ObjectPath Id="104" ObjectPathId="103" /><ObjectIdentityQuery Id="105" ObjectPathId="103" /><ObjectPath Id="107" ObjectPathId="106" /><ObjectPath Id="109" ObjectPathId="108" /><ObjectIdentityQuery Id="110" ObjectPathId="108" /><Query Id="111" ObjectPathId="108"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="90" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="93" ParentId="90" Name="GetDefaultSiteCollectionTermStore" /><Property Id="96" ParentId="93" Name="Groups" />${termGroupQuery}<Property Id="101" ParentId="98" Name="TermSets" />${termSetQuery}<Property Id="106" ParentId="103" Name="Terms" /><Method Id="108" ParentId="106" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.name)}</Parameter></Parameters></Method></ObjectPaths></Request>`;
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          data: data
        };

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
          return;
        }

        const term: Term | null = json[json.length - 1];
        if (!term) {
          cb();
          return;
        }

        delete term._ObjectIdentity_;
        delete term._ObjectType_;
        term.CreatedDate = new Date(Number(term.CreatedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
        term.Id = term.Id.replace('/Guid(', '').replace(')/', '');
        term.LastModifiedDate = new Date(Number(term.LastModifiedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
        logger.log(term);
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }
}

module.exports = new SpoTermGetCommand();