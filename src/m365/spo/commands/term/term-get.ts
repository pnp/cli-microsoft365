import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { Term } from './Term';
import { TermCollection } from './TermCollection';
import * as os from 'os';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl?: string;
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
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        webUrl: typeof args.options.webUrl !== 'undefined',
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
        option: '-u, --webUrl [webUrl]'
      },
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
        if (args.options.webUrl) {
          const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
          if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
          }
        }

        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.termGroupId && !validation.isValidGuid(args.options.termGroupId)) {
          return `${args.options.termGroupId} is not a valid GUID`;
        }

        if (args.options.termSetId && !validation.isValidGuid(args.options.termSetId)) {
          return `${args.options.termSetId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['id', 'name']
      },
      {
        options: ['termGroupId', 'termGroupName'],
        runsWhen: (args: CommandArgs) => {
          return args.options.name !== undefined;
        }
      },
      {
        options: ['termSetId', 'termSetName'],
        runsWhen: (args: CommandArgs) => {
          return args.options.name !== undefined;
        }
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoWebUrl: string = args.options.webUrl ? args.options.webUrl : await spo.getSpoAdminUrl(logger, this.debug);
      const res: ContextInfo = await spo.getRequestDigest(spoWebUrl);
      if (this.verbose) {
        logger.logToStderr(`Retrieving taxonomy term...`);
      }

      let data = '';

      if (args.options.id) {
        data = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="6" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="7" ParentId="6" Name="GetDefaultSiteCollectionTermStore" /><Method Id="13" ParentId="7" Name="GetTerm"><Parameters><Parameter Type="Guid">{${args.options.id}}</Parameter></Parameters></Method></ObjectPaths></Request>`;
      }
      else {
        data = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectIdentityQuery Id="3" ObjectPathId="1" /><ObjectPath Id="5" ObjectPathId="4" /><ObjectIdentityQuery Id="6" ObjectPathId="4" /><ObjectPath Id="8" ObjectPathId="7" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><ObjectPath Id="13" ObjectPathId="12" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectIdentityQuery Id="16" ObjectPathId="14" /><ObjectPath Id="18" ObjectPathId="17" /><SetProperty Id="19" ObjectPathId="17" Name="TrimUnavailable"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="20" ObjectPathId="17" Name="TermLabel"><Parameter Type="String">${formatting.escapeXml(args.options.name)}</Parameter></SetProperty><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="1" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="4" ParentId="1" Name="GetDefaultSiteCollectionTermStore" /><Property Id="7" ParentId="4" Name="Groups" /><Method Id="9" ParentId="7" Name="${args.options.termGroupName === undefined ? "GetById" : "GetByName"}"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termGroupName) || args.options.termGroupId}</Parameter></Parameters></Method><Property Id="12" ParentId="9" Name="TermSets" /><Method Id="14" ParentId="12" Name="${args.options.termSetName === undefined ? "GetById" : "GetByName"}"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termSetName) || args.options.termSetId}</Parameter></Parameters></Method><Constructor Id="17" TypeId="{61a1d689-2744-4ea3-a88b-c95bee9803aa}" /><Method Id="21" ParentId="14" Name="GetTerms"><Parameters><Parameter ObjectPathId="17" /></Parameters></Method></ObjectPaths></Request>`;
      }

      let term: Term;
      const csomResponse = await this.executeCsomCall(data, spoWebUrl, res);

      if (csomResponse === null) {
        throw `Term with id '${args.options.id}' could not be found.`;
      }
      else if (csomResponse._Child_Items_ !== undefined) {
        const terms: TermCollection = csomResponse;
        if (terms._Child_Items_.length === 0) {
          throw `Term with name '${args.options.name}' could not be found.`;
        }

        if (terms._Child_Items_.length > 1) {
          const disambiguationText = terms._Child_Items_.map(c => {
            return `- ${this.getTermId(c.Id)} - ${c.PathOfTerm}`;
          }).join(os.EOL);

          throw new Error(`Multiple terms with the specific term name found. Please disambiguate:${os.EOL}${disambiguationText}`);
        }

        term = terms._Child_Items_[0];
      }
      else {
        term = csomResponse;
      }

      delete term._ObjectIdentity_;
      delete term._ObjectType_;

      term.CreatedDate = this.parseTermDateToIsoString(term.CreatedDate);
      term.Id = this.getTermId(term.Id);
      term.LastModifiedDate = this.parseTermDateToIsoString(term.LastModifiedDate);

      logger.log(term);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private async executeCsomCall(data: string, spoWebUrl: string, res: ContextInfo): Promise<any> {
    const requestOptions: CliRequestOptions = {
      url: `${spoWebUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': res.FormDigestValue
      },
      data: data
    };

    const processQuery: string = await request.post(requestOptions);
    const json: ClientSvcResponse = JSON.parse(processQuery);
    const response: ClientSvcResponseContents = json[0];

    if (response.ErrorInfo) {
      throw response.ErrorInfo.ErrorMessage;
    }

    const responseObject = json[json.length - 1];
    return responseObject;
  }

  private getTermId(termId: string): string {
    return termId.replace('/Guid(', '').replace(')/', '');
  }

  private parseTermDateToIsoString(dateAsString: string): string {
    return new Date(Number(dateAsString.replace('/Date(', '').replace(')/', ''))).toISOString();
  }
}

module.exports = new SpoTermGetCommand();