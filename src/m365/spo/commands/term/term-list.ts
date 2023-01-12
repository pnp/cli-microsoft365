import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { Term } from './Term';
import { TermCollection } from './TermCollection';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  termGroupId?: string;
  termGroupName?: string;
  termSetId?: string;
  termSetName?: string;
  includeChildTerms?: boolean;
}

class SpoTermListCommand extends SpoCommand {
  public get name(): string {
    return commands.TERM_LIST;
  }

  public get description(): string {
    return 'Lists taxonomy terms from the given term set';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Name', 'ParentTermId'];
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
        termGroupId: typeof args.options.termGroupId !== 'undefined',
        termGroupName: typeof args.options.termGroupName !== 'undefined',
        termSetId: typeof args.options.termSetId !== 'undefined',
        termSetName: typeof args.options.termSetName !== 'undefined',
        includeChildTerms: !!args.options.includeChildTerms
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
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
      },
      {
        option: '--includeChildTerms'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
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
      { options: ['termGroupId', 'termGroupName'] },
      { options: ['termSetId', 'termSetName'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
      const res: ContextInfo = await spo.getRequestDigest(spoAdminUrl);

      if (this.verbose) {
        logger.logToStderr(`Retrieving taxonomy term sets...`);
      }

      const termGroupQuery: string = args.options.termGroupId ? `<Method Id="77" ParentId="75" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termGroupId}}</Parameter></Parameters></Method>` : `<Method Id="77" ParentId="75" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termGroupName)}</Parameter></Parameters></Method>`;
      const termSetQuery: string = args.options.termSetId ? `<Method Id="82" ParentId="80" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termSetId}}</Parameter></Parameters></Method>` : `<Method Id="82" ParentId="80" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termSetName)}</Parameter></Parameters></Method>`;
      const data = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="70" ObjectPathId="69" /><ObjectIdentityQuery Id="71" ObjectPathId="69" /><ObjectPath Id="73" ObjectPathId="72" /><ObjectIdentityQuery Id="74" ObjectPathId="72" /><ObjectPath Id="76" ObjectPathId="75" /><ObjectPath Id="78" ObjectPathId="77" /><ObjectIdentityQuery Id="79" ObjectPathId="77" /><ObjectPath Id="81" ObjectPathId="80" /><ObjectPath Id="83" ObjectPathId="82" /><ObjectIdentityQuery Id="84" ObjectPathId="82" /><ObjectPath Id="86" ObjectPathId="85" /><Query Id="87" ObjectPathId="85"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="69" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="72" ParentId="69" Name="GetDefaultSiteCollectionTermStore" /><Property Id="75" ParentId="72" Name="Groups" />${termGroupQuery}<Property Id="80" ParentId="77" Name="TermSets" />${termSetQuery}<Property Id="85" ParentId="82" Name="Terms" /></ObjectPaths></Request>`;

      const result = await this.executeCsomCall(data, spoAdminUrl, res);
      const terms: Term[] = [];
      if (result._Child_Items_ && result._Child_Items_.length > 0) {
        for (const term of result._Child_Items_) {
          this.setTermDetails(term);
          terms.push(term);
          if (args.options.includeChildTerms && term.TermsCount > 0) {
            await this.getChildTerms(spoAdminUrl, res, term);
          }
        }
      }

      if (!args.options.output || args.options.output === 'json') {
        logger.log(terms);
      }
      else if (!args.options.includeChildTerms) {
        // Converted to text friendly output
        logger.log(terms.map(i => {
          return {
            Id: i.Id,
            Name: i.Name
          };
        }));
      }
      else {
        // Converted to text friendly output
        const friendlyOutput: any[] = [];
        terms.forEach(term => {
          term.ParentTermId = '';
          friendlyOutput.push(term);
          if (term.Children && term.Children.length > 0) {
            this.getFriendlyChildTerms(term, friendlyOutput);
          }
        });
        logger.log(friendlyOutput.map(i => {
          return {
            Id: i.Id,
            Name: i.Name,
            ParentTermId: i.ParentTermId
          };
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getFriendlyChildTerms(term: Term, friendlyOutput: any[]): void {
    term.Children.forEach(childTerm => {
      childTerm.ParentTermId = term.Id;
      friendlyOutput.push(childTerm);
      if (childTerm.Children && childTerm.Children.length > 0) {
        this.getFriendlyChildTerms(childTerm, friendlyOutput);
      }
    });
  }

  private async getChildTerms(spoAdminUrl: string, res: ContextInfo, parentTerm: Term): Promise<void> {
    parentTerm.Children = [];
    const data = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="20" ObjectPathId="19" /><Query Id="21" ObjectPathId="19"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="CustomSortOrder" ScalarProperty="true" /><Property Name="CustomProperties" ScalarProperty="true" /><Property Name="LocalCustomProperties" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><Property Id="19" ParentId="16" Name="Terms" /><Identity Id="16" Name="${parentTerm._ObjectIdentity_}" /></ObjectPaths></Request>`;
    const result = await this.executeCsomCall(data, spoAdminUrl, res);
    if (result._Child_Items_ && result._Child_Items_.length > 0) {
      for (const term of result._Child_Items_) {
        this.setTermDetails(term);
        parentTerm.Children.push(term);
        if (term.TermsCount > 0) {
          await this.getChildTerms(spoAdminUrl, res, term);
        }
      }
    }
  }


  private setTermDetails(term: Term): void {
    term.CreatedDate = this.parseTermDateToIsoString(term.CreatedDate);
    term.Id = term.Id.replace('/Guid(', '').replace(')/', '');
    term.LastModifiedDate = this.parseTermDateToIsoString(term.LastModifiedDate);
  }

  private parseTermDateToIsoString(dateAsString: string): string {
    return new Date(Number(dateAsString.replace('/Date(', '').replace(')/', ''))).toISOString();
  }

  private async executeCsomCall(data: string, spoAdminUrl: string, res: ContextInfo): Promise<TermCollection> {
    const requestOptions: AxiosRequestConfig = {
      url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
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
    return json[json.length - 1];
  }
}

module.exports = new SpoTermListCommand();