import { v4 } from 'uuid';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { Term } from './Term.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.string().alias('n'),
  webUrl: z.string().optional().alias('u'),
  termSetId: z.string().optional(),
  termSetName: z.string().optional(),
  termGroupId: z.string().optional(),
  termGroupName: z.string().optional(),
  id: z.string().optional().alias('i'),
  description: z.string().optional().alias('d'),
  parentTermId: z.string().optional(),
  customProperties: z.string().optional(),
  localCustomProperties: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoTermAddCommand extends SpoCommand {
  public get name(): string {
    return commands.TERM_ADD;
  }

  public get description(): string {
    return 'Adds taxonomy term';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => {
        if (opts.webUrl) {
          return validation.isValidSharePointUrl(opts.webUrl) === true;
        }
        return true;
      }, { message: 'Invalid SharePoint URL' })
      .refine(opts => {
        if (opts.id && !validation.isValidGuid(opts.id)) {
          return false;
        }
        return true;
      }, { error: e => `${(e.input as any).id} is not a valid GUID` })
      .refine(opts => {
        if (opts.parentTermId && !validation.isValidGuid(opts.parentTermId)) {
          return false;
        }
        return true;
      }, { error: e => `${(e.input as any).parentTermId} is not a valid GUID` })
      .refine(opts => {
        if (opts.parentTermId && (opts.termSetId || opts.termSetName)) {
          return false;
        }
        return true;
      }, { message: 'Specify either parentTermId, termSetId or termSetName but not both' })
      .refine(opts => {
        if (opts.termGroupId && !validation.isValidGuid(opts.termGroupId)) {
          return false;
        }
        return true;
      }, { error: e => `${(e.input as any).termGroupId} is not a valid GUID` })
      .refine(opts => {
        if (!opts.termSetId && !opts.termSetName && !opts.parentTermId) {
          return false;
        }
        return true;
      }, { message: 'Specify termSetId, termSetName or parentTermId' })
      .refine(opts => {
        if (opts.termSetId && opts.termSetName) {
          return false;
        }
        return true;
      }, { message: 'Specify termSetId or termSetName but not both' })
      .refine(opts => {
        if (opts.termSetId && !validation.isValidGuid(opts.termSetId)) {
          return false;
        }
        return true;
      }, { error: e => `${(e.input as any).termSetId} is not a valid GUID` })
      .refine(opts => {
        if (opts.customProperties) {
          try {
            JSON.parse(opts.customProperties);
          }
          catch {
            return false;
          }
        }
        return true;
      }, { message: 'customProperties is not valid JSON' })
      .refine(opts => {
        if (opts.localCustomProperties) {
          try {
            JSON.parse(opts.localCustomProperties);
          }
          catch {
            return false;
          }
        }
        return true;
      }, { message: 'localCustomProperties is not valid JSON' })
      .refine(opts => [opts.termGroupId, opts.termGroupName].filter(x => x !== undefined).length === 1, {
        message: 'Specify either termGroupId or termGroupName, but not both.',
        params: { customCode: 'optionSet', options: ['termGroupId', 'termGroupName'] }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let term: Term;
    let formDigest: string;

    try {
      const spoWebUrl: string = args.options.webUrl ? args.options.webUrl : await spo.getSpoAdminUrl(logger, this.debug);
      const res: ContextInfo = await spo.getRequestDigest(spoWebUrl);
      formDigest = res.FormDigestValue;

      if (this.verbose) {
        await logger.logToStderr(`Adding taxonomy term...`);
      }

      const termGroupQuery: string = args.options.termGroupId ? `<Method Id="11" ParentId="9" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termGroupId}}</Parameter></Parameters></Method>` : `<Method Id="11" ParentId="9" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termGroupName)}</Parameter></Parameters></Method>`;
      const termParentQuery: string = args.options.parentTermId ?
        // get parent term by ID
        `<Method Id="16" ParentId="6" Name="GetTerm"><Parameters><Parameter Type="Guid">{${args.options.parentTermId}}</Parameter></Parameters></Method>` :
        // no parent term specified, add to term set
        args.options.termSetId ? `<Method Id="16" ParentId="14" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termSetId}}</Parameter></Parameters></Method>` : `<Method Id="16" ParentId="14" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termSetName)}</Parameter></Parameters></Method>`;
      const termId: string = args.options.id || v4();
      const data: string = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectIdentityQuery Id="13" ObjectPathId="11" /><ObjectPath Id="15" ObjectPathId="14" /><ObjectPath Id="17" ObjectPathId="16" /><ObjectIdentityQuery Id="18" ObjectPathId="16" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectIdentityQuery Id="21" ObjectPathId="19" /><Query Id="22" ObjectPathId="19"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" />${termGroupQuery}<Property Id="14" ParentId="11" Name="TermSets" />${termParentQuery}<Method Id="19" ParentId="16" Name="CreateTerm"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.name)}</Parameter><Parameter Type="Int32">1033</Parameter><Parameter Type="Guid">{${termId}}</Parameter></Parameters></Method></ObjectPaths></Request>`;

      const requestOptionsPost: any = {
        url: `${spoWebUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': res.FormDigestValue
        },
        data: data
      };

      const processQuery: string = await request.post(requestOptionsPost);
      const json: ClientSvcResponse = JSON.parse(processQuery);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }

      term = json[json.length - 1];

      let terms: string = undefined as any;
      if (!(!args.options.description &&
        !args.options.customProperties &&
        !args.options.localCustomProperties)) {
        if (this.verbose) {
          await logger.logToStderr(`Setting term properties...`);
        }

        const properties: string[] = [];
        let i: number = 127;
        if (args.options.description) {
          properties.push(`<Method Name="SetDescription" Id="${i++}" ObjectPathId="117"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.description)}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method>`);
          term.Description = args.options.description;
        }

        if (args.options.customProperties) {
          const customProperties: any = JSON.parse(args.options.customProperties);
          Object.keys(customProperties).forEach(k => {
            properties.push(`<Method Name="SetCustomProperty" Id="${i++}" ObjectPathId="117"><Parameters><Parameter Type="String">${formatting.escapeXml(k)}</Parameter><Parameter Type="String">${formatting.escapeXml(customProperties[k])}</Parameter></Parameters></Method>`);
          });
          term.CustomProperties = customProperties;
        }

        if (args.options.localCustomProperties) {
          const localCustomProperties: any = JSON.parse(args.options.localCustomProperties);
          Object.keys(localCustomProperties).forEach(k => {
            properties.push(`<Method Name="SetLocalCustomProperty" Id="${i++}" ObjectPathId="117"><Parameters><Parameter Type="String">${formatting.escapeXml(k)}</Parameter><Parameter Type="String">${formatting.escapeXml(localCustomProperties[k])}</Parameter></Parameters></Method>`);
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
          url: `${spoWebUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': formDigest
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${properties.join('')}<Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="${term._ObjectIdentity_}" /><Identity Id="109" Name="${termStoreObjectIdentity}" /></ObjectPaths></Request>`
        };

        terms = await request.post(requestOptions);
      }

      if (terms) {
        const json: ClientSvcResponse = JSON.parse(terms);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          throw response.ErrorInfo.ErrorMessage;
        }
      }

      delete term._ObjectIdentity_;
      delete term._ObjectType_;
      term.CreatedDate = new Date(Number(term.CreatedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
      term.Id = term.Id.replace('/Guid(', '').replace(')/', '');
      term.LastModifiedDate = new Date(Number(term.LastModifiedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
      await logger.log(term);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

export default new SpoTermAddCommand();