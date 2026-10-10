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
import { TermSet } from './TermSet.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.string().alias('n'),
  webUrl: z.string().optional().alias('u'),
  termGroupId: z.string().optional(),
  termGroupName: z.string().optional(),
  id: z.string().optional().alias('i'),
  description: z.string().optional().alias('d'),
  customProperties: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoTermSetAddCommand extends SpoCommand {
  public get name(): string {
    return commands.TERM_SET_ADD;
  }

  public get description(): string {
    return 'Adds taxonomy term set';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
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
      }, { error: e => `'${e.input}' is not a valid GUID.` })
      .refine(opts => {
        if (opts.termGroupId && !validation.isValidGuid(opts.termGroupId)) {
          return false;
        }
        return true;
      }, { error: e => `'${e.input}' is not a valid GUID.` })
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
      .refine(opts => [opts.termGroupId, opts.termGroupName].filter(x => x !== undefined).length === 1, {
        message: 'Specify either termGroupId or termGroupName, but not both.',
        params: { customCode: 'optionSet', options: ['termGroupId', 'termGroupName'] }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let termSet: TermSet;

    try {
      const spoWebUrl: string = args.options.webUrl ? args.options.webUrl : await spo.getSpoAdminUrl(logger, this.debug);
      const res: ContextInfo = await spo.getRequestDigest(spoWebUrl);
      const formDigest = res.FormDigestValue;

      if (this.verbose) {
        await logger.logToStderr(`Adding taxonomy term set...`);
      }

      const termGroupQuery: string = args.options.termGroupName ?
        `<Method Id="42" ParentId="40" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termGroupName)}</Parameter></Parameters></Method>` :
        `<Method Id="42" ParentId="40" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termGroupId}}</Parameter></Parameters></Method>`;
      const termSetId: string = args.options.id || v4();

      const requestOptionsQuery: any = {
        url: `${spoWebUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': res.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="35" ObjectPathId="34" /><ObjectIdentityQuery Id="36" ObjectPathId="34" /><ObjectPath Id="38" ObjectPathId="37" /><ObjectIdentityQuery Id="39" ObjectPathId="37" /><ObjectPath Id="41" ObjectPathId="40" /><ObjectPath Id="43" ObjectPathId="42" /><ObjectIdentityQuery Id="44" ObjectPathId="42" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectIdentityQuery Id="47" ObjectPathId="45" /><Query Id="48" ObjectPathId="45"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="34" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="37" ParentId="34" Name="GetDefaultSiteCollectionTermStore" /><Property Id="40" ParentId="37" Name="Groups" />${termGroupQuery}<Method Id="45" ParentId="42" Name="CreateTermSet"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.name)}</Parameter><Parameter Type="Guid">{${termSetId}}</Parameter><Parameter Type="Int32">1033</Parameter></Parameters></Method></ObjectPaths></Request>`
      };

      const processQuery: string = await request.post(requestOptionsQuery);
      const json: ClientSvcResponse = JSON.parse(processQuery);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }

      termSet = json[json.length - 1];

      let terms: string = undefined as any;
      if (!(!args.options.description &&
        !args.options.customProperties)) {
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
          await logger.logToStderr(`Setting term set properties...`);
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
          url: `${spoWebUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': formDigest
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${properties.join('')}<Method Name="CommitAll" Id="131" ObjectPathId="109" /></Actions><ObjectPaths><Identity Id="117" Name="${termSet._ObjectIdentity_}" /><Identity Id="109" Name="${termStoreObjectIdentity}" /></ObjectPaths></Request>`
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

      delete termSet._ObjectIdentity_;
      delete termSet._ObjectType_;
      termSet.CreatedDate = new Date(Number(termSet.CreatedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
      termSet.Id = termSet.Id.replace('/Guid(', '').replace(')/', '');
      termSet.LastModifiedDate = new Date(Number(termSet.LastModifiedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
      await logger.log(termSet);

    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoTermSetAddCommand();