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
  webUrl: z.string().optional().alias('u'),
  id: z.string().optional().alias('i'),
  name: z.string().optional().alias('n'),
  termGroupId: z.string().optional(),
  termGroupName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoTermSetGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TERM_SET_GET;
  }

  public get description(): string {
    return 'Gets information about the specified taxonomy term set';
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
      }, { error: e => `${(e.input as any).id} is not a valid GUID.` })
      .refine(opts => {
        if (opts.termGroupId && !validation.isValidGuid(opts.termGroupId)) {
          return false;
        }
        return true;
      }, { error: e => `${(e.input as any).termGroupId} is not a valid GUID.` })
      .refine(opts => [opts.id, opts.name].filter(x => x !== undefined).length === 1, {
        message: 'Specify either id or name, but not both.',
        params: { customCode: 'optionSet', options: ['id', 'name'] }
      })
      .refine(opts => [opts.termGroupId, opts.termGroupName].filter(x => x !== undefined).length === 1, {
        message: 'Specify either termGroupId or termGroupName, but not both.',
        params: { customCode: 'optionSet', options: ['termGroupId', 'termGroupName'] }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoWebUrl: string = args.options.webUrl ? args.options.webUrl : await spo.getSpoAdminUrl(logger, this.debug);
      const res: ContextInfo = await spo.getRequestDigest(spoWebUrl);
      if (this.verbose) {
        await logger.logToStderr(`Retrieving taxonomy term set...`);
      }

      const termGroupQuery: string = args.options.termGroupId ? `<Method Id="62" ParentId="60" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.termGroupId}}</Parameter></Parameters></Method>` : `<Method Id="62" ParentId="60" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.termGroupName)}</Parameter></Parameters></Method>`;
      const termSetQuery: string = args.options.id ? `<Method Id="67" ParentId="65" Name="GetById"><Parameters><Parameter Type="Guid">{${args.options.id}}</Parameter></Parameters></Method>` : `<Method Id="67" ParentId="65" Name="GetByName"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.name)}</Parameter></Parameters></Method>`;

      const requestOptions: any = {
        url: `${spoWebUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': res.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="55" ObjectPathId="54" /><ObjectIdentityQuery Id="56" ObjectPathId="54" /><ObjectPath Id="58" ObjectPathId="57" /><ObjectIdentityQuery Id="59" ObjectPathId="57" /><ObjectPath Id="61" ObjectPathId="60" /><ObjectPath Id="63" ObjectPathId="62" /><ObjectIdentityQuery Id="64" ObjectPathId="62" /><ObjectPath Id="66" ObjectPathId="65" /><ObjectPath Id="68" ObjectPathId="67" /><ObjectIdentityQuery Id="69" ObjectPathId="67" /><Query Id="70" ObjectPathId="67"><Query SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticMethod Id="54" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="57" ParentId="54" Name="GetDefaultSiteCollectionTermStore" /><Property Id="60" ParentId="57" Name="Groups" />${termGroupQuery}<Property Id="65" ParentId="62" Name="TermSets" />${termSetQuery}</ObjectPaths></Request>`
      };

      const processQuery: string = await request.post(requestOptions);
      const json: ClientSvcResponse = JSON.parse(processQuery);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }

      const termSet: TermSet = json[json.length - 1];
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

export default new SpoTermSetGetCommand();