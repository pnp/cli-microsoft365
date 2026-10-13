import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.object({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).alias('u'),
  listId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `${e.input} in option listId is not a valid GUID`
  }).optional(),
  listTitle: z.string().optional(),
  listUrl: z.string().optional(),
  id: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `${e.input} in option id is not a valid GUID`
  }).optional().alias('i'),
  title: z.string().optional().alias('t'),
  internalName: z.string().optional(),
  updateExistingLists: z.boolean().optional()
}).catchall(z.unknown());

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFieldSetCommand extends SpoCommand {
  public get name(): string {
    return commands.FIELD_SET;
  }

  public get description(): string {
    return 'Updates existing list or site column';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.title, opts.internalName].filter(o => o !== undefined).length === 1, {
        error: 'Specify one of the following options: id, title, internalName.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'title', 'internalName']
        }
      })
      .refine(opts => [opts.listId, opts.listTitle, opts.listUrl].filter(o => o !== undefined).length <= 1, {
        error: 'Specify either list id or title or list url, but not multiple'
      });
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const reqDigest = await spo.getRequestDigest(args.options.webUrl);
      const requestDigest = reqDigest.FormDigestValue;

      let fieldsParentIdentity = '<Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />';

      if (args.options.listId || args.options.listTitle || args.options.listUrl) {
        let requestData = '';
        if (args.options.listId) {
          requestData = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="664" ObjectPathId="663" /><Query Id="665" ObjectPathId="663"><Query SelectAllProperties="false"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="663" ParentId="7" Name="GetById"><Parameters><Parameter Type="Guid">${formatting.escapeXml(args.options.listId)}</Parameter></Parameters></Method><Property Id="7" ParentId="5" Name="Lists" /><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`;
        }
        else if (args.options.listTitle) {
          requestData = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="664" ObjectPathId="663" /><Query Id="665" ObjectPathId="663"><Query SelectAllProperties="false"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="663" ParentId="7" Name="GetByTitle"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.listTitle)}</Parameter></Parameters></Method><Property Id="7" ParentId="5" Name="Lists" /><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`;
        }
        else if (args.options.listUrl) {
          const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
          requestData = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="5"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticProperty Id="1" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="3" ParentId="1" Name="Web" /><Method Id="5" ParentId="3" Name="GetList"><Parameters><Parameter Type="String">${listServerRelativeUrl}</Parameter></Parameters></Method></ObjectPaths></Request>`;
        }

        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': requestDigest
          },
          data: requestData
        };

        const list = await request.post<string>(requestOptions);
        const json: ClientSvcResponse = JSON.parse(list);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          throw response.ErrorInfo.ErrorMessage;
        }

        const result: { _ObjectIdentity_: string; } = json[json.length - 1];
        fieldsParentIdentity = `<Identity Id="5" Name="${result._ObjectIdentity_}" />`;
      }

      // retrieve column CSOM object id
      const fieldQuery: string = args.options.id ?
        `<Method Id="663" ParentId="7" Name="GetById"><Parameters><Parameter Type="Guid">${formatting.escapeXml(args.options.id)}</Parameter></Parameters></Method>` :
        `<Method Id="663" ParentId="7" Name="GetByInternalNameOrTitle"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.title || args.options.internalName)}</Parameter></Parameters></Method>`;

      let requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': requestDigest
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="664" ObjectPathId="663" /><Query Id="665" ObjectPathId="663"><Query SelectAllProperties="false"><Properties /></Query></Query></Actions><ObjectPaths>${fieldQuery}<Property Id="7" ParentId="5" Name="Fields" />${fieldsParentIdentity}</ObjectPaths></Request>`
      };

      const field = await request.post<string>(requestOptions);
      let json: ClientSvcResponse = JSON.parse(field);
      let response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }

      const result: { _ObjectIdentity_: string; } = json[json.length - 1];
      const fieldId: string = result._ObjectIdentity_;

      requestOptions = {
        url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': requestDigest
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions>${this.getPayload(args.options)}<Method Name="UpdateAndPushChanges" Id="9000" ObjectPathId="663"><Parameters><Parameter Type="Boolean">${args.options.updateExistingLists ? 'true' : 'false'}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="663" Name="${fieldId}" /></ObjectPaths></Request>`
      };

      const res = await request.post<string>(requestOptions);
      json = JSON.parse(res);
      response = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private getPayload(options: any): string {
    const excludeOptions: string[] = [
      'webUrl',
      'listId',
      'listTitle',
      'listUrl',
      'id',
      'title',
      'internalName',
      'updateExistingLists',
      'debug',
      'verbose',
      'output'
    ];

    let i: number = 667;
    const payload: string = Object.keys(options).map(key => {
      return excludeOptions.indexOf(key) === -1 ? `<SetProperty Id="${i++}" ObjectPathId="663" Name="${key}"><Parameter Type="String">${formatting.escapeXml(options[key])}</Parameter></SetProperty>` : '';
    }).join('');

    return payload;
  }
}

export default new SpoFieldSetCommand();