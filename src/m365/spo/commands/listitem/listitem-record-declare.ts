import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import config from "../../../../config.js";
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from "../../../../utils/formatting.js";
import { ClientSvcResponse, ClientSvcResponseContents, spo } from "../../../../utils/spo.js";
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from "../../../../utils/validation.js";
import SpoCommand from "../../../base/SpoCommand.js";
import commands from "../../commands.js";

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(webUrl => validation.isValidSharePointUrl(webUrl) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  listId: z.string()
    .refine(listId => validation.isValidGuid(listId), {
      error: e => `${e.input} in option listId is not a valid GUID`
    })
    .optional()
    .alias('l'),
  listTitle: z.string().optional().alias('t'),
  listUrl: z.string().optional(),
  listItemId: z.number().int().positive().alias('i'),
  date: z.string()
    .refine(date => validation.isValidISODate(date), {
      error: e => `${e.input} in option date is not in ISO format (yyyy-mm-dd)`
    })
    .optional()
    .alias('d')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListItemRecordDeclareCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_RECORD_DECLARE;
  }

  public get description(): string {
    return "Declares the specified list item as a record";
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.listId, opts.listTitle, opts.listUrl].filter(o => o !== undefined).length === 1, {
        error: 'Specify one of the following options: listId, listTitle, listUrl.',
        params: {
          customCode: 'optionSet',
          options: ['listId', 'listTitle', 'listUrl']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let listId: string = '';

      if (args.options.listId) {
        listId = args.options.listId;
      }
      else {
        let requestUrl = `${args.options.webUrl}/_api/web`;

        if (args.options.listTitle) {
          requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;
        }
        else if (args.options.listUrl) {
          const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
          requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }

        const requestOptions: CliRequestOptions = {
          url: `${requestUrl}?$select=Id`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const list = await request.get<{ Id: string; }>(requestOptions);
        listId = list.Id;
      }

      const contextResponse = await spo.getRequestDigest(args.options.webUrl);
      const formDigestValue = contextResponse.FormDigestValue;

      const webIdentityResp = await spo.getCurrentWebIdentity(args.options.webUrl, formDigestValue);
      const webIdentity = webIdentityResp.objectIdentity;
      const requestBody: string = this.getDeclareRecordRequestBody(webIdentity, listId, args.options.listItemId, args.options.date || '');

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'Content-Type': 'text/xml',
          'X-RequestDigest': formDigestValue
        },
        data: requestBody
      };

      const res = await request.post<string>(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];

      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      else {
        const result: boolean = json[json.length - 1];
        await logger.log(result);
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  protected getDeclareRecordRequestBody(webIdentity: string, listId: string, id: number, date: string): string {
    let requestBody: string;
    if (date.length === 10) {
      requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><StaticMethod TypeId="{ea8e1356-5910-4e69-bc05-d0c30ed657fc}" Name="DeclareItemAsRecordWithDeclarationDate" Id="48"><Parameters><Parameter ObjectPathId="21" /><Parameter Type="DateTime">${date}</Parameter></Parameters></StaticMethod></Actions><ObjectPaths><Identity Id="21" Name="${webIdentity}:list:${listId}:item:${id},1" /></ObjectPaths></Request>`;
    }
    else {
      requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><StaticMethod TypeId="{ea8e1356-5910-4e69-bc05-d0c30ed657fc}" Name="DeclareItemAsRecord" Id="37"><Parameters><Parameter ObjectPathId="12" /></Parameters></StaticMethod></Actions><ObjectPaths><Identity Id="12" Name="${webIdentity}:list:${listId}:item:${id},1" /></ObjectPaths></Request>`;
    }

    return requestBody;
  }
}
export default new SpoListItemRecordDeclareCommand();