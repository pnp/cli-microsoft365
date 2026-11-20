import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import config from '../../../../config.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(webUrl => validation.isValidSharePointUrl(webUrl) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  listItemId: z.number().int().positive().alias('i'),
  listId: z.string()
    .refine(listId => validation.isValidGuid(listId), {
      error: e => `${e.input} in option listId is not a valid GUID`
    })
    .optional()
    .alias('l'),
  listTitle: z.string().optional().alias('t'),
  listUrl: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListItemRecordUndeclareCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_RECORD_UNDECLARE;
  }

  public get description(): string {
    return 'Undeclares list item as a record';
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

      if (this.debug) {
        await logger.logToStderr(`getting request digest for request`);
      }

      const reqDigest = await spo.getRequestDigest(args.options.webUrl);
      const formDigestValue = reqDigest.FormDigestValue;

      const objectIdentity = await spo.getCurrentWebIdentity(args.options.webUrl, formDigestValue);

      if (this.verbose) {
        await logger.logToStderr(`Undeclare list item as a record in list ${args.options.listId || args.options.listTitle || args.options.listUrl} in site ${args.options.webUrl}...`);
      }

      const requestOptions: any = {
        url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'Content-Type': 'text/xml',
          'X-RequestDigest': formDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><StaticMethod TypeId="{ea8e1356-5910-4e69-bc05-d0c30ed657fc}" Name="UndeclareItemAsRecord" Id="53"><Parameters><Parameter ObjectPathId="49" /></Parameters></StaticMethod></Actions><ObjectPaths><Identity Id="49" Name="${objectIdentity.objectIdentity}:list:${listId}:item:${args.options.listItemId},1" /></ObjectPaths></Request>`
      };

      await request.post(requestOptions);
      // REST post call doesn't return anything
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

export default new SpoListItemRecordUndeclareCommand();