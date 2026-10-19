import { z } from 'zod';
import { Auth } from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import config from '../../../../config.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.number().int().positive().alias('i'),
  listId: z.string()
    .refine(listId => validation.isValidGuid(listId), {
      error: e => `${e.input} in option listId is not a valid GUID`
    })
    .optional()
    .alias('l'),
  listTitle: z.string().optional().alias('t'),
  listUrl: z.string().optional(),
  webUrl: z.string()
    .refine(webUrl => validation.isValidSharePointUrl(webUrl) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListItemIsRecordCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_ISRECORD;
  }

  public get description(): string {
    return 'Checks if the specified list item is a record';
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
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let requestUrl = `${args.options.webUrl}/_api/web`;

    if (args.options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
    }
    else if (args.options.listTitle) {
      requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    let listId: string;

    if (this.debug) {
      await logger.logToStderr(`Retrieving access token for ${resource}...`);
    }

    try {
      if (typeof args.options.listId !== 'undefined') {
        if (this.verbose) {
          await logger.logToStderr(`List Id passed in as an argument.`);
        }

        listId = args.options.listId;
      }
      else {
        if (this.verbose) {
          await logger.logToStderr(`Getting list id for list ${args.options.listTitle ? args.options.listTitle : args.options.listUrl}`);
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
        await logger.logToStderr(`Getting request digest for request`);
      }

      const reqDigest = await spo.getRequestDigest(args.options.webUrl);
      const formDigestValue = reqDigest.FormDigestValue;

      const webIdentityResp = await spo.getCurrentWebIdentity(args.options.webUrl, formDigestValue);

      if (this.verbose) {
        await logger.logToStderr(`Checking if list item is a record in list ${args.options.listId ? args.options.listId : args.options.listTitle ? args.options.listTitle : args.options.listUrl} in site ${args.options.webUrl}...`);
      }

      const requestBody = this.getIsRecordRequestBody(webIdentityResp.objectIdentity, listId, args.options.id);
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

  private getIsRecordRequestBody(webIdentity: string, listId: string, id: number): string {
    const requestBody: string = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
            <Actions>
              <StaticMethod TypeId="{ea8e1356-5910-4e69-bc05-d0c30ed657fc}" Name="IsRecord" Id="1"><Parameters><Parameter ObjectPathId="14" /></Parameters></StaticMethod>
            </Actions>
            <ObjectPaths>
              <Identity Id="14" Name="${webIdentity}:list:${listId}:item:${id},1" />
            </ObjectPaths>
          </Request>`;

    return requestBody;
  }
}

export default new SpoListItemIsRecordCommand();