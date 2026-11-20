import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).alias('u'),
  listTitle: z.string().optional().alias('l'),
  listId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `${e.input} is not a valid GUID`
  }).optional(),
  listUrl: z.string().optional(),
  id: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `${e.input} is not a valid GUID`
  }).optional().alias('i'),
  title: z.string().optional().alias('t'),
  internalName: z.string().optional(),
  group: z.string().optional().alias('g'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFieldRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FIELD_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified list- or site column';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.title, opts.internalName, opts.group].filter(o => o !== undefined).length === 1, {
        error: 'Specify one of the following options: id, title, internalName, group.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'title', 'internalName', 'group']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let messageEnd: string;
    if (args.options.listId || args.options.listTitle) {
      messageEnd = `in list ${args.options.listId || args.options.listTitle}`;
    }
    else {
      messageEnd = `in site ${args.options.webUrl}`;
    }

    const removeField = async (listRestUrl: string, fieldId?: string | undefined, title?: string | undefined, internalName?: string | undefined): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing field ${fieldId || title || internalName} ${messageEnd}...`);
      }

      let fieldRestUrl: string;
      if (fieldId) {
        fieldRestUrl = `/getbyid('${formatting.encodeQueryParameter(fieldId)}')`;
      }
      else {
        fieldRestUrl = `/getbyinternalnameortitle('${formatting.encodeQueryParameter(title as string || internalName as string)}')`;
      }

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/${listRestUrl}fields${fieldRestUrl}`,
        method: 'POST',
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*',
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    };

    const prepareRemoval = async (): Promise<void> => {
      let listRestUrl: string = '';

      if (args.options.listId) {
        listRestUrl = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
      }
      else if (args.options.listTitle) {
        listRestUrl = `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/`;
      }
      else if (args.options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
        listRestUrl = `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
      }

      if (args.options.group) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving fields assigned to group ${args.options.group}...`);
        }
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/${listRestUrl}fields`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        try {
          const res = await request.get<any>(requestOptions);
          const filteredResults = res.value.filter((field: { Id: string | undefined, Group: string | undefined; }) => field.Group === args.options.group);
          if (this.verbose) {
            await logger.logToStderr(`${filteredResults.length} matches found...`);
          }

          const promises = [];
          for (let index = 0; index < filteredResults.length; index++) {
            promises.push(removeField(listRestUrl, filteredResults[index].Id));
          }

          await Promise.all(promises);
        }
        catch (err: any) {
          this.handleRejectedODataJsonPromise(err);
        }
      }
      else {
        try {
          await removeField(listRestUrl, args.options.id, args.options.title, args.options.internalName);
        }
        catch (err: any) {
          this.handleRejectedODataJsonPromise(err);
        }
      }
    };

    if (args.options.force) {
      await prepareRemoval();
    }
    else {
      const confirmMessage: string = `Are you sure you want to remove the ${args.options.group ? 'fields' : 'field'} ${args.options.id || args.options.title || args.options.internalName || 'from group ' + args.options.group} ${messageEnd}?`;

      const result = await cli.promptForConfirmation({ message: confirmMessage });

      if (result) {
        await prepareRemoval();
      }
    }
  }
}

export default new SpoFieldRemoveCommand();