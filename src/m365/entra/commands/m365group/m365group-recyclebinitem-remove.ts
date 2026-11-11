import { Group } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.guid().optional().alias('i'),
  displayName: z.string().optional().alias('d'),
  mailNickname: z.string().optional().alias('m'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraM365GroupRecycleBinItemRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_RECYCLEBINITEM_REMOVE;
  }

  public get description(): string {
    return 'Permanently deletes a Microsoft 365 Group from the recycle bin in the current tenant';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.displayName, opts.mailNickname].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: id, displayName or mailNickname.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'displayName', 'mailNickname']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeGroup: () => Promise<void> = async (): Promise<void> => {
      try {
        const groupId = await this.getGroupId(args.options);

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/directory/deletedItems/${groupId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeGroup();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the group '${args.options.id || args.options.displayName || args.options.mailNickname}'?` });

      if (result) {
        await removeGroup();
      }
    }
  }

  private async getGroupId(options: Options): Promise<string> {
    const { id, displayName, mailNickname } = options;

    if (id) {
      return id;
    }

    let filterValue: string = '';
    if (displayName) {
      filterValue = `displayName eq '${formatting.encodeQueryParameter(displayName)}'`;
    }

    if (mailNickname) {
      filterValue = `mailNickname eq '${formatting.encodeQueryParameter(mailNickname)}'`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=${filterValue}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: Group[] }>(requestOptions);
    const groups = response.value;

    if (groups.length === 0) {
      throw Error(`The specified group '${displayName || mailNickname}' does not exist.`);
    }

    if (groups.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', groups);
      const result = await cli.handleMultipleResultsFound<Group>(`Multiple groups with name '${displayName || mailNickname}' found.`, resultAsKeyValuePair);
      return result.id!;
    }

    return groups[0].id!;
  }
}

export default new EntraM365GroupRecycleBinItemRemoveCommand();