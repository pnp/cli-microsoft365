import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { CustomAction } from '../customaction/customaction.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { cli } from '../../../../cli/cli.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).alias('u'),
  title: z.string().optional().alias('t'),
  id: z.string().refine(validation.isValidGuid, { message: 'The value must be a valid GUID.' }).optional().alias('i'),
  clientSideComponentId: z.string().refine(validation.isValidGuid, { message: 'The value must be a valid GUID.' }).optional().alias('c'),
  scope: z.enum(['All', 'Site', 'Web']).optional().alias('s'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoCommandSetRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.COMMANDSET_REMOVE;
  }

  public get description(): string {
    return 'Removes a ListView Command Set that is added to a site.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(opts => [opts.title, opts.id, opts.clientSideComponentId].filter(x => x !== undefined).length === 1, {
      message: `Specify either 'title', 'id' or 'clientSideComponentId', but not multiple.`,
      params: { customCode: 'optionSet', options: ['title', 'id', 'clientSideComponentId'] }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing ListView Command Set ${args.options.clientSideComponentId || args.options.title || args.options.id} to site '${args.options.webUrl}'...`);
    }

    if (args.options.force) {
      await this.deleteCommandset(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove command set '${args.options.clientSideComponentId || args.options.title || args.options.id}'?` });

      if (result) {
        await this.deleteCommandset(args);
      }
    }
  }

  private async getCustomAction(options: Options): Promise<CustomAction> {
    let commandSets: CustomAction[] = [];

    if (options.id) {
      const commandSet = await spo.getCustomActionById(options.webUrl, options.id, options.scope);
      if (commandSet) {
        commandSets.push(commandSet);
      }
    }
    else if (options.title) {
      commandSets = await spo.getCustomActions(options.webUrl, options.scope, `(Title eq '${formatting.encodeQueryParameter(options.title as string)}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`);
    }
    else {
      commandSets = await spo.getCustomActions(options.webUrl, options.scope, `(ClientSideComponentId eq guid'${options.clientSideComponentId}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`);
    }

    if (commandSets.length === 0) {
      throw `No user commandsets with ${options.title && `title '${options.title}'` || options.clientSideComponentId && `ClientSideComponentId '${options.clientSideComponentId}'` || options.id && `id '${options.id}'`} found`;
    }

    if (commandSets.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', commandSets);
      return await cli.handleMultipleResultsFound<CustomAction>(`Multiple user commandsets with ${options.title ? `title '${options.title}'` : `ClientSideComponentId '${options.clientSideComponentId}'`} found.`, resultAsKeyValuePair);
    }

    return commandSets[0];
  }

  private async deleteCommandset(args: CommandArgs): Promise<void> {
    if (!args.options.scope) {
      args.options.scope = 'All';
    }

    try {
      const customAction = await this.getCustomAction(args.options);

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/${customAction.Scope === 3 ? "Web" : "Site"}/UserCustomActions('${formatting.encodeQueryParameter(customAction.Id)}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      await request.delete<CustomAction>(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoCommandSetRemoveCommand();