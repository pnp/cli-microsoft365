import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import { CustomAction } from '../customaction/customaction.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: 'Invalid SharePoint URL'
  }).alias('u'),
  title: z.string().optional().alias('t'),
  id: z.string().refine(validation.isValidGuid, { message: 'The value must be a valid GUID.' }).optional().alias('i'),
  clientSideComponentId: z.string().refine(validation.isValidGuid, { message: 'The value must be a valid GUID.' }).optional().alias('c'),
  scope: z.enum(['All', 'Site', 'Web']).optional().alias('s'),
  clientSideComponentProperties: z.boolean().optional().alias('p')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoCommandSetGetCommand extends SpoCommand {
  private static readonly baseLocation: string = 'ClientSideExtension.ListViewCommandSet';
  private static readonly allowedCommandSetLocations: string[] = [SpoCommandSetGetCommand.baseLocation, `${SpoCommandSetGetCommand.baseLocation}.CommandBar`, `${SpoCommandSetGetCommand.baseLocation}.ContextMenu`];

  public get name(): string {
    return commands.COMMANDSET_GET;
  }

  public get description(): string {
    return 'Gets a ListView Command Set that is added to a site.';
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
    try {
      if (this.verbose) {
        await logger.logToStderr(`Attempt to get a specific Command Set by property ${args.options.title || args.options.id || args.options.clientSideComponentId}.`);
      }

      let commandSet: CustomAction;
      if (args.options.id) {
        const customAction = await spo.getCustomActionById(args.options.webUrl, args.options.id, args.options.scope);

        if (customAction === undefined) {
          throw `Command set with id ${args.options.id} can't be found.`;
        }
        else if (!SpoCommandSetGetCommand.allowedCommandSetLocations.some(allowedLocation => allowedLocation === customAction.Location)) {
          throw `Custom action with id ${args.options.id} is not a command set.`;
        }
        commandSet = customAction!;
      }
      else if (args.options.clientSideComponentId) {
        const filter = `${this.getBaseFilter()} ClientSideComponentId eq guid'${args.options.clientSideComponentId}'`;
        const customActions = await spo.getCustomActions(args.options.webUrl, args.options.scope, filter);

        if (customActions.length === 0) {
          throw `No command set with clientSideComponentId '${args.options.clientSideComponentId}' found.`;
        }
        commandSet = customActions[0];
      }
      else if (args.options.title) {
        const filter = `${this.getBaseFilter()} Title eq '${formatting.encodeQueryParameter(args.options.title)}'`;
        const customActions = await spo.getCustomActions(args.options.webUrl, args.options.scope, filter);

        if (customActions.length === 1) {
          commandSet = customActions[0];
        }
        else if (customActions.length === 0) {
          throw `No command set with title '${args.options.title}' found.`;
        }
        else {
          const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', customActions);
          commandSet = await cli.handleMultipleResultsFound<CustomAction>(`Multiple command sets with title '${args.options.title}' found.`, resultAsKeyValuePair);
        }
      }

      if (!args.options.clientSideComponentProperties) {
        await logger.log(commandSet!);
      }
      else {
        const properties = formatting.tryParseJson(commandSet!.ClientSideComponentProperties);
        await logger.log(properties);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getBaseFilter(): string {
    return `startswith(Location,'${SpoCommandSetGetCommand.baseLocation}') and`;
  }
}

export default new SpoCommandSetGetCommand();