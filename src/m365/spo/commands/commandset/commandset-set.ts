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
  newClientSideComponentId: z.string().refine(validation.isValidGuid, { message: 'The value must be a valid GUID.' }).optional(),
  newTitle: z.string().optional(),
  description: z.string().optional(),
  listType: z.enum(['List', 'Library', 'SitePages']).optional().alias('l'),
  clientSideComponentProperties: z.string().optional(),
  scope: z.enum(['All', 'Site', 'Web']).optional().alias('s'),
  location: z.enum(['ContextMenu', 'CommandBar', 'Both']).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoCommandSetSetCommand extends SpoCommand {
  public get name(): string {
    return commands.COMMANDSET_SET;
  }

  public get description(): string {
    return 'Updates a ListView Command Set on a site.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.title, opts.id, opts.clientSideComponentId].filter(x => x !== undefined).length === 1, {
        message: `Specify either 'title', 'id' or 'clientSideComponentId', but not multiple.`,
        params: { customCode: 'optionSet', options: ['title', 'id', 'clientSideComponentId'] }
      })
      .refine(opts => !!opts.newTitle || opts.description !== undefined || !!opts.listType || !!opts.clientSideComponentProperties || !!opts.location || !!opts.newClientSideComponentId, {
        message: 'Please specify option to be updated'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Updating ListView Command Set ${args.options.id || args.options.title || args.options.clientSideComponentId} to site '${args.options.webUrl}'...`);
    }

    if (!args.options.scope) {
      args.options.scope = 'All';
    }

    const location: string = this.getLocation(args.options.location ? args.options.location : '');

    try {
      const requestBody: any = {};

      if (args.options.newTitle) {
        requestBody.Title = args.options.newTitle;
      }

      if (args.options.description !== undefined) {
        requestBody.Description = args.options.description;
      }

      if (args.options.location) {
        requestBody.Location = location;
      }

      if (args.options.listType) {
        requestBody.RegistrationId = this.getListTemplate(args.options.listType);
      }

      if (args.options.clientSideComponentProperties) {
        requestBody.ClientSideComponentProperties = args.options.clientSideComponentProperties;
      }

      if (args.options.newClientSideComponentId) {
        requestBody.ClientSideComponentId = args.options.newClientSideComponentId;
      }

      const customAction = await this.getCustomAction(args.options);

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/${customAction.Scope === 3 ? "Web" : "Site"}/UserCustomActions('${formatting.encodeQueryParameter(customAction.Id)}')`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'X-HTTP-Method': 'MERGE'
        },
        data: requestBody,
        responseType: 'json'
      };

      await request.post<CustomAction>(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
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

  private getLocation(location: string): string {
    switch (location) {
      case 'Both':
        return 'ClientSideExtension.ListViewCommandSet';
      case 'ContextMenu':
        return 'ClientSideExtension.ListViewCommandSet.ContextMenu';
      default:
        return 'ClientSideExtension.ListViewCommandSet.CommandBar';
    }
  }

  private getListTemplate(listTemplate: string): string {
    switch (listTemplate) {
      case 'SitePages':
        return '119';
      case 'Library':
        return '101';
      default:
        return '100';
    }
  }
}

export default new SpoCommandSetSetCommand();