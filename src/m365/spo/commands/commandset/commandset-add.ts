import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { CustomAction } from '../customaction/customaction.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  title: z.string().alias('t'),
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).alias('u'),
  listType: z.enum(['List', 'Library', 'SitePages']).alias('l'),
  clientSideComponentId: z.string().refine(validation.isValidGuid, { message: 'The value must be a valid GUID.' }).alias('i'),
  description: z.string().optional(),
  clientSideComponentProperties: z.string().optional(),
  scope: z.enum(['Site', 'Web']).optional().alias('s'),
  location: z.enum(['ContextMenu', 'CommandBar', 'Both']).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoCommandSetAddCommand extends SpoCommand {
  public get name(): string {
    return commands.COMMANDSET_ADD;
  }

  public get description(): string {
    return 'Adds a ListView Command Set to a site.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding ListView Command Set ${args.options.clientSideComponentId} to site '${args.options.webUrl}'...`);
    }

    if (!args.options.scope) {
      args.options.scope = 'Web';
    }

    const location: string | undefined = args.options.location && this.getLocation(args.options.location);
    const listType: string = this.getListTemplate(args.options.listType);

    try {
      const requestBody: any = {
        Title: args.options.title,
        Description: args.options.description,
        Location: location,
        ClientSideComponentId: args.options.clientSideComponentId,
        RegistrationId: listType,
        RegistrationType: 1
      };

      if (args.options.clientSideComponentProperties) {
        requestBody.ClientSideComponentProperties = args.options.clientSideComponentProperties;
      }

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/${args.options.scope}/UserCustomActions`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        data: requestBody,
        responseType: 'json'
      };
      const response = await request.post<CustomAction>(requestOptions);

      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
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

export default new SpoCommandSetAddCommand();