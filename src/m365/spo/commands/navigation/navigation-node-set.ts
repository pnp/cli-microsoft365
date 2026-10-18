import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { MenuStateNode } from './NavigationNode.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    message: 'The value of the option \'webUrl\' is not a valid SharePoint Online site URL.'
  }).alias('u'),
  id: z.coerce.number().refine((val: number) => !isNaN(val), {
    message: 'id is not a number'
  }),
  title: z.string().optional(),
  url: z.string().optional(),
  audienceIds: z.string().refine((val: string) => {
    if (val === '') {
      return true;
    }
    const items = val.split(',').map((s: string) => s.trim());
    if (items.length > 10) {
      return false;
    }
    const invalid = items.filter((item: string) => !validation.isValidGuid(item));
    return invalid.length === 0;
  }, { error: (issue: any) => {
    const val = issue.input as string;
    const items = val.split(',').map((s: string) => s.trim());
    if (items.length > 10) {
      return 'The maximum amount of audienceIds per navigation node exceeded. The maximum amount of audienceIds is 10.';
    }
    const invalid = items.filter((item: string) => !validation.isValidGuid(item));
    return `The following GUIDs are invalid for the option 'audienceIds': ${invalid.join(', ')}.`;
  } }).optional(),
  isExternal: z.boolean().optional(),
  openInNewWindow: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoNavigationNodeSetCommand extends SpoCommand {
  public get name(): string {
    return commands.NAVIGATION_NODE_SET;
  }

  public get description(): string {
    return 'Adds a navigation node to the specified site navigation';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => opts.audienceIds !== undefined || opts.url !== undefined || opts.isExternal !== undefined || opts.title !== undefined || opts.openInNewWindow !== undefined, {
        message: 'Please specify at least one property to update.',
        params: {
          customCode: 'required'
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Setting navigation node...`);
      }

      let url = args.options.url;
      if (url === '') {
        url = 'http://linkless.header/';
      }

      const requestBody: any = {
        Title: args.options.title,
        IsExternal: args.options.isExternal,
        Url: url
      };

      if (args.options.audienceIds !== undefined) {
        requestBody.AudienceIds = args.options.audienceIds === '' ? [] : args.options.audienceIds.split(',');
      }

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/navigation/GetNodeById(${args.options.id})`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata'
        },
        data: requestBody,
        responseType: 'json'
      };

      const response = await request.patch<any>(requestOptions);
      if (response['odata.null'] === true) {
        throw `Navigation node does not exist.`;
      }

      if (args.options.openInNewWindow !== undefined) {
        if (this.verbose) {
          await logger.logToStderr(`Making sure that the navigation node opens in a new window.`);
        }

        const idStr = args.options.id.toString();
        let menuState = await spo.getQuickLaunchMenuState(args.options.webUrl);
        let menuStateItem = this.getMenuStateNode(menuState.Nodes, idStr);
        if (!menuStateItem) {
          menuState = await spo.getTopNavigationMenuState(args.options.webUrl);
          menuStateItem = this.getMenuStateNode(menuState.Nodes, idStr);
        }
        menuStateItem!.OpenInNewWindow = args.options.openInNewWindow;
        await spo.saveMenuState(args.options.webUrl, menuState);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getMenuStateNode(nodes: MenuStateNode[], id: string): MenuStateNode {
    let menuNode = nodes.find((node: MenuStateNode) => node.Key.toString() === id.toString());
    if (menuNode === undefined) {
      for (const node of nodes.filter(node => node.Nodes.length > 0)) {
        menuNode = this.getMenuStateNode(node.Nodes, id);
        if (menuNode) {
          break;
        }
      }
    }
    return menuNode!;
  }
}

export default new SpoNavigationNodeSetCommand();