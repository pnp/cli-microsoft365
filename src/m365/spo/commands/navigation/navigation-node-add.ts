import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { MenuState, MenuStateNode, NavigationNode } from './NavigationNode.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    message: 'The value of the option \'webUrl\' is not a valid SharePoint Online site URL.'
  }).alias('u'),
  location: z.enum(['QuickLaunch', 'TopNavigationBar']).optional().alias('l'),
  title: z.string().alias('t'),
  url: z.string().optional(),
  parentNodeId: z.coerce.number().refine((val: number) => !isNaN(val), {
    message: 'parentNodeId is not a number'
  }).optional(),
  isExternal: z.boolean().optional(),
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
  openInNewWindow: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoNavigationNodeAddCommand extends SpoCommand {
  public get name(): string {
    return commands.NAVIGATION_NODE_ADD;
  }

  public get description(): string {
    return 'Adds a navigation node to the specified site navigation';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => [opts.location, opts.parentNodeId].filter(x => x !== undefined).length === 1, {
        message: 'Specify either location or parentNodeId, but not both.',
        params: {
          customCode: 'optionSet',
          options: ['location', 'parentNodeId']
        }
      });
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding navigation node...`);
    }

    const nodesCollection: string = args.options.parentNodeId ?
      `GetNodeById(${args.options.parentNodeId})/Children` :
      (args.options.location as string).toLowerCase();

    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/web/navigation/${nodesCollection}`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        AudienceIds: args.options.audienceIds?.split(','),
        Title: args.options.title,
        Url: args.options.url ?? 'http://linkless.header/',
        IsExternal: args.options.isExternal === true
      }
    };

    try {
      const res = await request.post<NavigationNode>(requestOptions);

      if (args.options.openInNewWindow) {
        if (this.verbose) {
          await logger.logToStderr(`Making sure that the newly added navigation node opens in a new window.`);
        }

        const id: string = res.Id.toString();

        let menuState: MenuState = args.options.location === 'TopNavigationBar' ? await spo.getTopNavigationMenuState(args.options.webUrl) : await spo.getQuickLaunchMenuState(args.options.webUrl);
        let menuStateItem: MenuStateNode = this.getMenuStateNode(menuState.Nodes, id);

        if (args.options.parentNodeId && !menuStateItem) {
          menuState = await spo.getTopNavigationMenuState(args.options.webUrl);
          menuStateItem = this.getMenuStateNode(menuState.Nodes, id);
        }

        menuStateItem.OpenInNewWindow = true;
        await spo.saveMenuState(args.options.webUrl, menuState);
      }

      await logger.log(res);
    }

    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getMenuStateNode(nodes: MenuStateNode[], id: string): MenuStateNode {
    let menuNode = nodes.find((node: MenuStateNode) => node.Key !== null && node.Key === id);
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

export default new SpoNavigationNodeAddCommand();