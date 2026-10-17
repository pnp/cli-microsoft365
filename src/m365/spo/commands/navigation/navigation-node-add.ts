import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { MenuState, MenuStateNode } from './NavigationNode.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  location: z.enum(['QuickLaunch', 'TopNavigationBar']).optional().alias('l'),
  title: z.string().min(1, 'Cannot be empty.').alias('t'),
  url: z.string().optional(),
  parentNodeId: z.int().positive().optional(),
  audienceIds: z.string()
    .refine(audienceIds => audienceIds.split(',').length <= 10, {
      error: 'The maximum amount of audienceIds per navigation node exceeded. The maximum amount of audienceIds is 10.'
    })
    .refine(audienceIds => validation.isValidGuidArray(audienceIds) === true, {
      error: e => `The following GUIDs are invalid for the option 'audienceIds': ${validation.isValidGuidArray(e.input as string)}.`
    })
    .optional(),
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

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.location, options.parentNodeId].filter(o => o !== undefined).length === 1, {
        error: `Specify either 'location' or 'parentNodeId', but not both.`,
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
    const { webUrl, location, parentNodeId, title, url, audienceIds, openInNewWindow } = args.options;

    if (this.verbose) {
      await logger.logToStderr(`Adding navigation node...`);
    }

    try {
      let menuState: MenuState;
      let siblingNodes: MenuStateNode[];

      if (parentNodeId) {
        const parent = await spo.getMenuStateNodeByKey(webUrl, parentNodeId.toString());
        menuState = parent.menuState;
        siblingNodes = parent.node.Nodes;
      }
      else {
        menuState = location === 'TopNavigationBar'
          ? await spo.getTopNavigationMenuState(webUrl)
          : await spo.getQuickLaunchMenuState(webUrl);
        siblingNodes = menuState.Nodes;
      }

      const existingKeys: (string | null)[] = siblingNodes.map(node => node.Key);
      const newNode: Partial<MenuStateNode> = {
        AudienceIds: audienceIds ? audienceIds.split(',') : [],
        IsDeleted: false,
        IsHidden: false,
        Key: null,
        Nodes: [],
        NodeType: 0,
        OpenInNewWindow: openInNewWindow ? true : null,
        SimpleUrl: url ?? '',
        Title: title
      };
      siblingNodes.push(newNode as MenuStateNode);
      await spo.saveMenuState(webUrl, menuState);

      if (this.verbose) {
        await logger.logToStderr(`Retrieving the added navigation node...`);
      }

      const updatedMenuState: MenuState = await spo.getMenuState(webUrl, menuState.StartingNodeKey);
      const updatedSiblingNodes: MenuStateNode[] | undefined = parentNodeId
        ? spo.findMenuStateNode(updatedMenuState.Nodes, parentNodeId.toString())?.Nodes
        : updatedMenuState.Nodes;
      const addedNode: MenuStateNode | undefined = updatedSiblingNodes?.find(node => !existingKeys.includes(node.Key) && node.Title === title);

      if (!addedNode) {
        throw `The navigation node was added, but it couldn't be retrieved.`;
      }

      await logger.log(addedNode);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoNavigationNodeAddCommand();