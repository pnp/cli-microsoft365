import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { MenuState, MenuStateNode, NavigationNode } from './NavigationNode';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  audienceIds?: string;
  isExternal?: boolean;
  location?: string;
  parentNodeId?: number;
  title: string;
  url: string;
  webUrl: string;
  openInNewWindow?: boolean
}

class SpoNavigationNodeAddCommand extends SpoCommand {
  private static readonly allowedLocations: string[] = ['QuickLaunch', 'TopNavigationBar'];

  public get name(): string {
    return commands.NAVIGATION_NODE_ADD;
  }

  public get description(): string {
    return 'Adds a navigation node to the specified site navigation';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        isExternal: !!args.options.isExternal,
        location: typeof args.options.location !== 'undefined',
        parentNodeId: typeof args.options.parentNodeId !== 'undefined',
        audienceIds: typeof args.options.audienceIds !== 'undefined',
        openInNewWindow: !!args.options.openInNewWindow
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --location [location]',
        autocomplete: SpoNavigationNodeAddCommand.allowedLocations
      },
      {
        option: '-t, --title <title>'
      },
      {
        option: '--url <url>'
      },
      {
        option: '--parentNodeId [parentNodeId]'
      },
      {
        option: '--isExternal'
      },
      {
        option: '--audienceIds [audienceIds]'
      },
      {
        option: '--openInNewWindow'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.parentNodeId) {
          if (isNaN(args.options.parentNodeId)) {
            return `${args.options.parentNodeId} is not a number`;
          }
        }
        else {
          if (!SpoNavigationNodeAddCommand.allowedLocations.some(allowedLocation => allowedLocation === args.options.location)) {
            return `${args.options.location} is not a valid value for the location option. Allowed values are ${SpoNavigationNodeAddCommand.allowedLocations.join('|')}`;
          }
        }

        if (args.options.audienceIds) {
          const audienceIdsSplitted = args.options.audienceIds.split(',');
          if (audienceIdsSplitted.length > 10) {
            return 'The maximum amount of audienceIds per navigation node exceeded. The maximum amount of auciendeIds is 10.';
          }

          if (!validation.isValidGuidArray(audienceIdsSplitted)) {
            return 'The option audienceIds contains one or more invalid GUIDs';
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['location', 'parentNodeId'] }
    );
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Adding navigation node...`);
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
        Url: args.options.url,
        IsExternal: args.options.isExternal === true
      }
    };

    try {
      const res = await request.post<NavigationNode>(requestOptions);

      if (args.options.openInNewWindow) {
        if (this.verbose) {
          logger.logToStderr(`Making sure that the newly added navigation node opens in a new window.`);
        }

        let menuState: MenuState | undefined;
        let menuStateItem: MenuStateNode | undefined;
        if (args.options.parentNodeId) {
          menuState = await spo.getQuickLaunchMenuState(args.options.webUrl);
          menuStateItem = this.getMenuStateNode(menuState.Nodes, res.Id.toString());
          if (!menuStateItem) {
            menuState = await spo.getTopNavigationMenuState(args.options.webUrl);
            menuStateItem = this.getMenuStateNode(menuState.Nodes, res.Id.toString());
          }
        }
        else {
          if (args.options.location === 'QuickLaunch') {
            menuState = await spo.getQuickLaunchMenuState(args.options.webUrl);
            menuStateItem = this.getMenuStateNode(menuState.Nodes, res.Id.toString());
          }
          else {
            menuState = await spo.getTopNavigationMenuState(args.options.webUrl);
            menuStateItem = this.getMenuStateNode(menuState.Nodes, res.Id.toString());
          }
        }
        menuStateItem!.OpenInNewWindow = true;
        await spo.saveMenuState(args.options.webUrl, menuState!);
      }

      logger.log(res);
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

module.exports = new SpoNavigationNodeAddCommand();