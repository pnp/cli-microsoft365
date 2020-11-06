import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ClientSideControl } from './ClientSideControl';
import { ClientSidePageProperties } from './ClientSidePageProperties';
import { ClientSidePage, ClientSidePart } from './clientsidepages';
import { Page } from './Page';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  name: string;
  webUrl: string;
  webPartData?: string;
  webPartProperties?: string;
}

class SpoPageControlSetCommand extends SpoCommand {

  public get name(): string {
    return `${commands.PAGE_CONTROL_SET}`;
  }

  public get description(): string {
    return 'Allows you to set or update control webpart data or properties on a modern page';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    Page
      .getPage(args.options.name, args.options.webUrl, logger, this.debug, this.verbose)
      .then((clientSidePage: ClientSidePage): Promise<ClientSidePageProperties> => {
        let control: ClientSidePart | null = clientSidePage.findControlById(args.options.id);

        if (control) {
          if (this.verbose) {
            logger.log(`Control with ID ${args.options.id} found on the page`);
          }

          // Check out the page
          return Page.checkout(args.options.name, args.options.webUrl, logger, this.debug, this.verbose);
        } else {
          return Promise.reject(`Control with ID ${args.options.id} not found on page ${args.options.name}`);
        }
      })
      .then((page: ClientSidePageProperties) => {
        // Update the web part data
        const canvasContent: ClientSideControl[] = JSON.parse(page.CanvasContent1);
        if (this.debug) {
          logger.log(canvasContent);
        }

        let csControl = canvasContent.find(c => c.id === args.options.id);

        if (csControl) {
          if (args.options.webPartData) {
            if (this.verbose) {
              logger.log('WebPart data:');
              logger.log(args.options.webPartData);
              logger.log('');
            }
        
            const webPartData = JSON.parse(args.options.webPartData);
        
            csControl.webPartData = {
              ...csControl.webPartData,
              ...webPartData,
              id: csControl.webPartData.id,
              instanceId: csControl.webPartData.instanceId
            };
            
            if (this.verbose) {
              logger.log('Updated webpart data:');
              logger.log(csControl.webPartData);
              logger.log('');
            }
          }
          
          if (args.options.webPartProperties) {
            if (this.verbose) {
              logger.log('webPartProperties data:');
              logger.log(args.options.webPartProperties);
              logger.log('');
            }
        
            const webpartProperties = JSON.parse(args.options.webPartProperties);
        
            csControl.webPartData.properties = {
              ...csControl.webPartData.properties,
              ...webpartProperties
            };        
        
            if (this.verbose) {
              logger.log('Updated webpart properties:');
              logger.log(csControl.webPartData.properties);
              logger.log('');
            }
          }
        
          return Page.save(args.options.name, args.options.webUrl, canvasContent, logger, this.debug, this.verbose);
        } else {
          return Promise.reject(`Control with ID ${args.options.id} not found on page ${args.options.name}`);
        }
      })
      .then(() => {
        if (this.verbose) {
          logger.log(chalk.green('DONE'));
        }
        cb();
      })
      .catch((err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'ID of the control to set new properties for'
      },
      {
        option: '-n, --name <name>',
        description: 'Name of the page where the control is located'
      },
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the page to retrieve is located'
      },
      {
        option: '--webPartData [webPartData]',
        description: 'JSON string with web part data as retrieved from the web part maintenance mode. Specify webPartProperties or webPartData but not both'
      },
      {
        option: '--webPartProperties [webPartProperties]',
        description: 'JSON string with web part data as retrieved from the web part maintenance mode. Specify webPartProperties or webPartData but not both'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.id)) {
      return `${args.options.id} is not a valid GUID`;
    }

    if (args.options.webPartProperties) {
      try {
        JSON.parse(args.options.webPartProperties);
      } catch (e) {
        return `Specified webPartProperties is not a valid JSON string. Input: ${args.options.webPartData}. Error: ${e}`;
      }
    }

    if (args.options.webPartData) {
      try {
        JSON.parse(args.options.webPartData);
      } catch (e) {
        return `Specified webPartData is not a valid JSON string. Input: ${args.options.webPartData}. Error: ${e}`;
      }
    }

    if (args.options.webPartData && args.options.webPartProperties) {
      return 'Specify webPartProperties or webPartData but not both';
    }

    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoPageControlSetCommand();
