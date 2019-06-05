import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { ClientSidePage, ClientSidePart } from './clientsidepages';
import { Page } from './Page';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  webUrl: string;
}

class SpoPageControlListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_CONTROL_LIST}`;
  }

  public get description(): string {
    return 'Lists controls on the specific modern page';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    Page
      .getPage(args.options.name, args.options.webUrl, cmd, this.debug, this.verbose)
      .then((clientSidePage: ClientSidePage): void => {
        let controls: ClientSidePart[] = [];
        clientSidePage.sections.forEach(s => {
          s.columns.forEach(c => {
            controls = controls.concat(c.controls);
          });
        });
        // remove the column property to be able to serialize the array to JSON
        controls.forEach(c => delete c.column);

        // remove the dynamicDataValues and dynamicDataPaths properties if they are null
        controls.forEach(c => {
          if (!c.dynamicDataPaths) {
            delete c.dynamicDataPaths
          }
          if (!c.dynamicDataValues) {
            delete c.dynamicDataValues
          }
        });

        if (args.options.output === 'json') {
          // drop the information about original classes from clientsidepages.ts
          cmd.log(JSON.parse(JSON.stringify(controls)));
        }
        else {
          cmd.log(controls.map(c => {
            return {
              id: c.id,
              type: SpoPageControlListCommand.getControlTypeDisplayName((c as any).controlType),
              title: (c as any).title
            };
          }));
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private static getControlTypeDisplayName(controlType: number): string {
    switch (controlType) {
      case 0:
        return 'Empty column';
      case 3:
        return 'Client-side web part';
      case 4:
        return 'Client-side text';
      default:
        return '' + controlType;
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'Name of the page to list controls of'
      },
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the page to retrieve is located'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.name) {
        return 'Required parameter name missing';
      }

      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    If the specified ${chalk.grey('name')} doesn't refer to an existing modern page, you will get
    a ${chalk.grey('File doesn\'t exists')} error.

  Examples:
  
    List controls on the modern page with name ${chalk.grey('home.aspx')}
      ${this.name} --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx
`);
  }
}

module.exports = new SpoPageControlListCommand();