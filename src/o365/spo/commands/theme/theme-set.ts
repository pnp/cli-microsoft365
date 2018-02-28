import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import * as fs from 'fs';
import * as path from 'path';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  filePath: string;
}

class ThemeSetCommand extends SpoCommand {

  public get name(): string {
    return commands.THEME_SET;
  }

  public get description(): string {
    return 'Add or update theme to tenant with the given palette';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Adding new theme to the tenant store...`);
        }

        const fullPath: string = path.resolve(args.options.filePath);

        if (this.verbose) {
          cmd.log(`Adding theme from ${fullPath} to tenant...`);
        }

        const palette: any = {
          "palette": JSON.parse(fs.readFileSync(fullPath, 'utf8'))
        }

        if (this.debug) {
          cmd.log('');
          cmd.log('Palette');
          cmd.log(JSON.stringify(palette));
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_api/thememanager/AddTenantTheme`,
          method: 'POST',
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          body: {
            "name": args.options.name,
            "themeJson": JSON.stringify(palette)
          },
          json: true
        };

        if (args.options.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((rawRes: string): void => {

        if (args.options.debug) {
          cmd.log('Response:');
          cmd.log(rawRes);
          cmd.log('');
        }

        if (args.options.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => {
        this.handleRejectedODataPromise(err, cmd, cb)
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
        option: '-n, --name <name>',
        description: 'name of the theme getting added'
      },
      {
        option: '-p, --filePath <filePath>',
        description: 'Absolute or relative path to the theme json file to add to the tenant theme store'
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

      if (!args.options.filePath) {
        return 'Required parameter file path missing';
      }

      const fullPath: string = path.resolve(args.options.filePath);

      if (!fs.existsSync(fullPath)) {
        return `File '${fullPath}' not found`;
      }

      if (fs.lstatSync(fullPath).isDirectory()) {
        return `Path '${fullPath}' points to a directory`;
      }
      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
      using the ${chalk.blue(commands.CONNECT)} command.
  
    Remarks:
    
      To add/update theme, you have to first connect to SharePoint using the
      ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
          
    Examples:
    
      To add or update theme to the tenant from absolute or relative path of given theme json file
      ${chalk.grey(config.delimiter)} ${commands.THEME_SET} -n Contoso-Blue -p /Users/rjesh/themes/contoso-blue.json
      
    More information:

      Create custom theme using Office Fabric theme generator tool, 
        copy the JSON output and save as JSON file.clear
      https://developer.microsoft.com/en-us/fabric#/styles/themegenerator
      `);
  }
}

module.exports = new ThemeSetCommand();