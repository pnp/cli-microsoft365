import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  pageName: string;
  webPartData: string;
}

class SpoAppPageUpdateCommand extends SpoCommand {
  public get name(): string {
    return `${commands.APPPAGE_UPDATE}`;
  }

  public get description(): string {
    return 'Updates the single-part app page';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/sitepages/Pages/UpdateFullPageApp`,
      headers: {
        'content-type': 'application/json;odata=nometadata',
         accept: 'application/json;odata=nometadata'
      },
      json: true,
      body: {
        serverRelativeUrl: `${Utils.getServerRelativePath(args.options.webUrl)}/SitePages/${args.options.pageName}`,
        webPartDataAsJson: args.options.webPartData
      }
    };

    request.post(requestOptions).then((res: any): void => {
      cmd.log(res);
      if (this.verbose) {
        cmd.log(vorpal.chalk.green('DONE'));
      }
      cb();
    }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the page to update is located'
      },
      {
        option: '-p, --pageName <pageName>',
        description: 'The name of the page to be updated'
      },
      {
        option: '-d, --webPartData <webPartData>',
        description: 'JSON string of the web part to update on the page'
      }
      
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {

      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }
      if (!args.options.pageName) {
        return 'Required parameter pageName missing';
      }
      if (!args.options.webPartData) {
        return 'Required parameter webPartData missing';
      }
      try {
        JSON.parse(args.options.webPartData);
      } catch (e) {
        return `Specified webPartData is not a valid JSON string. Error: ${e}`;
      }
      return true;
    };
  } 

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());

    log(`
    
  Examples:   
     
    Updates the single-part app page in a site with url 
    https://contoso.sharepoint.com, webpart data is stored in the
    ${chalk.grey('$webPartData')} variable
      ${this.name} --pageName "Contoso.aspx" --webUrl "https://contoso.sharepoint.com" --webPartData $webPartData 
`);
  }
}
module.exports = new SpoAppPageUpdateCommand();