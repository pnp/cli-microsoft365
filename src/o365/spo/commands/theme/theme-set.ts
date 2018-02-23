import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
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
  isInverted: boolean;
}

class SpoThemeSetCommand extends SpoCommand {
  public get name(): string {
    return commands.THEME_SET;
  }

  public get description(): string {
    return 'Add or update a theme';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.inverted = (!(!args.options.isInverted)).toString();
    return telemetryProps;
  }

  protected requiresTenantAdmin(): boolean {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Setting theme for the ${auth.site.url} tenant...`);
        }

        return this.getRequestDigest(cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const fullPath: string = path.resolve(args.options.filePath);

        if (this.verbose) {
          cmd.log(`Adding theme from ${fullPath} to tenant...`);
        }

        const palette: any = JSON.parse(fs.readFileSync(fullPath, 'utf8'));

        if (this.debug) {
          cmd.log('');
          cmd.log('Palette');
          cmd.log(JSON.stringify(palette));
        }

        const isInverted: boolean = args.options.isInverted ? true : false;

        const requestOptions: any = {
          url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'X-RequestDigest': res.FormDigestValue
          }),
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="UpdateTenantTheme" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.name)}</Parameter><Parameter Type="String">{"isInverted":${isInverted},"name":"${Utils.escapeXml(args.options.name)}","palette":${JSON.stringify(palette)}}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];

        if (response.ErrorInfo) {
          cmd.log(new CommandError(response.ErrorInfo.ErrorMessage));
        }
        else {
          const result: boolean = json[json.length - 1];
          cmd.log(result);
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-n, --name <name>',
      description: 'Name of the theme to add or update'
    },
    {
      option: '-p, --filePath <filePath>',
      description: 'Absolute or relative path to the theme json file'
    },
    {
      option: '--isInverted',
      description: 'Set to specify that the theme is inverted'
    }];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.name) {
        return 'Required parameter name missing';
      }

      if (!args.options.filePath) {
        return 'Required parameter filePath missing';
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
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online tenant admin site,
  using the ${chalk.blue(commands.CONNECT)} command.

  Remarks:
  
    To add or update a theme, you have to first connect to a tenant admin site using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso-admin.sharepoint.com`)}.
        
  Examples:
  
    Add or update a theme from a theme JSON file
      ${chalk.grey(config.delimiter)} ${commands.THEME_SET} --name Contoso-Blue --filePath /Users/rjesh/themes/contoso-blue.json

    Add or update an inverted theme from a theme JSON file
      ${chalk.grey(config.delimiter)} ${commands.THEME_SET} --name Contoso-Blue --filePath /Users/rjesh/themes/contoso-blue.json --isInverted
    
  More information:

    SharePoint site theming
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview

    Theme Generator
      https://developer.microsoft.com/en-us/fabric#/styles/themegenerator
      `);
  }
}

module.exports = new SpoThemeSetCommand();