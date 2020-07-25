import commands from '../../commands';
import flowCommands from '../../../flow/commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import request from '../../../../request';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import * as path from 'path';
import * as fs from 'fs';
import { Connector } from './Connector';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  connector: string;
  environment: string;
  outputFolder?: string;
}

class PaConnectorExportCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.CONNECTOR_EXPORT;
  }

  public get description(): string {
    return 'Exports the specified power automate or power apps custom connector';
  }

  public alias(): string[] | undefined {
    return [flowCommands.CONNECTOR_EXPORT];
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const outputFolder = path.resolve(args.options.outputFolder || '.', args.options.connector);

    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.PowerApps/apis/${encodeURIComponent(args.options.connector)}?api-version=2016-11-01&$filter=environment%20eq%20%27${encodeURIComponent(args.options.environment)}%27%20and%20IsCustomApi%20eq%20%27True%27`,
      headers: {
        accept: 'application/json'
      },
      json: true
    };

    let connector: Connector;

    if (this.verbose) {
      cmd.log('Downloading connector...');
    }
    request
      .get<Connector>(requestOptions)
      .then((connectorRes: Connector): Promise<string> => {
        connector = connectorRes;

        if (!connector.properties) {
          return Promise.reject('Properties not present in the api registration information.');
        }

        if (this.verbose) {
          cmd.log(`Creating output folder ${outputFolder}...`);
        }
        fs.mkdirSync(outputFolder);

        const settings: any = {
          apiDefinition: "apiDefinition.swagger.json",
          apiProperties: "apiProperties.json",
          connectorId: args.options.connector,
          environment: args.options.environment,
          icon: "icon.png",
          powerAppsApiVersion: "2016-11-01",
          powerAppsUrl: "https://api.powerapps.com"
        };
        if (this.verbose) {
          cmd.log('Exporting settings...');
        }
        fs.writeFileSync(path.join(outputFolder, 'settings.json'), JSON.stringify(settings, null, 2), 'utf8');

        const propertiesWhitelist: string[] = [
          "connectionParameters",
          "iconBrandColor",
          "capabilities",
          "policyTemplateInstances"
        ];

        const apiProperties: any = {
          properties: JSON.parse(JSON.stringify(connector.properties))
        };
        Object.keys(apiProperties.properties).forEach(k => {
          if (propertiesWhitelist.indexOf(k) < 0) {
            delete apiProperties.properties[k];
          }
        });
        if (this.verbose) {
          cmd.log('Exporting API properties...');
        }
        fs.writeFileSync(path.join(outputFolder, 'apiProperties.json'), JSON.stringify(apiProperties, null, 2), 'utf8');

        if (connector.properties.apiDefinitions &&
          connector.properties.apiDefinitions.originalSwaggerUrl) {
          if (this.verbose) {
            cmd.log(`Downloading swagger from ${connector.properties.apiDefinitions.originalSwaggerUrl}...`);
          }
          return request
            .get({
              url: connector.properties.apiDefinitions.originalSwaggerUrl,
              headers: {
                'x-anonymous': true
              }
            });
        }
        else {
          if (this.debug) {
            cmd.log('originalSwaggerUrl not set. Skipping');
          }
          return Promise.resolve('');
        }
      })
      .then((swagger: string): Promise<any> => {
        if (swagger && swagger.length > 0) {
          if (this.debug) {
            cmd.log('Downloaded swagger');
            cmd.log(swagger);
          }
          if (this.verbose) {
            cmd.log('Exporting swagger...');
          }
          fs.writeFileSync(path.join(outputFolder, 'apiDefinition.swagger.json'), swagger, 'utf8');
        }

        if (connector.properties.iconUri) {
          if (this.verbose) {
            cmd.log(`Downloading icon from ${connector.properties.iconUri}...`);
          }
          return request
            .get({
              url: connector.properties.iconUri,
              encoding: null,
              headers: {
                'x-anonymous': true
              }
            });
        }
        else {
          if (this.debug) {
            cmd.log('iconUri not set. Skipping');
          }
          return Promise.resolve();
        }
      })
      .then((icon: any): void => {
        if (icon) {
          if (this.verbose) {
            cmd.log('Exporting icon...');
          }
          const iconBuffer: Buffer = Buffer.from(icon, 'utf8');
          fs.writeFileSync(path.join(outputFolder, 'icon.png'), iconBuffer);
        }
        else {
          if (this.debug) {
            cmd.log('No icon retrieved');
          }
        }
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }
        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-e, --environment <environment>',
        description: 'The name of the environment where the custom connector to export is located'
      },
      {
        option: '-c, --connector <connector>',
        description: 'The name of the custom connector to export'
      },
      {
        option: '--outputFolder [outputFolder]',
        description: 'Path where the folder with connector\'s files should be saved. If not specified, will create the connector\'s folder in the current folder.'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.outputFolder &&
        !fs.existsSync(path.resolve(args.options.outputFolder))) {
        return `Specified output folder ${args.options.outputFolder} doesn't exist`;
      }

      const outputFolder = path.resolve(args.options.outputFolder || '.', args.options.connector);
      if (fs.existsSync(outputFolder)) {
        return `Connector output folder ${outputFolder} already exists`;
      }

      return true;
    };
  }
}

module.exports = new PaConnectorExportCommand();