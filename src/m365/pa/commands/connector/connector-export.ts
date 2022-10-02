import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import PowerAppsCommand from '../../../base/PowerAppsCommand';
import flowCommands from '../../../flow/commands';
import commands from '../../commands';
import { Connector } from './Connector';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  connector: string;
  environment: string;
  outputFolder?: string;
}

class PaConnectorExportCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.CONNECTOR_EXPORT;
  }

  public get description(): string {
    return 'Exports the specified power automate or power apps custom connector';
  }

  public alias(): string[] | undefined {
    return [flowCommands.CONNECTOR_EXPORT];
  }

  constructor() {
    super();
  
    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }
  
  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        outputFolder: typeof args.options.outputFolder !== 'undefined'
      });
    });
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environment <environment>'
      },
      {
        option: '-c, --connector <connector>'
      },
      {
        option: '--outputFolder [outputFolder]'
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.outputFolder &&
          !fs.existsSync(path.resolve(args.options.outputFolder))) {
          return `Specified output folder ${args.options.outputFolder} doesn't exist`;
        }
    
        const outputFolder = path.resolve(args.options.outputFolder || '.', args.options.connector);
        if (fs.existsSync(outputFolder)) {
          return `Connector output folder ${outputFolder} already exists`;
        }
    
        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const outputFolder = path.resolve(args.options.outputFolder || '.', args.options.connector);

    const requestOptions: any = {
      url: `${this.resource}/providers/Microsoft.PowerApps/apis/${encodeURIComponent(args.options.connector)}?api-version=2016-11-01&$filter=environment%20eq%20%27${encodeURIComponent(args.options.environment)}%27%20and%20IsCustomApi%20eq%20%27True%27`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    let connector: Connector;

    if (this.verbose) {
      logger.logToStderr('Downloading connector...');
    }

    try {
      connector = await request.get<Connector>(requestOptions);

      if (!connector.properties) {
        throw 'Properties not present in the api registration information.';
      }

      if (this.verbose) {
        logger.logToStderr(`Creating output folder ${outputFolder}...`);
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
        logger.logToStderr('Exporting settings...');
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
        logger.logToStderr('Exporting API properties...');
      }
      fs.writeFileSync(path.join(outputFolder, 'apiProperties.json'), JSON.stringify(apiProperties, null, 2), 'utf8');

      let swagger = '';
      if (connector.properties.apiDefinitions &&
        connector.properties.apiDefinitions.originalSwaggerUrl) {
        if (this.verbose) {
          logger.logToStderr(`Downloading swagger from ${connector.properties.apiDefinitions.originalSwaggerUrl}...`);
        }
        swagger = await request
          .get({
            url: connector.properties.apiDefinitions.originalSwaggerUrl,
            headers: {
              'x-anonymous': 'true'
            }
          });
      }
      else {
        if (this.debug) {
          logger.logToStderr('originalSwaggerUrl not set. Skipping');
        }
      }

      if (swagger && swagger.length > 0) {
        if (this.debug) {
          logger.logToStderr('Downloaded swagger');
          logger.logToStderr(swagger);
        }
        if (this.verbose) {
          logger.logToStderr('Exporting swagger...');
        }
        fs.writeFileSync(path.join(outputFolder, 'apiDefinition.swagger.json'), swagger, 'utf8');
      }

      let icon = '';
      if (connector.properties.iconUri) {
        if (this.verbose) {
          logger.logToStderr(`Downloading icon from ${connector.properties.iconUri}...`);
        }
        icon = await request
          .get({
            url: connector.properties.iconUri,
            responseType: 'arraybuffer',
            headers: {
              'x-anonymous': 'true'
            }
          });
      }
      else {
        if (this.debug) {
          logger.logToStderr('iconUri not set. Skipping');
        }
      }

      if (icon) {
        if (this.verbose) {
          logger.logToStderr('Exporting icon...');
        }
        const iconBuffer: Buffer = Buffer.from(icon, 'utf8');
        fs.writeFileSync(path.join(outputFolder, 'icon.png'), iconBuffer);
      }
      else {
        if (this.debug) {
          logger.logToStderr('No icon retrieved');
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PaConnectorExportCommand();