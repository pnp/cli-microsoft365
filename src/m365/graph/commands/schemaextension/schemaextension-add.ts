import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  description: string;
  id: string;
  owner: string;
  properties: string;
  targetTypes: string;
}

class GraphSchemaExtensionAdd extends GraphCommand {
  public get name(): string {
    return `${commands.SCHEMAEXTENSION_ADD}`;
  }

  public get description(): string {
    return 'Creates a Microsoft Graph schema extension';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Adding schema extension with id '${args.options.id}'...`);
    }

    const targetTypes: string[] = args.options.targetTypes.split(',').map(t => t.trim());
    const properties: any = JSON.parse(args.options.properties);

    const requestOptions: any = {
      url: `${this.resource}/v1.0/schemaExtensions`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      body: {
        id: args.options.id,
        description: args.options.description,
        owner: args.options.owner,
        targetTypes,
        properties
      },
      json: true
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        cmd.log(res);

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: `The unique identifier for the schema extension definition`
      },
      {
        option: '-d, --description [description]',
        description: 'Description of the schema extension'
      },
      {
        option: '--owner <owner>',
        description: `The ID of the Azure AD application that is the owner of the schema extension`
      },
      {
        option: '-t, --targetTypes <targetTypes>',
        description: `Comma-separated list of Microsoft Graph resource types the schema extension targets`
      },
      {
        option: '-p, --properties <properties>',
        description: `The collection of property names and types that make up the schema extension definition formatted as a JSON string`
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.owner && !Utils.isValidGuid(args.options.owner)) {
        return `The specified owner '${args.options.owner}' is not a valid App Id`;
      }

      return this.validateProperties(args.options.properties);
    };
  }

  private validateProperties(propertiesString: string): boolean | string {
    let result: boolean | string = false;

    try {
      const properties: any = JSON.parse(propertiesString);

      // If the properties object is not an array
      if (properties.length === undefined) {
        
        result = 'The specified JSON string is not an array';

      } else {

        for (let i: number = 0; i < properties.length; i++) {
          const property: any = properties[i];
          if (!property.name) {
            
            result = `Property ${JSON.stringify(property)} misses name`;

          }
          if (!this.isValidPropertyType(property.type)) {
            
            result = `${property.type} is not a valid property type. Valid types are: Binary, Boolean, DateTime, Integer and String`;

          }
        }

        if(typeof result !== "string") {
          result = true;
        };
      }
    }
    catch (e) {
      result = e;
    }

    return result;
  }

  private isValidPropertyType(propertyType: string): boolean {
    if (!propertyType) {
      return false;
    }

    return ['Binary', 'Boolean', 'DateTime', 'Integer', 'String'].indexOf(propertyType) > -1;
  }
}

module.exports = new GraphSchemaExtensionAdd();