import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  owner: string;
  description?: string;
  status?: string;
  targetTypes?: string;
  properties?: string;
}

class GraphSchemaExtensionSetCommand extends GraphCommand {
  public get name(): string {
    return commands.SCHEMAEXTENSION_SET;
  }

  public get description(): string {
    return 'Updates a Microsoft Graph schema extension';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.description = typeof args.options.description !== 'undefined';
    telemetryProps.properties = typeof args.options.properties !== 'undefined';
    telemetryProps.targetTypes = typeof args.options.targetTypes !== 'undefined';
    telemetryProps.status = args.options.status;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Updating schema extension with id '${args.options.id}'...`);
    }

    // The default request body always contains owner
    const body: {
      owner: string;
      description?: string;
      status?: string;
      targetTypes?: string[];
      properties?: any;
    } = {
      owner: args.options.owner
    };

    // Add the description to request body if any
    if (args.options.description) {
      if (this.debug) {
        cmd.log(`Will update description to '${args.options.description}'...`);
      }
      body.description = args.options.description;
    }

    // Add the status to request body if any
    if (args.options.status) {
      if (this.debug) {
        cmd.log(`Will update status to '${args.options.status}'...`);
      }
      body.status = args.options.status;
    }

    // Add the target types to request body if any
    const targetTypes: string[] = args.options.targetTypes
      ? args.options.targetTypes.split(',').map(t => t.trim())
      : [];
    if (targetTypes.length > 0) {
      if (this.debug) {
        cmd.log(`Will update targetTypes to '${args.options.targetTypes}'...`);
      }
      body.targetTypes = targetTypes;
    }

    // Add the properties to request body if any
    const properties: any = args.options.properties
      ? JSON.parse(args.options.properties)
      : null;
    if (properties) {
      if (this.debug) {
        cmd.log(`Will update properties to '${args.options.properties}'...`);
      }
      body.properties = properties;
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/schemaExtensions/${args.options.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      body,
      json: true
    };

    request
      .patch(requestOptions)
      .then((res: any): void => {
        if (this.debug) {
          cmd.log("Schema extension successfully updated.");
        }

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
        option: '--owner <owner>',
        description: `The ID of the Azure AD application that is the owner of the schema extension`
      },
      {
        option: '-d, --description [description]',
        description: 'Description of the schema extension'
      },
      {
        option: '-s, --status [status]',
        description: `The lifecycle state of the schema extension. Accepted values are 'Available' or 'Deprecated'`
      },
      {
        option: '-t, --targetTypes [targetTypes]',
        description: `Comma-separated list of Microsoft Graph resource types the schema extension targets`
      },
      {
        option: '-p, --properties [properties]',
        description: `The collection of property names and types that make up the schema extension definition formatted as a JSON string`
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.owner)) {
        return `The specified owner '${args.options.owner}' is not a valid App Id`;
      }

      if (!args.options.status && !args.options.properties && !args.options.targetTypes && !args.options.description) {
        return `No updates were specified. Please specify at least one argument among --status, --targetTypes, --description or --properties`
      }

      const validStatusValues = ['Available', 'Deprecated'];
      if (args.options.status && validStatusValues.indexOf(args.options.status) < 0) {
        return `Status option is invalid. Valid statuses are: Available or Deprecated`;
      }

      if (args.options.properties) {
        return this.validateProperties(args.options.properties);
      }

      return true;
    };
  }

  private validateProperties(propertiesString: string): boolean | string {
    let properties: any = null;
    try {
      properties = JSON.parse(propertiesString);
    }
    catch (e) {
      return 'The specified properties is not a valid JSON string';
    }

    // If the properties object is not an array
    if (properties.length === undefined) {
      return 'The specified properties JSON string is not an array';
    }

    for (let i: number = 0; i < properties.length; i++) {
      const property: any = properties[i];
      if (!property.name) {
        return `Property ${JSON.stringify(property)} misses name`;
      }
      if (!this.isValidPropertyType(property.type)) {
        return `${property.type} is not a valid property type. Valid types are: Binary, Boolean, DateTime, Integer and String`;
      }
    }

    return true;
  }

  private isValidPropertyType(propertyType: string): boolean {
    if (!propertyType) {
      return false;
    }

    return ['Binary', 'Boolean', 'DateTime', 'Integer', 'String'].indexOf(propertyType) > -1;
  }
}

module.exports = new GraphSchemaExtensionSetCommand();