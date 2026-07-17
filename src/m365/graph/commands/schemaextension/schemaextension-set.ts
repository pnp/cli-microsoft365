import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i'),
  owner: z.string(),
  description: z.string().optional().alias('d'),
  status: z.enum(['Available', 'Deprecated']).optional().alias('s'),
  targetTypes: z.string().optional().alias('t'),
  properties: z.string().optional().alias('p')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class GraphSchemaExtensionSetCommand extends GraphCommand {
  public get name(): string {
    return commands.SCHEMAEXTENSION_SET;
  }

  public get description(): string {
    return 'Updates a Microsoft Graph schema extension';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => validation.isValidGuid(options.owner), {
        error: e => `The specified owner '${(e.input as Options).owner}' is not a valid App Id`,
        path: ['owner']
      })
      .refine(options => options.status || options.properties || options.targetTypes || options.description, {
        error: 'No updates were specified. Please specify at least one argument among --status, --targetTypes, --description or --properties'
      })
      .refine(options => {
        if (!options.properties) {
          return true;
        }
        return this.validateProperties(options.properties) === true;
      }, {
        error: e => `${this.validateProperties((e.input as Options).properties!)}`,
        path: ['properties']
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Updating schema extension with id '${args.options.id}'...`);
    }

    // The default request data always contains owner
    const data: {
      owner: string;
      description?: string;
      status?: string;
      targetTypes?: string[];
      properties?: any;
    } = {
      owner: args.options.owner
    };

    // Add the description to request data if any
    if (args.options.description) {
      if (this.debug) {
        await logger.logToStderr(`Will update description to '${args.options.description}'...`);
      }
      data.description = args.options.description;
    }

    // Add the status to request data if any
    if (args.options.status) {
      if (this.debug) {
        await logger.logToStderr(`Will update status to '${args.options.status}'...`);
      }
      data.status = args.options.status;
    }

    // Add the target types to request data if any
    const targetTypes: string[] = args.options.targetTypes
      ? args.options.targetTypes.split(',').map(t => t.trim())
      : [];
    if (targetTypes.length > 0) {
      if (this.debug) {
        await logger.logToStderr(`Will update targetTypes to '${args.options.targetTypes}'...`);
      }
      data.targetTypes = targetTypes;
    }

    // Add the properties to request data if any
    const properties: any = args.options.properties
      ? JSON.parse(args.options.properties)
      : null;
    if (properties) {
      if (this.debug) {
        await logger.logToStderr(`Will update properties to '${args.options.properties}'...`);
      }
      data.properties = properties;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/schemaExtensions/${args.options.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      data,
      responseType: 'json'
    };

    try {
      await request.patch(requestOptions);

      if (this.debug) {
        await logger.logToStderr("Schema extension successfully updated.");
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private validateProperties(propertiesString: string): boolean | string {
    let properties: any;
    try {
      properties = JSON.parse(propertiesString);
    }
    catch {
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

export default new GraphSchemaExtensionSetCommand();