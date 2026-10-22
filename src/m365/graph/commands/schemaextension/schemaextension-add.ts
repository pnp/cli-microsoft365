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
  description: z.string().optional().alias('d'),
  owner: z.string(),
  targetTypes: z.string().alias('t'),
  properties: z.string().alias('p')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class GraphSchemaExtensionAddCommand extends GraphCommand {
  public get name(): string {
    return commands.SCHEMAEXTENSION_ADD;
  }

  public get description(): string {
    return 'Creates a Microsoft Graph schema extension';
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
      .refine(options => {
        return this.validateProperties(options.properties) === true;
      }, {
        error: e => `${this.validateProperties((e.input as Options).properties)}`,
        path: ['properties']
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding schema extension with id '${args.options.id}'...`);
    }

    const targetTypes: string[] = args.options.targetTypes.split(',').map(t => t.trim());
    const properties: any = JSON.parse(args.options.properties);

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/schemaExtensions`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      data: {
        id: args.options.id,
        description: args.options.description,
        owner: args.options.owner,
        targetTypes,
        properties
      },
      responseType: 'json'
    };

    try {
      const res = await request.post<any>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private validateProperties(propertiesString: string): boolean | string {
    let result: boolean | string = false;

    try {
      const properties: any = JSON.parse(propertiesString);

      // If the properties object is not an array
      if (properties.length === undefined) {
        result = 'The specified JSON string is not an array';
      }
      else {
        for (let i: number = 0; i < properties.length; i++) {
          const property: any = properties[i];
          if (!property.name) {
            result = `Property ${JSON.stringify(property)} misses name`;
          }

          if (!this.isValidPropertyType(property.type)) {
            result = `${property.type} is not a valid property type. Valid types are: Binary, Boolean, DateTime, Integer and String`;
          }
        }

        if (typeof result !== "string") {
          result = true;
        }
      }
    }
    catch (e: any) {
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

export default new GraphSchemaExtensionAddCommand();