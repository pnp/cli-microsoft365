import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): Promise<{}> => {
        if (this.verbose) {
          cmd.log(`Adding schema extension with id '${args.options.id}'...`);
        }

        const targetTypes: string[] = args.options.targetTypes.split(',').map(t => t.trim());
        const properties: any = JSON.parse(args.options.properties);

        const requestOptions: any = {
          url: `${auth.service.resource}/v1.0/schemaExtensions`,
          headers: {
            authorization: `Bearer ${auth.service.accessToken}`,
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

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        cmd.log(res);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
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
      if (!args.options.id) {
        return 'Required option id is missing';
      }

      if (!args.options.owner) {
        return 'Required option owner is missing';
      }

      if (args.options.owner && !Utils.isValidGuid(args.options.owner)) {
        return `The specified owner '${args.options.owner}' is not a valid App Id`;
      }

      if (!args.options.targetTypes) {
        return 'Required option targetTypes is missing';
      }

      if (!args.options.properties) {
        return 'Required option targetTypes is missing';
      }

      return this.validateProperties(args.options.properties);
    };
  }

  private validateProperties(propertiesString: string): boolean | string {
    try {
      const properties: any = JSON.parse(propertiesString);

      // If the properties object is not an array
      if (properties.length === undefined) {
        return 'The specified JSON string is not an array';
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
    catch (e) {
      return e;
    }
  }

  private isValidPropertyType(propertyType: string): boolean {
    if (!propertyType) {
      return false;
    }

    return ['Binary', 'Boolean', 'DateTime', 'Integer', 'String'].indexOf(propertyType) > -1;
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to the Microsoft Graph
    using the ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    To create a schema extension, you have to first log in to
    the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

    To create a schema extension, you have to specify a unique ID for the schema
    extension. You can assign a value in one of two ways:

    - concatenate the name of one of your verified domains with a name for
      the schema extension to form a unique string in format
      ${chalk.grey(`{domainName}_{schemaName}`)}, eg. ${chalk.grey(`contoso_mySchema`)}. 

      NOTE: Only verified domains under the following top-level domains are
      supported: .com,.net, .gov, .edu or .org.

    - provide a schema name, and let Microsoft Graph use that schema name to
      complete the id assignment in this format:
      ${chalk.grey(`ext{8-random-alphanumeric-chars}_{schema-name}`)}, eg.
      ${chalk.grey(`extkvbmkofy_mySchema`)}.
      
    The schema extension ID cannot be changed after creation.

    The schema extension owner is the ID of the Azure AD application that is
    the owner of the schema extension. Once set, this property is read-only
    and cannot be changed.

    The target types are the set of Microsoft Graph resource types (that support
    schema extensions) that this schema extension definition can be applied to
    This option is specified as a comma-separated list.

    When specifying the JSON string of properties on Windows, you
    have to escape double quotes in a specific way. Considering the following
    value for the properties option: {"Foo":"Bar"},
    you should specify the value as ${chalk.grey('\`"{""Foo"":""Bar""}"\`')}.
    In addition, when using PowerShell, you should use the --% argument.

  Examples:
  
    Create a schema extension
      ${chalk.grey(config.delimiter)} ${this.name} --id MySchemaExtension --description "My schema extension" --targetTypes Group --owner 62375ab9-6b52-47ed-826b-58e47e0e304b --properties \`"[{""name"":""myProp1"",""type"":""Integer""},{""name"":""myProp2"",""type"":""String""}]\`

    Create a schema extension with a verified domain
      ${chalk.grey(config.delimiter)} ${this.name} --id contoso_MySchemaExtension --description "My schema extension" --targetTypes Group --owner 62375ab9-6b52-47ed-826b-58e47e0e304b --properties \`"[{""name"":""myProp1"",""type"":""Integer""},{""name"":""myProp2"",""type"":""String""}]\`

    Create a schema extension in PowerShell
      ${chalk.grey(config.delimiter)} ${this.name} --id MySchemaExtension --description "My schema extension" --targetTypes Group --owner 62375ab9-6b52-47ed-826b-58e47e0e304b --% --properties \`"[{""name"":""myProp1"",""type"":""Integer""},{""name"":""myProp2"",""type"":""String""}]\`
`);
  }
}

module.exports = new GraphSchemaExtensionAdd();