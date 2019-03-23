import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
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
  id: string;
  description: string;
  owner: string;
  targetTypes: string;
  properties: string;
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
      .then((): request.RequestPromise => {
        if (this.verbose) {
          cmd.log(`Adding schema extension with id '${args.options.id}'...`);
        }

        const targetTypes: string[] = args.options.targetTypes.split(',').map(t => t.trim());
        let properties: any = JSON.parse(args.options.properties);

        const requestOptions: any = {
          url: `${auth.service.resource}/v1.0/schemaExtensions`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          }),
          body: {
            id: args.options.id,
            description: args.options.description,
            owner: args.options.owner,
            targetTypes,
            properties
          },
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }

        cmd.log(`Schema extension added with Id '${args.options.id}'`);

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
        description: `The Id of the Azure AD application that is the owner of the schema extension`
      },
      {
        option: '-t, --targetTypes <types>',
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

      if (!this._validatePropertiesStringValue(args.options.properties)) {
        return `The specified value for properties option '${args.options.properties}' is not a valid properties definition`;
      }

      return true;
    };
  }

  private _validatePropertiesStringValue(properties: string): boolean {
    
    const checkValidJson = Utils.isValidJsonString(properties);
    if (!checkValidJson.isValid) {
      return false;
    }

    return this._validateProperties(checkValidJson.parsedObject);
  }

  private _validateProperties(properties: any): boolean {
    // If the properties object is not an array
    if (properties.length === undefined) {
      return false;
    }

    const invalidProps = (properties as any[]).filter(p => !p.name || !this._isValidPropertyType(p.type));
    console.log('invalid props: ');
    console.log(JSON.stringify(invalidProps));
    return invalidProps.length == 0;
  }

  private _isValidPropertyType(propertyType: string): boolean {
    if (!propertyType) {
      return false;
    }

    switch (propertyType.toLowerCase()) {
      case "binary":
      case "boolean":
      case "datetime":
      case "integer":
      case "string":
        return true;
      default:
        return false;
    }
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

    To create a schema extension, you have to specify a unique ID for the schema extension
    You can assign a value in one of two ways:
    - Concatenate the name of one of your verified domains with a name for the schema extension to form a unique string in this format, {domainName}_{schemaName}.
    As an example, contoso_mySchema. 
    NOTE: Only verified domains under the following top-level domains are supported: .com,.net, .gov, .edu or .org.
    - Provide a schema name, and let Microsoft Graph use that schema name to complete the id assignment in this format: ext{8-random-alphanumeric-chars}_{schema-name}.
    An example would be extkvbmkofy_mySchema.
    This property cannot be changed after creation.

    The schema extension owner is the Id of the Azure AD application that is the owner of the schema extension.
    It must be specified, otherwise, Micrsoft Graph will return an 'Unauthorized' error 
    Once set, this property is read-only and cannot be changed.

    The target types are the set of Microsoft Graph resource types (that support schema extensions) that this schema extension definition can be applied to
    This option is specified as a comma-separated list

    When specifying the JSON string of properties on Windows, you
    have to escape double quotes in a specific way. Considering the following
    value for the _properties_ option: {"Foo":"Bar"},
    you should specify the value as \`"{""Foo"":""Bar""}"\`.
    In addition, when using PowerShell, you should use the --% argument.

  Examples:
  
   Create a schema extension
      ${chalk.grey(config.delimiter)} ${this.name} --id MySchemaExtension --description "My schema extension" --targetTypes Group --owner 62375ab9-6b52-47ed-826b-58e47e0e304b --properties \`"[{""name"":""myProp1"",""type"":""Integer""},{""name"":""myProp2"",""type"":""String""}]\`

   Create a schema extension with a verified domain (Let's assume you own contoso.com and it is registered as verified domain on your tenant)
      ${chalk.grey(config.delimiter)} ${this.name} --id contoso_MySchemaExtension --description "My schema extension" --targetTypes Group --owner 62375ab9-6b52-47ed-826b-58e47e0e304b --properties \`"[{""name"":""myProp1"",""type"":""Integer""},{""name"":""myProp2"",""type"":""String""}]\`

   Create a schema extension (in a PowerShell console)
      ${chalk.grey(config.delimiter)} ${this.name} --id MySchemaExtension --description "My schema extension" --targetTypes Group --owner 62375ab9-6b52-47ed-826b-58e47e0e304b --% --properties \`"[{""name"":""myProp1"",""type"":""Integer""},{""name"":""myProp2"",""type"":""String""}]\`
`);
  }
}

module.exports = new GraphSchemaExtensionAdd();