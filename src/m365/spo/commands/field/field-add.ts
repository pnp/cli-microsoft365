import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listTitle?: string;
  xml: string;
  options?: string;
}

class SpoFieldAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.FIELD_ADD}`;
  }

  public get description(): string {
    return 'Adds a new list or site column using the CAML field definition';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getRequestDigest(args.options.webUrl)
      .then((res: ContextInfo): Promise<{}> => {
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/${(args.options.listTitle ? `lists/getByTitle('${encodeURIComponent(args.options.listTitle)}')/` : '')}fields/CreateFieldAsXml`,
          headers: {
            'X-RequestDigest': res.FormDigestValue,
            accept: 'application/json;odata=nometadata'
          },
          body: {
            parameters: {
              SchemaXml: args.options.xml,
              Options: this.getOptions(args.options.options)
            }
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
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getOptions(options?: string): number {
    let optionsValue: number = 0;

    if (!options) {
      return optionsValue;
    }

    options.split(',').forEach(o => {
      o = o.trim();
      switch (o) {
        case 'DefaultValue':
          optionsValue += 0;
          break;
        case 'AddToDefaultContentType':
          optionsValue += 1;
          break;
        case 'AddToNoContentType':
          optionsValue += 2;
          break;
        case 'AddToAllContentTypes':
          optionsValue += 4;
          break;
        case 'AddFieldInternalNameHint':
          optionsValue += 8;
          break;
        case 'AddFieldToDefaultView':
          optionsValue += 16;
          break;
        case 'AddFieldCheckDisplayName':
          optionsValue += 32;
          break;
      }
    });

    return optionsValue;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'Absolute URL of the site where the field should be created'
      },
      {
        option: '-l, --listTitle [listTitle]',
        description: 'Title of the list where the field should be created (if it should be created as a list column)'
      },
      {
        option: '-x, --xml <xml>',
        description: 'CAML field definition'
      },
      {
        option: '--options [options]',
        description: 'The options to use to add to the field. Allowed values: DefaultValue|AddToDefaultContentType|AddToNoContentType|AddToAllContentTypes|AddFieldInternalNameHint|AddFieldToDefaultView|AddFieldCheckDisplayName'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter url missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (!args.options.xml) {
        return 'Required parameter xml missing';
      }

      if (args.options.options) {
        let optionsError: string | boolean = true;
        const options: string[] = ['DefaultValue', 'AddToDefaultContentType', 'AddToNoContentType', 'AddToAllContentTypes', 'AddFieldInternalNameHint', 'AddFieldToDefaultView', 'AddFieldCheckDisplayName'];
        args.options.options.split(',').forEach(o => {
          o = o.trim();
          if (options.indexOf(o) < 0) {
            optionsError = `${o} is not a valid value for the options argument. Allowed values are DefaultValue|AddToDefaultContentType|AddToNoContentType|AddToAllContentTypes|AddFieldInternalNameHint|AddFieldToDefaultView|AddFieldCheckDisplayName`;
          }
        });
        return optionsError;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    If the specified field already exists, you will get a
    ${chalk.grey('A duplicate field name "your-field" was found.')} error.

  Examples:
  
    Create a date time site column
      ${this.name} --webUrl https://contoso.sharepoint.com/sites/contoso-sales --xml '\`<Field Type="DateTime" DisplayName="Start date-time" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateTime" Group="PnP Columns" FriendlyDisplayFormat="Disabled" ID="{5ee2dd25-d941-455a-9bdb-7f2c54aed11b}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertStartDateTime" Name="PnPAlertStartDateTime"><Default>[today]</Default></Field>\`'
    
    Create a URL list column
      ${this.name} --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listTitle Events --xml '\`<Field Type="URL" DisplayName="More information link" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Hyperlink" Group="PnP Columns" ID="{6085e32a-339b-4da7-ab6d-c1e013e5ab27}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertMoreInformation" Name="PnPAlertMoreInformation"></Field>\`'

    Create a URL list column and add it to all content types
      ${this.name} --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listTitle Events --xml '\`<Field Type="URL" DisplayName="More information link" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="Hyperlink" Group="PnP Columns" ID="{6085e32a-339b-4da7-ab6d-c1e013e5ab27}" SourceID="{4f118c69-66e0-497c-96ff-d7855ce0713d}" StaticName="PnPAlertMoreInformation" Name="PnPAlertMoreInformation"></Field>\`' --options AddToAllContentTypes

  More information:

    AddFieldOptions enumeration
      https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.addfieldoptions.aspx
`);
  }
}

module.exports = new SpoFieldAddCommand();