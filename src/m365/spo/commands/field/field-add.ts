import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ContextInfo } from '../../spo';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getRequestDigest(args.options.webUrl)
      .then((res: ContextInfo): Promise<{}> => {
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/${(args.options.listTitle ? `lists/getByTitle('${encodeURIComponent(args.options.listTitle)}')/` : '')}fields/CreateFieldAsXml`,
          headers: {
            'X-RequestDigest': res.FormDigestValue,
            accept: 'application/json;odata=nometadata'
          },
          data: {
            parameters: {
              SchemaXml: args.options.xml,
              Options: this.getOptions(args.options.options)
            }
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        logger.log(res);

        if (this.verbose) {
          logger.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
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

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
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
  }
}

module.exports = new SpoFieldAddCommand();