import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { cli } from '../../../../cli/cli.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { Page } from './Page.js';
import { ClientSidePageProperties } from './ClientSidePageProperties.js';
import { PageControl } from './PageControl.js';
import { ClientSideControl } from './ClientSideControl.js';
import { urlUtil } from '../../../../utils/urlUtil.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint URL.`
    })
    .alias('u'),
  pageName: z.string().alias('n'),
  id: z.uuid().alias('i'),
  draft: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoPageControlRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_CONTROL_REMOVE;
  }

  public get description(): string {
    return 'Removes a control from a modern page';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (!args.options.force) {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to delete control '${args.options.id}' on page '${args.options.pageName}'?` });

      if (!result) {
        return;
      }
    }

    try {
      if (this.verbose) {
        await logger.logToStderr(`Getting page properties for page '${args.options.pageName}'...`);
      }

      const pageName = urlUtil.removeLeadingSlashes(args.options.pageName.toLowerCase().endsWith('.aspx') ? args.options.pageName : `${args.options.pageName}.aspx`);
      let requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/SitePages/Pages/GetByUrl('SitePages/${formatting.encodeQueryParameter(pageName)}')?$select=CanvasContent1`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const pageProps = await request.get<ClientSidePageProperties>(requestOptions);
      if (!pageProps.CanvasContent1) {
        throw `Page '${pageName}' doesn't contain canvas control '${args.options.id}'.`;
      }

      const pageControls: PageControl[] = JSON.parse(pageProps.CanvasContent1);
      const hasControl = pageControls.some(control => control.id?.toLowerCase() === args.options.id.toLowerCase());

      if (!hasControl) {
        throw `Control with ID '${args.options.id}' was not found on page '${pageName}'.`;
      }

      if (this.verbose) {
        await logger.logToStderr('Checking out page...');
      }
      const page = await Page.checkout(pageName, args.options.webUrl, logger, this.verbose);
      const canvasContent: ClientSideControl[] = JSON.parse(page.CanvasContent1);

      if (this.verbose) {
        await logger.logToStderr(`Removing control with ID '${args.options.id}' from page...`);
      }

      const pageContent = canvasContent.filter(control => !control.id || control.id.toLowerCase() !== args.options.id.toLowerCase());

      requestOptions = {
        url: `${args.options.webUrl}/_api/SitePages/Pages/GetByUrl('SitePages/${formatting.encodeQueryParameter(pageName)}')/SavePageAsDraft`,
        headers: {
          'content-type': 'application/json;odata=nometadata',
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          CanvasContent1: JSON.stringify(pageContent)
        }
      };
      await request.patch(requestOptions);

      if (!args.options.draft) {
        if (this.verbose) {
          await logger.logToStderr(`Republishing page...`);
        }

        await Page.publishPage(args.options.webUrl, pageName);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageControlRemoveCommand();