import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  id: z.int().positive(),
  title: z.string().min(1, 'Cannot be empty.').optional(),
  url: z.string().optional(),
  audienceIds: z.string()
    .refine(audienceIds => audienceIds === '' || audienceIds.split(',').length <= 10, {
      error: 'The maximum amount of audienceIds per navigation node exceeded. The maximum amount of audienceIds is 10.'
    })
    .refine(audienceIds => audienceIds === '' || validation.isValidGuidArray(audienceIds) === true, {
      error: e => `The following GUIDs are invalid for the option 'audienceIds': ${validation.isValidGuidArray(e.input as string)}.`
    })
    .optional(),
  openInNewWindow: z.boolean().optional()
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoNavigationNodeSetCommand extends SpoCommand {
  public get name(): string {
    return commands.NAVIGATION_NODE_SET;
  }

  public get description(): string {
    return 'Updates a SharePoint navigation node';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.title, options.url, options.audienceIds, options.openInNewWindow].some(o => o !== undefined), {
        error: 'Specify at least one property to update.'
      });
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const { webUrl, id, title, url, audienceIds, openInNewWindow } = args.options;

    if (this.verbose) {
      await logger.logToStderr(`Updating navigation node with id ${id}...`);
    }

    try {
      const { menuState, node } = await spo.getMenuStateNodeByKey(webUrl, id.toString());

      if (title !== undefined) {
        node.Title = title;
      }

      if (url !== undefined) {
        node.SimpleUrl = url;
      }

      if (audienceIds !== undefined) {
        node.AudienceIds = audienceIds === '' ? [] : audienceIds.split(',');
      }

      if (openInNewWindow !== undefined) {
        node.OpenInNewWindow = openInNewWindow;
      }

      await spo.saveMenuState(webUrl, menuState);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoNavigationNodeSetCommand();