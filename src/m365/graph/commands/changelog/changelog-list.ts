import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils';
import AnonymousCommand from '../../../base/AnonymousCommand';
import { Changelog, ChangelogItem } from '../../Changelog';
import commands from '../../commands';
import * as Parser from 'rss-parser';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  versions?: string;
  changeType?: string;
  services?: string;
  startDate?: string;
  endDate?: string;
}

class GraphChangelogListCommand extends AnonymousCommand {
  private allowedVersions: string[] = ['beta', 'v1.0'];
  private allowedChangeTypes: string[] = ['Addition', 'Change', 'Deletion', 'Deprecation'];
  private allowedServices: string[] = [
    'Applications', 'Calendar', 'Change notifications', 'Cloud communications', 
    'Compliance', 'Cross-device experiences', 'Customer booking', 'Device and app management', 
    'Education', 'Files', 'Financials', 'Groups', 
    'Identity and access', 'Mail', 'Notes', 'Notifications', 
    'People and workplace intelligence', 'Personal contacts', 'Reports', 'Search', 
    'Security', 'Sites and lists', 'Tasks and plans', 'Teamwork', 
    'To-do tasks', 'Users', 'Workbooks and charts'
  ];

  public get name(): string {
    return commands.CHANGELOG_LIST;
  }

  public get description(): string {
    return 'Gets an overview of specific API-level changes in Microsoft Graph v1.0 and beta';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.versions = typeof args.options.versions !== 'undefined';
    telemetryProps.changeType = typeof args.options.changeType !== 'undefined';
    telemetryProps.services = typeof args.options.services !== 'undefined';
    telemetryProps.startDate = typeof args.options.startDate !== 'undefined';
    telemetryProps.endDate = typeof args.options.endDate !== 'undefined';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['guid', 'category', 'title', 'description', 'pubDate'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const parser = new Parser();
    const allowedChangeType = args.options.changeType && this.allowedChangeTypes.find(x => x.toLocaleLowerCase() === args.options.changeType!.toLocaleLowerCase());
    const searchParam = args.options.changeType ? `/?filterBy=${allowedChangeType}` : '';

    parser
      .parseURL(`https://developer.microsoft.com/en-us/graph/changelog/rss${searchParam}`)
      .then((output: any) => {
        const changelog: Changelog = this.filterThroughOptions(args.options, this.mapChangelog(output));
        
        logger.log(changelog.items);
        cb();
      });
  }

  private filterThroughOptions(options: Options, changelog: Changelog): Changelog {
    let items: ChangelogItem[] = changelog.items;

    if (options.services) {
      const allowedServices: string[] = this.allowedServices
        .filter(allowedService => options.services!.toLocaleLowerCase().split(',').includes(allowedService.toLocaleLowerCase()));

      items = changelog.items.filter(item => allowedServices.includes(item.title));
    }

    if (options.versions) {
      const allowedVersions: string[] = this.allowedVersions
        .filter(allowedVersion => options.versions!.toLocaleLowerCase().split(',').includes(allowedVersion.toLocaleLowerCase()));

      items = items.filter(item => allowedVersions.includes(item.category));
    }

    if (options.startDate) {
      const startDate: Date = new Date(options.startDate);

      items = items.filter(item => item.pubDate >= startDate);
    }

    if (options.endDate) {
      const endDate: Date = new Date(options.endDate);

      items = items.filter(item => item.pubDate <= endDate);
    }

    // Make sure everything is unique based on the item guid
    items = [...new Map(items.map((item) => [item.guid, item])).values()];

    // Order items by date desc
    changelog.items = items.sort((itemA, itemB) => Number(itemB.pubDate) - Number(itemA.pubDate));

    return changelog;
  }

  private mapChangelog(output: any): Changelog {
    const changelogItems: ChangelogItem[] = output['items'].map((item: any) => ({
      guid: item['guid'],
      category: item['categories'][1],
      title: item['title'],
      description: item['contentSnippet'],
      pubDate: new Date(item['isoDate'])
    }) as ChangelogItem);

    return ({
      title: output['title'],
      url: output['link'],
      description: output['description'],
      items: changelogItems
    }) as Changelog;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-v, --versions [versions]',
        autocomplete: this.allowedVersions
      },
      {
        option: "-c, --changeType [changeType]",
        autocomplete: this.allowedChangeTypes
      },
      {
        option: "-s, --services [services]",
        autocomplete: this.allowedServices
      },
      {
        option: "--startDate [startDate]"
      },
      {
        option: "--endDate [endDate]"
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (
      args.options.versions && 
      args.options.versions.toLocaleLowerCase().split(',').some(x => !this.allowedVersions.map(y => y.toLocaleLowerCase()).includes(x))) {
      return `The verions contains an invalid value. Specify either ${this.allowedVersions.join(', ')} as properties`;
    }

    if (
      args.options.changeType && 
      !this.allowedChangeTypes.map(x => x.toLocaleLowerCase()).includes(args.options.changeType.toLocaleLowerCase())) {
      return `The change type contain an invalid value. Specify either ${this.allowedChangeTypes.join(', ')} as properties`;
    }

    if (
      args.options.services && 
      args.options.services.toLocaleLowerCase().split(',').some(x => !this.allowedServices.map(y => y.toLocaleLowerCase()).includes(x))) {
      return `The services contains invalid value. Specify either ${this.allowedServices.join(', ')} as properties`;
    }

    if (args.options.startDate && !validation.isValidISODate(args.options.startDate)) {
      return 'The startDate is not a valid ISO date string';
    }

    if (args.options.endDate && !validation.isValidISODate(args.options.endDate)) {
      return 'The endDate is not a valid ISO date string';
    }

    return true;
  }
}

module.exports = new GraphChangelogListCommand();