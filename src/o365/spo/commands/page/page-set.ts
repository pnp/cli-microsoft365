import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';
import { Auth } from '../../../../Auth';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  webUrl: string;
  layoutType?: string;
  promoteAs?: string;
  commentsEnabled?: string;
  publish: boolean;
  publishMessage?: string;
}

class SpoPageSetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_SET}`;
  }

  public get description(): string {
    return 'Updates modern page properties';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.layoutType = args.options.layoutType;
    telemetryProps.promoteAs = args.options.promoteAs;
    telemetryProps.commentsEnabled = args.options.commentsEnabled || false;
    telemetryProps.publish = args.options.publish || false;
    telemetryProps.publishMessage = typeof args.options.publishMessage !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';
    let requestDigest: string = '';
    let pageName: string = args.options.name;
    if (!pageName.endsWith('.aspx')) {
      pageName += '.aspx';
    }
    const serverRelativeSiteUrl: string = `${args.options.webUrl.substr(args.options.webUrl.indexOf('/', 8))}/sitepages/${pageName}`;

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        siteAccessToken = accessToken;

        if (this.verbose) {
          cmd.log(`Retrieving request digest...`);
        }

        return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise | Promise<void> => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }

        requestDigest = res.FormDigestValue;

        if (!args.options.layoutType) {
          return Promise.resolve();
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeSiteUrl}')/ListItemAllFields`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'X-RequestDigest': requestDigest,
            'content-type': 'application/json;odata=nometadata',
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
            accept: 'application/json;odata=nometadata'
          }),
          body: {
            PageLayoutType: args.options.layoutType
          },
          json: true
        };

        if (args.options.layoutType === 'Article') {
          requestOptions.body.PromotedState = 0;
          requestOptions.body.BannerImageUrl = {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `${resource}/_layouts/15/images/sitepagethumbnail.png`
          };
        }

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: any): request.RequestPromise | Promise<void> => {
        if (!args.options.promoteAs) {
          return Promise.resolve();
        }

        const requestOptions: any = {
          json: true
        };

        switch (args.options.promoteAs) {
          case 'HomePage':
            requestOptions.url = `${args.options.webUrl}/_api/web/rootfolder`;
            requestOptions.headers = Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'X-RequestDigest': requestDigest,
              'X-HTTP-Method': 'MERGE',
              'IF-MATCH': '*',
              'content-type': 'application/json;odata=nometadata',
              accept: 'application/json;odata=nometadata'
            });
            requestOptions.body = {
              WelcomePage: `SitePages/${pageName}`
            };
            break;
          case 'NewsPage':
            requestOptions.url = `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeSiteUrl}')/ListItemAllFields`;
            requestOptions.headers = Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'X-RequestDigest': requestDigest,
              'X-HTTP-Method': 'MERGE',
              'IF-MATCH': '*',
              'content-type': 'application/json;odata=nometadata',
              accept: 'application/json;odata=nometadata'
            });
            requestOptions.body = {
              PromotedState: 2,
              FirstPublishedDate: new Date().toISOString().replace('Z', '')
            }
            break;
        }

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: any): request.RequestPromise | Promise<void> => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        if (typeof args.options.commentsEnabled === 'undefined') {
          return Promise.resolve();
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeSiteUrl}')/ListItemAllFields/SetCommentsDisabled(${args.options.commentsEnabled === 'false'})`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'X-RequestDigest': requestDigest,
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: any): request.RequestPromise | Promise<void> => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        if (!args.options.publish) {
          return Promise.resolve();
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeSiteUrl}')/Publish('${encodeURIComponent(args.options.publishMessage || '').replace(/'/g, '%39')}')`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'X-RequestDigest': requestDigest,
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          }),
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
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'Name of the page to update'
      },
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the page to update is located'
      },
      {
        option: '-l, --layoutType [layoutType]',
        description: 'Layout of the page. Allowed values Article|Home',
        autocomplete: ['Article', 'Home']
      },
      {
        option: '-p, --promoteAs [promoteAs]',
        description: 'Update the page purpose. Allowed values HomePage|NewsPage',
        autocomplete: ['HomePage', 'NewsPage']
      },
      {
        option: '--commentsEnabled [commentsEnabled]',
        description: 'Set to true, to enable comments on the page. Allowed values true|false',
        autocomplete: ['true', 'false']
      },
      {
        option: '--publish',
        description: 'Set to publish the page'
      },
      {
        option: '--publishMessage [publishMessage]',
        description: 'Message to set when publishing the page'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.name) {
        return 'Required parameter name missing';
      }

      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (args.options.layoutType &&
        args.options.layoutType !== 'Article' &&
        args.options.layoutType !== 'Home') {
        return `${args.options.layoutType} is not a valid option for layoutType. Allowed values Article|Home`;
      }

      if (args.options.promoteAs &&
        args.options.promoteAs !== 'HomePage' &&
        args.options.promoteAs !== 'NewsPage') {
        return `${args.options.promoteAs} is not a valid option for promoteAs. Allowed values HomePage|NewsPage`;
      }

      if (args.options.promoteAs === 'HomePage' && args.options.layoutType !== 'Home') {
        return 'You can only promote home pages as site home page';
      }

      if (args.options.promoteAs === 'NewsPage' && args.options.layoutType === 'Home') {
        return 'You can only promote article pages as news article';
      }

      if (typeof args.options.commentsEnabled !== 'undefined' &&
        args.options.commentsEnabled !== 'true' &&
        args.options.commentsEnabled !== 'false') {
        return `${args.options.commentsEnabled} is not a valid value for commentsEnabled. Allowed values true|false`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site using the
    ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To update a modern page, you have to first connect to a SharePoint site using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    If you try to create a page with a name of a page that already exists, you
    will get a ${chalk.grey('The file doesn\'t exists')} error.

    If you choose to promote the page using the ${chalk.blue('promoteAs')} option
    or enable page comments, you will see the result only after publishing
    the page.

  Examples:

    Change the layout of the existing page to Article
      ${chalk.grey(config.delimiter)} ${this.name} --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --layoutType Article

    Promote the existing article page as a news article
      ${chalk.grey(config.delimiter)} ${this.name} --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --promoteAs NewsPage

    Change the page's layout to Home and set it as the site's home page
      ${chalk.grey(config.delimiter)} ${this.name} --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --layoutType Home --promoteAs HomePage

    Enable comments on the existing page
      ${chalk.grey(config.delimiter)} ${this.name} --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --commentsEnabled true

    Publish existing page
      ${chalk.grey(config.delimiter)} ${this.name} --name page.aspx --webUrl https://contoso.sharepoint.com/sites/a-team --publish
`);
  }
}

module.exports = new SpoPageSetCommand();