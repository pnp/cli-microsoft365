import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { formatting } from '../../../../utils/formatting';
const command: Command = require('./file-sharinglink-set');

describe(commands.FILE_SHARINGLINK_SET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const webUrl = 'https://contoso.sharepoint.com';
  const listId = 'eb15c4c1-6820-4462-aff9-87e1fb98590b';
  const fileId = 'f09c4efe-b8c0-4e89-a166-03418661b89b';
  const sharingLinkId = '7c9f97c9-1bda-433c-9364-bb83e81771ee';
  const fileUrl = '/sites/project-x/documents/SharedFile.docx';
  const fileDetailsResponse = {
    ListId: listId,
    UniqueId: fileId
  };

  const shareLinkResponseText = {
    id: '7c9f97c9-1bda-433c-9364-bb83e81771ee',
    link: 'https://contoso.sharepoint.com/:b:/g/EbZx4QPyndlGp6HV-gvSPksBftmUNAiXjm0y-_527_fI9g',
    scope: 0
  };

  const sharingInformationResponse = {
    d: {
      __metadata: {
        id: 'https://contoso.sharepoint.com/_api/web/Lists(@a1)/GetItemByUniqueId(@a2)/GetSharingInformation',
        uri: 'https://contoso.sharepoint.com/_api/web/Lists(@a1)/GetItemByUniqueId(@a2)/GetSharingInformation',
        type: 'SP.Sharing.SharingInformation'
      },
      pickerSettings: { __deferred: [Object] },
      anonymousLinkExpirationRestrictionDays: -1,
      anyoneLinkTrackUsers: false,
      blockPeoplePickerAndSharing: false,
      canAddExternalPrincipal: true,
      canAddInternalPrincipal: true,
      canRequestAccessForGrantAccess: true,
      canSendEmail: true,
      canUseSimplifiedRoles: true,
      currentRole: 3,
      customizedExternalSharingServiceUrl: '',
      defaultLinkKind: 5,
      defaultShareLinkPermission: 2,
      defaultShareLinkScope: 0,
      defaultShareLinkToExistingAccess: false,
      directUrl: 'https://contoso.sharepoint.com/:b:/r/Shared%20Documents/Document.docx?csf=1&web=1',
      displayName: 'Document.docx',
      enforceIBSegmentFiltering: false,
      enforceSPOSearch: false,
      fileExtension: 'docx',
      hasUniquePermissions: true,
      isStubFile: false,
      itemUniqueId: 'f09c4efe-b8c0-4e89-a166-03418661b89b',
      itemUrl: 'https://contoso.sharepoint.com/_api/v2.0/drives/b!cLoLZkprP0qqo4PG4KBxvLNbttihbPJNpL4O_givJYCARvYucPShT4HI2ULs_paW/items/017S3RUYVWOHQQH4U53FDKPIOV7IF5EPSL',
      microserviceShareUiUrl: 'https://www.odwebp.svc.ms/share',
      outlookEndpointHostUrl: 'outlook.office.com',
      sensitivityLabelInformation: null,
      sharedObjectType: 1,
      shareUiUrl: 'https://conoso.sharepoint.com/Shared%20Documents/Forms/AllItems.aspx?p=12&id=%2FShared%20Documents%2FDocument%2Edocx&clientId=clientId',
      sharingAbilities: {
        __metadata: [Object],
        anonymousLinkAbilities: [Object],
        anyoneLinkAbilities: [Object],
        directSharingAbilities: [Object],
        organizationLinkAbilities: [Object],
        peopleSharingLinkAbilities: [Object]
      },
      sharingLinkTemplates: {
        __metadata: [Object], templates: {
          __metadata: { type: 'Collection(SP.Sharing.SharingLinkDefaultTemplate)' },
          results: [
            {
              linkDetails: {
                __metadata: { type: 'SP.SharingLinkInfo' },
                AllowsAnonymousAccess: true,
                ApplicationId: null,
                BlocksDownload: false,
                Created: '2023-01-09T20:19:55.215Z',
                CreatedBy: {
                  __metadata: [Object],
                  directoryObjectId: null,
                  email: 'john@contoso.onmicrosoft.com',
                  expiration: null,
                  id: 10,
                  isActive: true,
                  isExternal: false,
                  jobTitle: null,
                  loginName: 'i:0#.f|membership|john@contoso.onmicrosoft.com',
                  name: 'John Doe',
                  principalType: 1,
                  userId: null,
                  userPrincipalName: 'john@contoso.onmicrosoft.com'
                },
                Description: null,
                Embeddable: false,
                Expiration: '',
                HasExternalGuestInvitees: false,
                Invitations: { __metadata: [Object], results: [] },
                IsActive: true,
                IsAddressBarLink: false,
                IsCreateOnlyLink: false,
                IsDefault: true,
                IsEditLink: true,
                IsFormsLink: false,
                IsManageListLink: false,
                IsReviewLink: false,
                IsUnhealthy: false,
                LastModified: '2023-01-09T20:19:55.215Z',
                LastModifiedBy: {
                  __metadata: [Object],
                  directoryObjectId: null,
                  email: 'john@contoso.onmicrosoft.com',
                  expiration: null,
                  id: 10,
                  isActive: true,
                  isExternal: false,
                  jobTitle: null,
                  loginName: 'i:0#.f|membership|john@contoso.onmicrosoft.com',
                  name: 'John Doe',
                  principalType: 1,
                  userId: null,
                  userPrincipalName: 'john@contoso.onmicrosoft.com'
                },
                LimitUseToApplication: false,
                LinkKind: 5,
                PasswordLastModified: '',
                PasswordLastModifiedBy: null,
                RedeemedUsers: { __metadata: [Object], results: [] },
                RequiresPassword: false,
                RestrictedShareMembership: false,
                Scope: 0,
                ShareId: '7c9f97c9-1bda-433c-9364-bb83e81771ee',
                ShareTokenString: 'share=EbZx4QPyndlGp6HV-gvSPksBLEqz3k4gJxPFPm4f4tZtVA',
                SharingLinkStatus: 0,
                TrackLinkUsers: false,
                Url: 'https://contoso.sharepoint.com/:b:/g/EbZx4QPyndlGp6HV-gvSPksBLEqz3k4gJxPFPm4f4tZtVA'
              },
              passwordProtected: false,
              role: 2,
              scope: 0,
              shareKind: 5,
              trackLinkUsers: false
            },
            {}],
          linkDetails: [Object],
          passwordProtected: false,
          role: 2,
          scope: 0,
          shareKind: 5,
          trackLinkUsers: false
        }
      },
      sharingStatus: 2,
      showExternalSharingWarning: false,
      siteIBMode: 'Open',
      siteIBSegmentIDs: { __metadata: [Object], results: [] },
      standardRolesModified: true,
      userIsSharingViaMCS: null,
      userPhotoCdnBaseUrl: 'https://contoso.sharepoint.com/_vti_bin/afdcache.ashx/_userprofile/userphoto.jpg?_oat_=1673331428_28ca8e187ff361085b115bc79aa0ad8cb06fe6f6d81cf4ed5e6622ec67707075&P1=1673303395&P2=864232212&P3=1&P4=ln0ppMcWTpAaN6GRFeFG4zIfkAYGDEP%2b%2fndebReEd5LPqTjr62Hgubd0Q7W%2fUudwaqWIhiWjdAO5dNl8jLgvPROXRk4yZAaneMd0%2ffRubh5CLc6jVD7gEozoxslVD0xQnYN1Ol4P1eUjnO1TjoZEN5MV%2b8dBPlCGYWmn8YWNGPpfWal9rWzfzSNRhPhM6ndgxPIzY2DVfSU%2bfMqPfEEKjqewPul8RgAbHh9T%2f7HrZDoNuA53CPD3oujItPYsQ%2fHoNjM5IUpBFWrkt6amBtXmwUZKuJQ97kQgN1GJ0V3ysTJzwaSLG7T2kTnkiRRkWXyOhp%2fh1e6clXJl2M87Gzrmeg%3d%3d',
      webTemplateId: 68,
      webUrl: 'https://contoso.sharepoint.com'
    }
  };

  const shareLinkResponse = {
    d: {
      ShareLink: {
        __metadata: {
          type: "SP.Sharing.ShareLinkResponse"
        },
        sharingLinkInfo: {
          __metadata: {
            type: "SP.SharingLinkInfo"
          },
          AllowsAnonymousAccess: true,
          ApplicationId: null,
          BlocksDownload: false,
          Created: "2023-01-09T20:20:22.999Z",
          CreatedBy: {
            __metadata: {
              type: "SP.Sharing.Principal"
            },
            directoryObjectId: null,
            email: "john@contoso.onmicrosoft.com",
            expiration: null,
            id: 10,
            isActive: true,
            isExternal: false,
            jobTitle: null,
            loginName: "i:0#.f|membership|john@contoso.onmicrosoft.com",
            name: "John Doe",
            principalType: 1,
            userId: null,
            userPrincipalName: "john@contoso.onmicrosoft.com"
          },
          Description: null,
          Embeddable: false,
          Expiration: "2023-10-31T23:00:00.000Z",
          HasExternalGuestInvitees: false,
          Invitations: {
            __metadata: {
              type: "Collection(SP.Sharing.LinkInvitation)"
            },
            results: []
          },
          IsActive: true,
          IsAddressBarLink: false,
          IsCreateOnlyLink: false,
          IsDefault: true,
          IsEditLink: false,
          IsFormsLink: false,
          IsManageListLink: false,
          IsReviewLink: false,
          IsUnhealthy: false,
          LastModified: "2023-01-09T21:14:53.181Z",
          LastModifiedBy: {
            __metadata: {
              type: "SP.Sharing.Principal"
            },
            directoryObjectId: null,
            email: "john@contoso.onmicrosoft.com",
            expiration: null,
            id: 10,
            isActive: true,
            isExternal: false,
            jobTitle: null,
            loginName: "i:0#.f|membership|john@contoso.onmicrosoft.com",
            name: "John Doe",
            principalType: 1,
            userId: null,
            userPrincipalName: "john@contoso.onmicrosoft.com"
          },
          LimitUseToApplication: false,
          LinkKind: 4,
          PasswordLastModified: "",
          PasswordLastModifiedBy: null,
          RedeemedUsers: {
            __metadata: {
              type: "Collection(SP.Sharing.LinkInvitation)"
            },
            results: []
          },
          RequiresPassword: false,
          RestrictedShareMembership: false,
          Scope: 0,
          ShareId: "7c9f97c9-1bda-433c-9364-bb83e81771ee",
          ShareTokenString: "share=EbZx4QPyndlGp6HV-gvSPksBftmUNAiXjm0y-_527_fI9g",
          SharingLinkStatus: 2,
          TrackLinkUsers: false,
          Url: "https://contoso.sharepoint.com/:b:/g/EbZx4QPyndlGp6HV-gvSPksBftmUNAiXjm0y-_527_fI9g"
        }
      }
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_SHARINGLINK_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates a sharing link from a file specified by the id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileById('${fileId}')?$select=ListId,UniqueId`) {
        return fileDetailsResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/Lists(@a1)/GetItemByUniqueId(@a2)/GetSharingInformation?@a1='${fileDetailsResponse.ListId}'&@a2='${fileDetailsResponse.UniqueId}'&$Expand=sharingLinkTemplates`) {
        return sharingInformationResponse;
      }

      if (opts.url === `${webUrl}/_api/web/Lists(@a1)/GetItemByUniqueId(@a2)/ShareLink?@a1='${fileDetailsResponse.ListId}'&@a2='${fileDetailsResponse.UniqueId}'`) {
        return shareLinkResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, fileId: fileId, output: 'json', id: sharingLinkId, role: 'read', verbose: true } } as any);
    assert(loggerLogSpy.calledWith(shareLinkResponse.d.ShareLink.sharingLinkInfo));
  });

  it('updates a sharing link from a file specified by the url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileUrl)}')?$select=ListId,UniqueId`) {
        return fileDetailsResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/Lists(@a1)/GetItemByUniqueId(@a2)/GetSharingInformation?@a1='${fileDetailsResponse.ListId}'&@a2='${fileDetailsResponse.UniqueId}'&$Expand=sharingLinkTemplates`) {
        return sharingInformationResponse;
      }

      if (opts.url === `${webUrl}/_api/web/Lists(@a1)/GetItemByUniqueId(@a2)/ShareLink?@a1='${fileDetailsResponse.ListId}'&@a2='${fileDetailsResponse.UniqueId}'`) {
        return shareLinkResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, fileUrl: fileUrl, output: 'json', id: sharingLinkId, expirationDateTime: '2023-01-09', verbose: true } } as any);
    assert(loggerLogSpy.calledWith(shareLinkResponse.d.ShareLink.sharingLinkInfo));
  });

  it('updates a sharing link from a file specified by the id with text output', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileById('${fileId}')?$select=ListId,UniqueId`) {
        return fileDetailsResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/Lists(@a1)/GetItemByUniqueId(@a2)/GetSharingInformation?@a1='${fileDetailsResponse.ListId}'&@a2='${fileDetailsResponse.UniqueId}'&$Expand=sharingLinkTemplates`) {
        return sharingInformationResponse;
      }

      if (opts.url === `${webUrl}/_api/web/Lists(@a1)/GetItemByUniqueId(@a2)/ShareLink?@a1='${fileDetailsResponse.ListId}'&@a2='${fileDetailsResponse.UniqueId}'`) {
        return shareLinkResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, fileId: fileId, output: 'text', id: sharingLinkId, role: 'write', verbose: true } } as any);
    assert(loggerLogSpy.calledWith(shareLinkResponseText));
  });

  it('throws error when file not found by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileById('${fileId}')?$select=ListId,UniqueId`) {
        throw { error: { 'odata.error': { message: { value: 'File Not Found.' } } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, fileId: fileId, id: sharingLinkId, role: 'read', verbose: true } } as any),
      new CommandError(`File Not Found.`));
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', fileId: fileId, id: sharingLinkId, role: 'read' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the fileId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: 'invalid', id: sharingLinkId, role: 'read' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: fileId, id: 'invalid', role: 'read' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the expirationDateTime option is not a valid date', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: fileId, id: sharingLinkId, expirationDateTime: 'invalid date' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid role specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: fileId, id: sharingLinkId, role: 'invalid role' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if options are valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', fileId: fileId, id: sharingLinkId, role: 'read' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
