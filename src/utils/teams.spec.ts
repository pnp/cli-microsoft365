import * as assert from 'assert';
import * as sinon from 'sinon';
import request from "../request";
import auth from '../Auth';
import { teams } from './teams';
import { sinonUtil } from "./sinonUtil";
import { formatting } from './formatting';
import { aadGroup } from './aadGroup';
import { Logger } from '../cli/Logger';

describe('utils/teams', () => {
  let logger: Logger;
  let log: string[];

  const teamsResponse = {
    id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844',
    createdDateTime: '2017-11-29T03:27:05Z',
    displayName: 'Finance',
    description: 'This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.',
    classification: null,
    specialization: 'none',
    visibility: 'Public',
    webUrl: 'https://teams.microsoft.com/l/team/19:ASjdflg-xKFnjueOwbm3es6HF2zx3Ki57MyfDFrjeg01%40thread.tacv2/conversations?groupId=1caf7dcd-7e83-4c3a-94f7-932a1299c844&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35',
    isArchived: false,
    isMembershipLimitedToOwners: false,
    discoverySettings: {
      showInTeamsSearchAndSuggestions: false
    },
    memberSettings: {
      allowCreateUpdateChannels: true,
      allowCreatePrivateChannels: true,
      allowDeleteChannels: true,
      allowAddRemoveApps: true,
      allowCreateUpdateRemoveTabs: true,
      allowCreateUpdateRemoveConnectors: true
    },
    guestSettings: {
      allowCreateUpdateChannels: false,
      allowDeleteChannels: false
    },
    messagingSettings: {
      allowUserEditMessages: true,
      allowUserDeleteMessages: true,
      allowOwnerDeleteMessages: true,
      allowTeamMentions: true,
      allowChannelMentions: true
    },
    funSettings: {
      allowGiphy: true,
      giphyContentRating: 'moderate',
      allowStickersAndMemes: true,
      allowCustomMemes: true
    }
  };

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
  });

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    auth.service.connected = true;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('gets the specific team by name correct dynamics url as admin', async () => {
    const groupId = '00000000-0000-0000-0000-000000000000';

    sinon.stub(aadGroup, 'getGroupIdByDisplayName').resolves(groupId);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter(groupId)}`)) {
        return teamsResponse;
      }

      throw 'Invalid request';
    });

    const actual = await teams.getTeamByName('Teams name', logger, true);
    assert.strictEqual(actual, teamsResponse);
  });
});