import * as sinon from 'sinon';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { Page } from './Page';
import { ClientSidePage } from './clientsidepages';

describe('Page', () => {
  let log: string[];
  let cmdInstance: any;

  beforeEach(() => {
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  it('correctly handles error when parsing modern page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.resolve({
        ListItemAllFields: {
          ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
          CanvasContent1: '<div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSetti"></div></div>'
        }
      });
    });

    Page
      .getPage('page.aspx', 'https://contoso.sharepoint.com', 'abc', cmdInstance, false, false)
      .then((page: ClientSidePage): void => {
        done(new Error('Parsing page didn\'t fail while expected'));
      }, (error: any): void => {
        done();
      });
  });
});