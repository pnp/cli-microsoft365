import * as sinon from 'sinon';
import * as assert from 'assert';
import Utils from '../Utils';
import { FileTokenStorage, TokensFile } from './FileTokenStorage';
import * as fs from 'fs';

describe('FileTokenStorage', () => {
  const fileStorage = new FileTokenStorage();

  afterEach(() => {
    Utils.restore([
      fs.existsSync,
      fs.readFileSync,
      fs.writeFile
    ]);
  })

  it('fails retrieving password from file if the token file doesn\'t exist', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    fileStorage
      .get('mock')
      .then(() => {
        done('Expected fail but passed instead');
      }, (err) => {
        try {
          assert.equal(err, 'File not found');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('fails retrieving password from file if the token file is empty', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '');
    fileStorage
      .get('mock')
      .then(() => {
        done('Expected fail but passed instead');
      }, (err) => {
        done();
      });
  });

  it('fails retrieving password from file if the token file contains invalid JSON string', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '');
    fileStorage
      .get('mock')
      .then(() => {
        done('Expected fail but passed instead');
      }, (err) => {
        done();
      });
  });

  it('fails retrieving password from file if the token file contains an empty object', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '{}');
    fileStorage
      .get('mock')
      .then(() => {
        done('Expected fail but passed instead');
      }, (err) => {
        try {
          assert.equal(err, 'Token for service mock not found');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('fails retrieving password from file if the token file doesn\'t contain any passwords', (done) => {
    const tokensFile: TokensFile = {
      services: {}
    };
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify(tokensFile));
    fileStorage
      .get('mock')
      .then(() => {
        done('Expected fail but passed instead');
      }, (err) => {
        try {
          assert.equal(err, 'Token for service mock not found');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('fails retrieving password from file if the token file doesn\'t contain password for the specified service', (done) => {
    const tokensFile: TokensFile = {
      services: {
        'SPO': 'abc'
      }
    };
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify(tokensFile));
    fileStorage
      .get('mock')
      .then(() => {
        done('Expected fail but passed instead');
      }, (err) => {
        try {
          assert.equal(err, 'Token for service mock not found');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('returns password from file when the token file contains password for the specified service', (done) => {
    const tokensFile: TokensFile = {
      services: {
        'mock': 'abc'
      }
    };
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify(tokensFile));
    fileStorage
      .get('mock')
      .then((password) => {
        try {
          assert.equal(password, 'abc');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('saves the password in the file when the file doesn\'t exist', (done) => {
    const expected: TokensFile = {
      services: {
        'mock': 'abc'
      }
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    sinon.stub(fs, 'writeFile').callsFake((path, token) => { actual = token; }).callsArgWith(3, null);
    fileStorage
      .set('mock', 'abc')
      .then(() => {
        try {
          assert.equal(actual, JSON.stringify(expected));
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('saves the password in the file when the file is empty', (done) => {
    const expected: TokensFile = {
      services: {
        'mock': 'abc'
      }
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '');
    sinon.stub(fs, 'writeFile').callsFake((path, token) => { actual = token; }).callsArgWith(3, null);
    fileStorage
      .set('mock', 'abc')
      .then(() => {
        try {
          assert.equal(actual, JSON.stringify(expected));
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('saves the password in the file when the file contains an empty JSON object', (done) => {
    const expected: TokensFile = {
      services: {
        'mock': 'abc'
      }
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '{}');
    sinon.stub(fs, 'writeFile').callsFake((path, token) => { actual = token; }).callsArgWith(3, null);
    fileStorage
      .set('mock', 'abc')
      .then(() => {
        try {
          assert.equal(actual, JSON.stringify(expected));
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('saves the password in the file when the file contains no passwords', (done) => {
    const expected: TokensFile = {
      services: {
        'mock': 'abc'
      }
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '{services:{}}');
    sinon.stub(fs, 'writeFile').callsFake((path, token) => { actual = token; }).callsArgWith(3, null);
    fileStorage
      .set('mock', 'abc')
      .then(() => {
        try {
          assert.equal(actual, JSON.stringify(expected));
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('adds the password to the file when the file contains passwords', (done) => {
    const expected: TokensFile = {
      services: {
        'SPO': 'def',
        'mock': 'abc'
      }
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({ services: { 'SPO': 'def' } }));
    sinon.stub(fs, 'writeFile').callsFake((path, token) => { actual = token; }).callsArgWith(3, null);
    fileStorage
      .set('mock', 'abc')
      .then(() => {
        try {
          assert.equal(actual, JSON.stringify(expected));
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('correctly handles error when writing to the file failed', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    sinon.stub(fs, 'writeFile').callsFake(() => { }).callsArgWith(3, { message: 'An error has occurred' });
    fileStorage
      .set('mock', 'abc')
      .then(() => {
        done('Fail expected but passed instead');
      }, (err) => {
        try {
          assert.equal(err, 'An error has occurred');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('succeeds with removing if the token file doesn\'t exist', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    fileStorage
      .remove('mock')
      .then(() => {
        done();
      }, (err) => {
        done('Pass expected but failed instead');
      });
  });

  it('succeeds with removing if the token file is empty', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '');
    fileStorage
      .remove('mock')
      .then(() => {
        done();
      }, (err) => {
        done('Pass expected but failed instead');
      });
  });

  it('succeeds with removing if the token file contains empty JSON object', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => '{}');
    fileStorage
      .remove('mock')
      .then(() => {
        done();
      }, (err) => {
        done('Pass expected but failed instead');
      });
  });

  it('succeeds with removing if the token file contains no services', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({ services: {} }));
    sinon.stub(fs, 'writeFile').callsFake(() => { }).callsArgWith(3, null);
    fileStorage
      .remove('mock')
      .then(() => {
        done();
      }, (err) => {
        done('Pass expected but failed instead');
      });
  });

  it('removes password for the specified service from the token file keeping other passwords intact', (done) => {
    const expected: TokensFile = {
      services: {
        'SPO': 'abc'
      }
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      services: {
        'SPO': 'abc',
        'mock': 'def'
      }
    }));
    sinon.stub(fs, 'writeFile').callsFake((path, tokens) => { actual = tokens; }).callsArgWith(3, null);
    fileStorage
      .remove('mock')
      .then(() => {
        try {
          assert.equal(actual, JSON.stringify(expected));
          done();
        }
        catch (e) {
          done(e);
        }
      }, (err) => {
        done('Pass expected but failed instead');
      });
  });

  it('removes password for the specified service from the token file', (done) => {
    const expected: TokensFile = {
      services: {}
    };
    let actual: string = '';
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      services: {
        'mock': 'def'
      }
    }));
    sinon.stub(fs, 'writeFile').callsFake((path, tokens) => { actual = tokens; }).callsArgWith(3, null);
    fileStorage
      .remove('mock')
      .then(() => {
        try {
          assert.equal(actual, JSON.stringify(expected));
          done();
        }
        catch (e) {
          done(e);
        }
      }, (err) => {
        done('Pass expected but failed instead');
      });
  });

  it('succeeds when password successfully removed from the token file', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      services: {
        'mock': 'def'
      }
    }));
    sinon.stub(fs, 'writeFile').callsFake(() => {}).callsArgWith(3, null);
    fileStorage
      .remove('mock')
      .then(() => {
        try {
          done();
        }
        catch (e) {
          done(e);
        }
      }, (err) => {
        done('Pass expected but failed instead');
      });
  });

  it('correctly handles error when writing updated tokens to the token file', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      services: {
        'mock': 'def'
      }
    }));
    sinon.stub(fs, 'writeFile').callsFake(() => {}).callsArgWith(3, { message: 'An error has occurred' });
    fileStorage
      .remove('mock')
      .then(() => {
        done('Fail expected but passed instead');
      }, (err) => {
        try {
          assert.equal(err, 'An error has occurred');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });
});