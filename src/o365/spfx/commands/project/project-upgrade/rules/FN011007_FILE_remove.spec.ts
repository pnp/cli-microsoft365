import * as assert from 'assert';
import * as path from 'path';
import { Finding } from '../Finding';
import { Project } from '../model';
import { FN011007_FILE_remove } from './FN011007_FILE_remove';

describe('FN011007_FILE_remove', () => {
  let findings: Finding[];
  let rule: FN011007_FILE_remove;

  beforeEach(() => {
    findings = [];
  });

  it('doesn\'t return notification file doesn\'t exist', () => {
    rule = new FN011007_FILE_remove('dummy.json');
    const project: Project = {
      path: path.join(__dirname, '../test-projects/spfx-102-webpart-react').replace('dist', 'src'),
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });

  it('returns a notification if file exists', () => {
    rule = new FN011007_FILE_remove('/typings/tsd.d.ts');
    const project: Project = {
      path: path.join(__dirname, '../test-projects/spfx-102-webpart-react').replace('dist', 'src'),
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 1);
  });
});