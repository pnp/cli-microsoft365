import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { Finding } from '../Finding';
import { Project, TsFile } from '../../model';
import { FN016004_TS_property_pane_property_import } from './FN016004_TS_property_pane_property_import';
import Utils from '../../../../../../Utils';
import { TsRule } from './TsRule';

describe('FN016004_TS_property_pane_property_import', () => {
  let findings: Finding[];
  let rule: FN016004_TS_property_pane_property_import;
  afterEach(() => {
    Utils.restore([
      fs.existsSync,
      fs.readFileSync,
      (TsRule as any).getParentOfType
    ]);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN016004_TS_property_pane_property_import();
  });

  it('returns empty resolution by default', () => {
    assert.strictEqual(rule.resolution, '');
  });

  it('doesn\'t return notifications if no .ts files found', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if specified .ts file not found', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp',
      tsFiles: [
        new TsFile('foo')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if @microsoft/sp-webpart-base import has values that are suppose to be there', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `import {
      BaseClientSideWebPart
    } from '@microsoft/sp-webpart-base';`);
    const project: Project = {
      path: '/usr/tmp',
      tsFiles: [
        new TsFile('foo')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if @microsoft/sp-webpart-base import has values that are not suppose to be there', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `import {
      BaseClientSideWebPart,
      IPropertyPaneConfiguration,
      PropertyPaneTextField,
      PropertyPaneLabel
    } from '@microsoft/sp-webpart-base';`);
    const project: Project = {
      path: '/usr/tmp',
      tsFiles: [
        new TsFile('foo')
      ]
    };
    rule.visit(project, findings);
    assert(findings[0].occurrences[0].resolution.indexOf('import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneLabel } from "@microsoft/sp-property-pane";') > -1);
  });

  it('does not return an empty import when all imports are moved to @microsoft/sp-property-pane', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `import {
      IPropertyPaneField,
      PropertyPaneFieldType
    } from '@microsoft/sp-webpart-base';`);
    const project: Project = {
      path: '/usr/tmp',
      tsFiles: [
        new TsFile('foo')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings[0].occurrences[0].resolution, 'import { IPropertyPaneField, PropertyPaneFieldType } from "@microsoft/sp-property-pane";');
  });

  it('does not add PropertyPaneCustomField when it is not used', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `import { IPropertyPaneCustomFieldProps } from '@microsoft/sp-webpart-base';`);
    const project: Project = {
      path: '/usr/tmp',
      tsFiles: [
        new TsFile('foo')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings[0].occurrences[0].resolution, 'import { IPropertyPaneCustomFieldProps } from "@microsoft/sp-property-pane";');
  });
});