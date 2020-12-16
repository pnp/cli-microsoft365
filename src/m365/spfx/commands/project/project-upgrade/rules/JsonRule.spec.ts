import * as assert from 'assert';
import * as parse from 'json-to-ast';
import { JsonFile, Project } from '../../model';
import { Finding } from '../Finding';
import { JsonRule } from './JsonRule';

class MockJsonRule extends JsonRule {
  get id(): string {
    return 'FN000000';
  }

  get title(): string {
    return 'Mock rule';
  }

  get description(): string {
    return 'Mock JSON rule';
  }

  get severity(): string {
    return 'Required';
  }

  visit(project: Project, findings: Finding[]): void {
  }

  public getAstNode(jsonFile: JsonFile, jsonProperty: string): parse.ASTNode | undefined {
    return this.getAstNodeFromFile(jsonFile, jsonProperty);
  }
}

describe('JsonRule', () => {
  let rule: MockJsonRule;

  beforeEach(() => {
    rule = new MockJsonRule();
  })

  it('has empty resolution', () => {
    assert.strictEqual('', rule.resolution);
  });

  it('has resolution type set to json', () => {
    assert.strictEqual('json', rule.resolutionType);
  });

  it('has empty file', () => {
    assert.strictEqual('', rule.file);
  });

  it('returns the root node when the node has not children', () => {
    const jsonFile: JsonFile = { source: JSON.stringify({}) };
    jsonFile.ast = parse(jsonFile.source!);
    assert.strictEqual(rule.getAstNode(jsonFile, ''), jsonFile.ast);
  });

  it('returns the root node when no property specified', () => {
    const jsonFile: JsonFile = { source: JSON.stringify({ prop: 'value' }, null, 2) };
    jsonFile.ast = parse(jsonFile.source!);
    assert.strictEqual(rule.getAstNode(jsonFile, ''), jsonFile.ast);
  });

  it('returns correct line number for the specified root property', () => {
    const jsonFile: JsonFile = { source: JSON.stringify({ prop: 'value' }, null, 2) };
    const node = rule.getAstNode(jsonFile, 'prop');
    assert.notStrictEqual(typeof node, 'undefined', 'Node not found');
    assert.strictEqual(node?.loc?.start.line, 2, 'Incorrect line number');
  });

  it('returns correct line number for the specified child property', () => {
    const jsonFile: JsonFile = { source: JSON.stringify({
      prop: {
        child: 'value'
      }
    }, null, 2) };
    const node = rule.getAstNode(jsonFile, 'prop.child');
    assert.notStrictEqual(typeof node, 'undefined', 'Node not found');
    assert.strictEqual(node?.loc?.start.line, 3, 'Incorrect line number');
  });

  it('returns root property node if the child node not found', () => {
    const jsonFile: JsonFile = { source: JSON.stringify({
      prop: {
        child: 'value'
      }
    }, null, 2) };
    const node = rule.getAstNode(jsonFile, 'prop.child1');
    assert.notStrictEqual(typeof node, 'undefined', 'Node not found');
    assert.strictEqual(node?.loc?.start.line, 2, 'Incorrect line number');
  });

  it('returns correct line number for a string array value', () => {
    const jsonFile: JsonFile = { source: JSON.stringify({
      prop: [
        'value1',
        'value2'
      ]
    }, null, 2) };
    const node = rule.getAstNode(jsonFile, 'prop[value2]');
    assert.notStrictEqual(typeof node, 'undefined', 'Node not found');
    assert.strictEqual(node?.loc?.start.line, 4, 'Incorrect line number');
  });

  it('returns correct line number for a specified object from an array', () => {
    const jsonFile: JsonFile = { source: JSON.stringify({
      prop: [
        {},
        {}
      ]
    }, null, 2) };
    const node = rule.getAstNode(jsonFile, 'prop[1]');
    assert.notStrictEqual(typeof node, 'undefined', 'Node not found');
    assert.strictEqual(node?.loc?.start.line, 4, 'Incorrect line number');
  });

  it('returns correct line number for the child property of the specified object from an array', () => {
    const jsonFile: JsonFile = { source: JSON.stringify({
      prop: [
        {},
        {
          child: 'value'
        }
      ]
    }, null, 2) };
    const node = rule.getAstNode(jsonFile, 'prop[1].child');
    assert.notStrictEqual(typeof node, 'undefined', 'Node not found');
    assert.strictEqual(node?.loc?.start.line, 5, 'Incorrect line number');
  });

  it('returns line number of the parent node if a child property of the specified object from an array not found', () => {
    const jsonFile: JsonFile = { source: JSON.stringify({
      prop: [
        {},
        {
          child: 'value'
        }
      ]
    }, null, 2) };
    const node = rule.getAstNode(jsonFile, 'prop[1].child1');
    assert.notStrictEqual(typeof node, 'undefined', 'Node not found');
    assert.strictEqual(node?.loc?.start.line, 4, 'Incorrect line number');
  });
});