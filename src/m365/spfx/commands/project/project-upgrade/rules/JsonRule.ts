import * as parse from 'json-to-ast';
import { Finding } from "../";
import { JsonFile } from '../../model';
import { OccurrencePosition } from '../Occurrence';
import { Rule } from "./Rule";

export abstract class JsonRule extends Rule {
  get resolution(): string {
    return '';
  }

  get resolutionType(): string {
    return 'json';
  }

  get file(): string {
    return '';
  }

  protected addFindingWithPosition(findings: Finding[], node: parse.ASTNode | undefined): void {
    this.addFindingWithOccurrences([{
      file: this.file,
      resolution: this.resolution,
      position: this.getPositionFromNode(node)
    }], findings);
  }

  protected getPositionFromNode(node: parse.ASTNode | undefined): OccurrencePosition {
    if (!node || !node.loc) {
      return { line: 1, character: 1 };
    }

    return {
      line: node.loc.start.line,
      character: node.loc.start.column
    };
  }

  protected getAstNodeFromFile(jsonFile: JsonFile, jsonProperty: string): parse.ASTNode | undefined {
    if (!jsonFile.source) {
      return undefined;
    }

    if (!jsonFile.ast) {
      jsonFile.ast = parse(jsonFile.source);
    }

    return this.getAstNodeForProperty(jsonFile.ast as parse.ArrayNode, jsonProperty);
  }

  private getAstNodeForProperty(node: parse.ArrayNode, jsonProperty: string): parse.ASTNode | undefined {
    if (node.children.length === 0) {
      return node;
    }

    if (jsonProperty === '') {
      return node;
    }

    const jsonPropertyChunks = jsonProperty.split('.');
    let currentProperty = jsonPropertyChunks[0];
    currentProperty = currentProperty.replace(/;#/g, '.');
    let isArray = false;
    let arrayElement: string | undefined;
    if (currentProperty.endsWith(']')) {
      isArray = true;
      const pos = currentProperty.indexOf('[') + 1;
      // get array element from the property name
      arrayElement = currentProperty.substr(pos, currentProperty.length - pos - 1);
      // remove array element from the property name
      currentProperty = currentProperty.substr(0, pos - 1);
    }

    for (let i = 0; i < node.children.length; i++) {
      let currentNode: parse.PropertyNode = node.children[i] as unknown as parse.PropertyNode;
      
      if (currentNode.key.value !== currentProperty) {
        continue;
      }

      if (isArray) {
        const arrayIndex = parseInt(arrayElement as string);
        const arrayElements = (currentNode.value as parse.ArrayNode).children;

        if (isNaN(arrayIndex)) {
          for (let j = 0; j < arrayElements.length; j++) {
            if ((arrayElements[j] as parse.LiteralNode).value === arrayElement) {
              currentNode = arrayElements[j] as unknown as parse.PropertyNode;
              break;
            }
          }
        }
        else {
          if (arrayIndex < arrayElements.length) {
            currentNode = arrayElements[arrayIndex] as unknown as parse.PropertyNode;
          }
        }
      }

      // if this is the last chunk, return current node
      if (jsonPropertyChunks.length === 1) {
        return currentNode;
      }

      // more chunks left, remove current from the array, and look for child nodes
      jsonPropertyChunks.splice(0, 1);
      return this.getAstNodeForProperty((isArray ? currentNode : currentNode.value) as unknown as parse.ArrayNode, jsonPropertyChunks.join('.'));
    }

    return node;
  }
}