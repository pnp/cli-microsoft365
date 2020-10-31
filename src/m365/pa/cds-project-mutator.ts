import * as path from 'path';
import { DOMParser } from 'xmldom';

/*
 * Logic extracted from bolt.module.solution.dll
 * Version: 0.4.3
 * Class: bolt.module.solution.CdsProjectMutator
 */
export default class CdsProjectMutator {
  private _cdsProjectDocument: Document;
  private _cdsProject: HTMLElement;
  private _cdsNamespace: string;

  public get cdsProjectDocument(): Document {
    return this._cdsProjectDocument;
  }

  public constructor(document: string) {
    this._cdsProjectDocument = new DOMParser().parseFromString(document, 'text/xml');
    this._cdsProject = this._cdsProjectDocument.documentElement;
    this._cdsNamespace = this._cdsProject.lookupNamespaceURI('') || '';
  }

  public addProjectReference(referencedProjectPath: string): void {
    if (!this.doesProjectReferenceExists(referencedProjectPath)) {
      const projectReferenceElement = this.createProjectReferenceElement(referencedProjectPath);
      var projectReferenceItemGroup = this.getProjectReferenceItemGroup();
      if (projectReferenceItemGroup) {
        this.addProjectReferenceElement(projectReferenceItemGroup, projectReferenceElement);
      }
      else {
        projectReferenceItemGroup = this.createProjectReferenceItemGroup(projectReferenceElement);
        this.addProjectReferenceItemGroupElement(projectReferenceItemGroup);
      }
    }
  }

  private doesProjectReferenceExists(referencedProjectPath: string): boolean {
    return this.getNamedGroups('ItemGroup').some(itemGroup => {
      return this.getProjectReferencesFromItemGroup(itemGroup).some(projectReference => {
        const projectReferencePath = projectReference.getAttributeNode('Include');
        return (projectReferencePath && path.normalize(projectReferencePath.value).toLowerCase() === referencedProjectPath.toLowerCase());
      });
    });
  }

  private getNamedGroups(name: string): Element[] {
    return Array.from(this._cdsProject.getElementsByTagNameNS(this._cdsNamespace, name));
  }

  private getProjectReferencesFromItemGroup(itemGroup: Element): Element[] {
    return Array.from(itemGroup.getElementsByTagNameNS(this._cdsNamespace, 'ProjectReference'));
  }

  private getProjectReferenceItemGroup(): Element | null {
    const itemGroups = this.getNamedGroups('ItemGroup').filter(itemGroup => this.getProjectReferencesFromItemGroup(itemGroup).length > 0);
    return itemGroups.length > 0 ? itemGroups[0] : null;
  }

  private createProjectReferenceElement(referencedProjectPath: string): Node {
    var projectReferenceElement = this._cdsProjectDocument.createElementNS(this._cdsNamespace, 'ProjectReference');
    projectReferenceElement.setAttributeNS(this._cdsNamespace, 'Include', referencedProjectPath);
    return projectReferenceElement;
  }

  private createProjectReferenceItemGroup(projectReferenceElement: Node): Element {
    var projectReferenceItemGroup = this._cdsProjectDocument.createElementNS(this._cdsNamespace, 'ItemGroup');
    projectReferenceItemGroup.appendChild(this._cdsProjectDocument.createTextNode('\n  '));
    this.addProjectReferenceElement(projectReferenceItemGroup, projectReferenceElement);
    return projectReferenceItemGroup;
  }

  private addProjectReferenceElement(projectReferenceItemGroup: Element, projectReferenceElement: Node): void {
    projectReferenceItemGroup.appendChild(this._cdsProjectDocument.createTextNode('  '));
    projectReferenceItemGroup.appendChild(projectReferenceElement);
    projectReferenceItemGroup.appendChild(this._cdsProjectDocument.createTextNode('\n  '));
  }

  private addProjectReferenceItemGroupElement(projectReferenceItemGroup: Element): void {
    const itemGroups = this.getNamedGroups('ItemGroup');
    if (itemGroups.length > 0) {
      this._cdsProject.insertBefore(projectReferenceItemGroup, itemGroups[itemGroups.length - 1].nextSibling);
      this._cdsProject.insertBefore(this._cdsProjectDocument.createTextNode('\n\n  '), projectReferenceItemGroup);
    }
    else {
      const propertyGroups = this.getNamedGroups('PropertyGroup');
      if (propertyGroups.length > 0) {
        this._cdsProject.insertBefore(projectReferenceItemGroup, propertyGroups[propertyGroups.length - 1].nextSibling);
        this._cdsProject.insertBefore(this._cdsProjectDocument.createTextNode('\n\n  '), projectReferenceItemGroup);
      }
      else {
        const importGroups = this.getNamedGroups('Import');
        if (importGroups.length > 0) {
          this._cdsProject.insertBefore(projectReferenceItemGroup, importGroups[0].nextSibling);
          this._cdsProject.insertBefore(this._cdsProjectDocument.createTextNode('\n\n  '), projectReferenceItemGroup);
        }
        else {
          this._cdsProject.appendChild(this._cdsProjectDocument.createTextNode('\n  '));
          this._cdsProject.appendChild(projectReferenceItemGroup);
          this._cdsProject.appendChild(this._cdsProjectDocument.createTextNode('\n'));
        }
      }
    }
  }
}