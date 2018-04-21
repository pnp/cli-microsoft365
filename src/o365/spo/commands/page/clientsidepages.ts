// Reused by courtesy of PnPJS
// original at: https://github.com/pnp/pnpjs/blob/b18b30ae563d302d5400ef12aeba20f2c8cb5749/packages/sp/src/clientsidepages.ts

/**
 * Interface defining an object with a known property type
 */
export interface TypedHash<T> {
  [key: string]: T;
}

/**
 * Provides functionality to extend the given object by doing a shallow copy
 *
 * @param target The object to which properties will be copied
 * @param source The source object from which properties will be copied
 * @param noOverwrite If true existing properties on the target are not overwritten from the source
 *
 */
function extend(target: any, source: any, noOverwrite = false): any {

  if (!objectDefinedNotNull(source)) {
    return target;
  }

  // ensure we don't overwrite things we don't want overwritten
  const check: (o: any, i: string) => Boolean = noOverwrite ? (o, i) => !(i in o) : () => true;

  return Object.getOwnPropertyNames(source)
    .filter((v: string) => check(target, v))
    .reduce((t: any, v: string) => {
      t[v] = source[v];
      return t;
    }, target);
}

/**
 * Determines if an object is both defined and not null
 * @param obj Object to test
 */
function objectDefinedNotNull(obj: any): boolean {
  return typeof obj !== "undefined" && obj !== null;
}

/**
 * Gets a random GUID value
 *
 * http://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript
 */
function getGUID(): string {
  let d = new Date().getTime();
  const guid = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, function (c) {
    const r = (d + Math.random() * 16) % 16 | 0;
    d = Math.floor(d / 16);
    return (c === "x" ? r : (r & 0x3 | 0x8)).toString(16);
  });
  return guid;
}

/**
 * Page promotion state
 */
export const enum PromotedState {
  /**
   * Regular client side page
   */
  NotPromoted = 0,
  /**
   * Page that will be promoted as news article after publishing
   */
  PromoteOnPublish = 1,
  /**
   * Page that is promoted as news article
   */
  Promoted = 2,
}

/**
 * Type describing the available page layout types for client side "modern" pages
 */
export type ClientSidePageLayoutType = "Article" | "Home";

/**
 * Column size factor. Max value is 12 (= one column), other options are 8,6,4 or 0
 */
export type CanvasColumnFactorType = 0 | 2 | 4 | 6 | 8 | 12;

/**
 * Gets the next order value 1 based for the provided collection
 * 
 * @param collection Collection of orderable things
 */
function getNextOrder(collection: { order: number }[]): number {

  if (collection.length < 1) {
    return 1;
  }

  return Math.max.apply(null, collection.map(i => i.order)) + 1;
}

/**
 * After https://stackoverflow.com/questions/273789/is-there-a-version-of-javascripts-string-indexof-that-allows-for-regular-expr/274094#274094
 * 
 * @param this Types the called context this to a string in which the search will be conducted
 * @param regex A regex or string to match
 * @param startpos A starting position from which the search will begin
 */
function regexIndexOf(this: string, regex: RegExp | string, startpos = 0) {
  const indexOf = this.substring(startpos).search(regex);
  return (indexOf >= 0) ? (indexOf + (startpos)) : indexOf;
}

/**
 * Gets an attribute value from an html string block
 * 
 * @param html HTML to search
 * @param attrName The name of the attribute to find
 */
function getAttrValueFromString(html: string, attrName: string): string | null {
  const reg = new RegExp(`${attrName}="([^"]*?)"`, "i");
  const match = reg.exec(html);
  return match && match.length > 0 ? match[1] : null;
}

/**
 * Finds bounded blocks of markup bounded by divs, ensuring to match the ending div even with nested divs in the interstitial markup
 * 
 * @param html HTML to search
 * @param boundaryStartPattern The starting pattern to find, typically a div with attribute
 * @param collector A func to take the found block and provide a way to form it into a useful return that is added into the return array
 */
function getBoundedDivMarkup<T>(html: string, boundaryStartPattern: RegExp | string, collector: (s: string) => T): T[] {

  const blocks: T[] = [];

  if (typeof html === "undefined" || html === null) {
    return blocks;
  }

  // remove some extra whitespace if present
  const cleanedHtml = html.replace(/[\t\r\n]/g, "");

  // find the first div
  let startIndex = regexIndexOf.call(cleanedHtml, boundaryStartPattern);

  if (startIndex < 0) {
    // we found no blocks in the supplied html
    return blocks;
  }

  // this loop finds each of the blocks
  while (startIndex > -1) {

    // we have one open div counting from the one found above using boundaryStartPattern so we need to ensure we find it's close
    let openCounter = 1;
    let searchIndex = startIndex + 1;
    let nextDivOpen = -1;
    let nextCloseDiv = -1;

    // this loop finds the </div> tag that matches the opening of the control
    while (true) {

      // find both the next opening and closing div tags from our current searching index
      nextDivOpen = regexIndexOf.call(cleanedHtml, /<div[^>]*>/i, searchIndex);
      nextCloseDiv = regexIndexOf.call(cleanedHtml, /<\/div>/i, searchIndex);

      if (nextDivOpen < 0) {
        // we have no more opening divs, just set this to simplify checks below
        nextDivOpen = cleanedHtml.length + 1;
      }

      // determine which we found first, then increment or decrement our counter
      // and set the location to begin searching again
      if (nextDivOpen < nextCloseDiv) {
        openCounter++;
        searchIndex = nextDivOpen + 1;
      } else if (nextCloseDiv < nextDivOpen) {
        openCounter--;
        searchIndex = nextCloseDiv + 1;
      }

      // once we have no open divs back to the level of the opening control div
      // meaning we have all of the markup we intended to find
      if (openCounter === 0) {

        // get the bounded markup, +6 is the size of the ending </div> tag
        const markup = cleanedHtml.substring(startIndex, nextCloseDiv + 6).trim();

        // save the control data we found to the array
        blocks.push(collector(markup));

        // get out of our while loop
        break;
      }

      if (openCounter > 1000 || openCounter < 0) {
        // this is an arbitrary cut-off but likely we will not have 1000 nested divs
        // something has gone wrong above and we are probably stuck in our while loop
        // let's get out of our while loop and not hang everything
        throw new Error("getBoundedDivMarkup exceeded depth parameters.");
      }
    }

    // get the start of the next control
    startIndex = regexIndexOf.call(cleanedHtml, boundaryStartPattern, nextCloseDiv);
  }

  return blocks;
}

/**
 * Normalizes the order value for all the sections, columns, and controls to be 1 based and stepped (1, 2, 3...)
 * 
 * @param collection The collection to normalize
 */
function reindex(collection?: { order: number, columns?: { order: number }[], controls?: { order: number }[] }[]): void {
  if (!collection) {
    return;
  }

  for (let i = 0; i < collection.length; i++) {
    collection[i].order = i + 1;
    if (collection[i].hasOwnProperty("columns")) {
      reindex(collection[i].columns);
    } else if (collection[i].hasOwnProperty("controls")) {
      reindex(collection[i].controls);
    }
  }
}

/**
 * Represents the data and methods associated with client side "modern" pages
 */
export class ClientSidePage {
  public sections: CanvasSection[] = [];
  public commentsDisabled = false;

  /**
   * Converts a json object to an escaped string appropriate for use in attributes when storing client-side controls
   * 
   * @param json The json object to encode into a string
   */
  public static jsonToEscapedString(json: any): string {

    return JSON.stringify(json)
      .replace(/"/g, "&quot;")
      .replace(/:/g, "&#58;")
      .replace(/{/g, "&#123;")
      .replace(/}/g, "&#125;");
  }

  /**
   * Converts an escaped string from a client-side control attribute to a json object
   * 
   * @param escapedString 
   */
  public static escapedStringToJson<T = any>(escapedString: string | null): T {
    if (!escapedString) {
      return {} as any;
    }

    return JSON.parse(escapedString
      .replace(/&quot;/g, `"`)
      .replace(/&#58;/g, ":")
      .replace(/&#123;/g, "{")
      .replace(/&#125;/g, "}"));
  }

  /**
   * Add a section to this page
   */
  public addSection(): CanvasSection {

    const section = new CanvasSection(this, getNextOrder(this.sections));
    this.sections.push(section);
    return section;
  }

  /**
   * Converts this page's content to html markup
   */
  public toHtml(): string {

    // trigger reindex of the entire tree
    reindex(this.sections);

    const html: string[] = [];

    html.push("<div>");

    for (let i = 0; i < this.sections.length; i++) {
      html.push(this.sections[i].toHtml());
    }

    html.push("</div>");

    return html.join("");
  }

  /**
   * Loads this page instance's content from the supplied html
   * 
   * @param html html string representing the page's content
   */
  public static fromHtml(html: string): ClientSidePage {
    const page: ClientSidePage = new ClientSidePage();

    // reset sections
    page.sections = [];

    // gather our controls from the supplied html
    getBoundedDivMarkup(html, /<div\b[^>]*data-sp-canvascontrol[^>]*?>/i, markup => {

      // get the control type
      const ct = /controlType&quot;&#58;(\d*?),/i.exec(markup);

      // if no control type is present this is a column which we give type 0 to let us process it
      const controlType = ct == null || ct.length < 2 ? 0 : parseInt(ct[1], 10);

      let control: CanvasControl | null = null;

      switch (controlType) {
        case 0:
          // empty canvas column
          control = new CanvasColumn(null, 0);
          control.fromHtml(markup);
          page.mergeColumnToTree(<CanvasColumn>control);
          break;
        case 3:
          // client side webpart
          control = new ClientSideWebpart("");
          control.fromHtml(markup);
          page.mergePartToTree(<ClientSidePart>control);
          break;
        case 4:
          // client side text
          control = new ClientSideText();
          control.fromHtml(markup);
          page.mergePartToTree(<ClientSidePart>control);
          break;
      }
    });

    // refresh all the orders within the tree
    reindex(page.sections);

    return page;
  }

  /**
   * Finds a control by the specified instance id
   * 
   * @param id Instance id of the control to find
   */
  public findControlById<T extends ClientSidePart = ClientSidePart>(id: string): T | null {
    return this.findControl((c) => c.id === id);
  }

  /**
   * Finds a control within this page's control tree using the supplied predicate
   * 
   * @param predicate Takes a control and returns true or false, if true that control is returned by findControl
   */
  public findControl<T extends ClientSidePart = ClientSidePart>(predicate: (c: ClientSidePart) => boolean): T | null {
    // check all sections
    for (let i = 0; i < this.sections.length; i++) {
      // check all columns
      for (let j = 0; j < this.sections[i].columns.length; j++) {
        // check all controls
        for (let k = 0; k < this.sections[i].columns[j].controls.length; k++) {
          // check to see if the predicate likes this control
          if (predicate(this.sections[i].columns[j].controls[k])) {
            return <T>this.sections[i].columns[j].controls[k];
          }
        }
      }
    }

    // we found nothing so give nothing back
    return null;
  }

  /**
   * Merges the control into the tree of sections and columns for this page
   * 
   * @param control The control to merge
   */
  private mergePartToTree(control: ClientSidePart): void {

    let section: CanvasSection | null = null;
    let column: CanvasColumn | null = null;

    const sections = this.sections.filter(s => control.controlData && s.order === control.controlData.position.zoneIndex);
    if (sections.length < 1) {
      section = new CanvasSection(this, control.controlData ? control.controlData.position.zoneIndex : 0);
      this.sections.push(section);
    } else {
      section = sections[0];
    }

    const columns = section.columns.filter(c => control.controlData && c.order === control.controlData.position.sectionIndex);
    if (columns.length < 1) {
      column = new CanvasColumn(section, control.controlData ? control.controlData.position.sectionIndex : 0, control.controlData ? control.controlData.position.sectionFactor : 0);
      section.columns.push(column);
    } else {
      column = columns[0];
    }

    control.column = column;
    column.addControl(control);
  }

  /**
   * Merges the supplied column into the tree
   * 
   * @param column Column to merge
   * @param position The position data for the column
   */
  private mergeColumnToTree(column: CanvasColumn): void {

    let section: CanvasSection | null = null;
    const sections = this.sections.filter(s => column.controlData && s.order === column.controlData.position.zoneIndex);

    if (sections.length < 1) {
      section = new CanvasSection(this, column.controlData ? column.controlData.position.zoneIndex : 0);
      this.sections.push(section);
    } else {
      section = sections[0];
    }

    column.section = section;
    section.columns.push(column);
  }
}

export class CanvasSection {

  /**
   * Used to track this object inside the collection at runtime
   */
  private _memId: string;

  constructor(public page: ClientSidePage, public order: number, public columns: CanvasColumn[] = []) {
    this._memId = getGUID();
  }

  /**
   * Default column (this.columns[0]) for this section
   */
  public get defaultColumn(): CanvasColumn {

    if (this.columns.length < 1) {
      this.addColumn(12);
    }

    return this.columns[0];
  }

  /**
   * Adds a new column to this section
   */
  public addColumn(factor: CanvasColumnFactorType): CanvasColumn {

    const column = new CanvasColumn(this, getNextOrder(this.columns), factor);
    this.columns.push(column);
    return column;
  }

  /**
   * Adds a control to the default column for this section
   * 
   * @param control Control to add to the default column
   */
  public addControl(control: ClientSidePart): this {
    this.defaultColumn.addControl(control);
    return this;
  }

  public toHtml(): string {

    const html = [];

    for (let i = 0; i < this.columns.length; i++) {
      html.push(this.columns[i].toHtml());
    }

    return html.join("");
  }

  /**
   * Removes this section and all contained columns and controls from the collection
   */
  public remove(): void {
    this.page.sections = this.page.sections.filter(section => section._memId !== this._memId);
    reindex(this.page.sections);
  }
}

export abstract class CanvasControl {

  constructor(
    protected controlType: number | undefined,
    public dataVersion: string | null,
    public column: CanvasColumn | null = null,
    public order = 1,
    public id: string | undefined = getGUID(),
    public controlData: ClientSideControlData | null = null) { }

  /**
   * Value of the control's "data-sp-controldata" attribute
   */
  public get jsonData(): string {
    return ClientSidePage.jsonToEscapedString(this.getControlData());
  }

  public abstract toHtml(index: number): string;

  public fromHtml(html: string): void {
    this.controlData = ClientSidePage.escapedStringToJson<ClientSideControlData>(getAttrValueFromString(html, "data-sp-controldata"));
    this.dataVersion = getAttrValueFromString(html, "data-sp-canvasdataversion");
    this.controlType = this.controlData.controlType;
    this.id = this.controlData.id;
  }

  protected abstract getControlData(): ClientSideControlData;
}

export class CanvasColumn extends CanvasControl {

  constructor(
    public section: CanvasSection | null,
    public order: number,
    public factor: CanvasColumnFactorType = 12,
    public controls: ClientSidePart[] = [],
    dataVersion = "1.0") {
    super(0, dataVersion);
  }

  public addControl(control: ClientSidePart): this {
    control.column = this;
    this.controls.push(control);
    return this;
  }

  public getControl<T extends ClientSidePart>(index: number): T {
    return <T>this.controls[index];
  }

  public toHtml(): string {
    const html = [];

    if (this.controls.length < 1) {

      html.push(`<div data-sp-canvascontrol="" data-sp-canvasdataversion="${this.dataVersion}" data-sp-controldata="${this.jsonData}"></div>`);

    } else {

      for (let i = 0; i < this.controls.length; i++) {
        html.push(this.controls[i].toHtml(i + 1));
      }
    }

    return html.join("");
  }

  public fromHtml(html: string): void {
    super.fromHtml(html);

    this.controlData = ClientSidePage.escapedStringToJson<ClientSideControlData>(getAttrValueFromString(html, "data-sp-controldata"));
    this.factor = this.controlData.position.sectionFactor;
    this.order = this.controlData.position.sectionIndex;
  }

  public getControlData(): ClientSideControlData {
    return {
      displayMode: 2,
      position: {
        sectionFactor: this.factor,
        sectionIndex: this.order,
        zoneIndex: this.section ? this.section.order : 0,
      },
    };
  }

  /**
   * Removes this column and all contained controls from the collection
   */
  public remove(): void {
    if (this.section) {
      this.section.columns = this.section.columns.filter(column => column.id !== this.id);
      if (this.column) {
        reindex(this.column.controls);
      }
    }
  }
}

/**
 * Abstract class with shared functionality for parts
 */
export abstract class ClientSidePart extends CanvasControl {

  /**
   * Removes this column and all contained controls from the collection
   */
  public remove(): void {
    if (this.column) {
      this.column.controls = this.column.controls.filter(control => control.id !== this.id);
      reindex(this.column.controls);
    }
  }
}

export class ClientSideText extends ClientSidePart {

  private _text: string;

  constructor(text = "") {
    super(4, "1.0");

    this._text = text;
  }

  /**
   * The text markup of this control
   */
  public get text(): string {
    return this._text;
  }

  public set text(text: string) {

    if (!text.startsWith("<p>")) {
      text = `<p>${text}</p>`;
    }

    this._text = text;
  }

  public getControlData(): ClientSideControlData {

    return {
      controlType: this.controlType,
      editorType: "CKEditor",
      id: this.id,
      position: {
        controlIndex: this.order,
        sectionFactor: this.column ? this.column.factor : 0,
        sectionIndex: this.column ? this.column.order : 0,
        zoneIndex: this.column && this.column.section ? this.column.section.order : 0,
      },
    };
  }

  public toHtml(index: number): string {

    // set our order to the value passed in
    this.order = index;

    const html: string[] = [];

    html.push(`<div data-sp-canvascontrol="" data-sp-canvasdataversion="${this.dataVersion}" data-sp-controldata="${this.jsonData}">`);
    html.push("<div data-sp-rte=\"\">");
    html.push(`${this.text}`);
    html.push("</div>");
    html.push("</div>");

    return html.join("");
  }

  public fromHtml(html: string): void {

    super.fromHtml(html);

    const match = /<div[^>]*data-sp-rte[^>]*>(.*?)<\/div>$/i.exec(html);

    this.text = match && match.length > 1 ? match[1] : "";
  }
}

export class ClientSideWebpart extends ClientSidePart {

  constructor(public title: string,
    public description = "",
    public propertieJson: TypedHash<any> = {},
    public webPartId = "",
    protected htmlProperties = "",
    protected serverProcessedContent: ServerProcessedContent | null = null,
    protected canvasDataVersion: string | null = "1.0") {
    super(3, "1.0");
  }

  public static fromComponentDef(definition: ClientSidePageComponent): ClientSideWebpart {
    const part = new ClientSideWebpart("");
    part.import(definition);
    return part;
  }

  public import(component: ClientSidePageComponent): void {
    this.webPartId = component.Id.replace(/^\{|\}$/g, "").toLowerCase();
    const manifest: ClientSidePageComponentManifest = JSON.parse(component.Manifest);
    this.title = manifest.preconfiguredEntries[0].title.default;
    this.description = manifest.preconfiguredEntries[0].description.default;
    this.dataVersion = "";
    this.propertieJson = this.parseJsonProperties(manifest.preconfiguredEntries[0].properties);
  }

  public setProperties<T = any>(properties: T): this {
    this.propertieJson = extend(this.propertieJson, properties);
    return this;
  }

  public getProperties<T = any>(): T {
    return <T>this.propertieJson;
  }

  public toHtml(index: number): string {

    // set our order to the value passed in
    this.order = index;

    // will form the value of the data-sp-webpartdata attribute
    const data = {
      dataVersion: this.dataVersion,
      description: this.description,
      id: this.webPartId,
      instanceId: this.id,
      properties: this.propertieJson,
      title: this.title,
    };

    const html: string[] = [];

    html.push(`<div data-sp-canvascontrol="" data-sp-canvasdataversion="${this.canvasDataVersion}" data-sp-controldata="${this.jsonData}">`);

    html.push(`<div data-sp-webpart="" data-sp-webpartdataversion="${this.dataVersion}" data-sp-webpartdata="${ClientSidePage.jsonToEscapedString(data)}">`);

    html.push(`<div data-sp-componentid>`);
    html.push(this.webPartId);
    html.push("</div>");

    html.push(`<div data-sp-htmlproperties="">`);
    html.push(this.renderHtmlProperties());
    html.push("</div>");

    html.push("</div>");
    html.push("</div>");

    return html.join("");
  }

  public fromHtml(html: string): void {

    super.fromHtml(html);

    const webPartData = ClientSidePage.escapedStringToJson<ClientSideWebpartData>(getAttrValueFromString(html, "data-sp-webpartdata"));

    this.title = webPartData.title;
    this.description = webPartData.description;
    this.webPartId = webPartData.id;
    this.canvasDataVersion = getAttrValueFromString(html, "data-sp-canvasdataversion");
    this.dataVersion = getAttrValueFromString(html, "data-sp-webpartdataversion");
    this.setProperties(webPartData.properties);

    if (typeof webPartData.serverProcessedContent !== "undefined") {
      this.serverProcessedContent = webPartData.serverProcessedContent;
    }

    // get our html properties
    const htmlProps = getBoundedDivMarkup(html, /<div\b[^>]*data-sp-htmlproperties[^>]*?>/i, markup => {
      return markup.replace(/^<div\b[^>]*data-sp-htmlproperties[^>]*?>/i, "").replace(/<\/div>$/i, "");
    });

    this.htmlProperties = htmlProps.length > 0 ? htmlProps[0] : "";
  }

  public getControlData(): ClientSideControlData {

    return {
      controlType: this.controlType,
      id: this.id,
      position: {
        controlIndex: this.order,
        sectionFactor: this.column ? this.column.factor : 0,
        sectionIndex: this.column ? this.column.order : 0,
        zoneIndex: this.column && this.column.section ? this.column.section.order : 0,
      },
      webPartId: this.webPartId,
    };

  }

  protected renderHtmlProperties(): string {

    const html: string[] = [];

    if (typeof this.serverProcessedContent === "undefined" || this.serverProcessedContent === null) {

      html.push(this.htmlProperties);

    } else if (typeof this.serverProcessedContent !== "undefined") {

      if (typeof this.serverProcessedContent.searchablePlainTexts !== "undefined") {

        const keys = Object.keys(this.serverProcessedContent.searchablePlainTexts);
        for (let i = 0; i < keys.length; i++) {
          html.push(`<div data-sp-prop-name="${keys[i]}" data-sp-searchableplaintext="true">`);
          html.push(this.serverProcessedContent.searchablePlainTexts[keys[i]]);
          html.push("</div>");
        }
      }

      if (typeof this.serverProcessedContent.imageSources !== "undefined") {

        const keys = Object.keys(this.serverProcessedContent.imageSources);
        for (let i = 0; i < keys.length; i++) {
          html.push(`<img data-sp-prop-name="${keys[i]}" src="${this.serverProcessedContent.imageSources[keys[i]]}" />`);
        }
      }

      if (typeof this.serverProcessedContent.links !== "undefined") {

        const keys = Object.keys(this.serverProcessedContent.links);
        for (let i = 0; i < keys.length; i++) {
          html.push(`<a data-sp-prop-name="${keys[i]}" href="${this.serverProcessedContent.links[keys[i]]}"></a>`);
        }
      }
    }

    return html.join("");
  }

  protected parseJsonProperties(props: TypedHash<any>): any {

    // If the web part has the serverProcessedContent property then keep this one as it might be needed as input to render the web part HTML later on
    if (typeof props.webPartData !== "undefined" && typeof props.webPartData.serverProcessedContent !== "undefined") {
      this.serverProcessedContent = props.webPartData.serverProcessedContent;
    } else if (typeof props.serverProcessedContent !== "undefined") {
      this.serverProcessedContent = props.serverProcessedContent;
    } else {
      this.serverProcessedContent = null;
    }

    if (typeof props.webPartData !== "undefined" && typeof props.webPartData.properties !== "undefined") {
      return props.webPartData.properties;
    } else if (typeof props.properties !== "undefined") {
      return props.properties;
    } else {
      return props;
    }
  }
}

/**
 * Client side webpart object (retrieved via the _api/web/GetClientSideWebParts REST call)
 */
export interface ClientSidePageComponent {
  /**
   * Component type for client side webpart object
   */
  ComponentType: number;
  /**
   * Id for client side webpart object
   */
  Id: string;
  /**
   * Manifest for client side webpart object
   */
  Manifest: string;
  /**
   * Manifest type for client side webpart object
   */
  ManifestType: number;
  /**
   * Name for client side webpart object
   */
  Name: string;
  /**
   * Status for client side webpart object
   */
  Status: number;
}

interface ClientSidePageComponentManifest {
  alias: string;
  componentType: "WebPart" | "" | null;
  disabledOnClassicSharepoint: boolean;
  hiddenFromToolbox: boolean;
  id: string;
  imageLinkPropertyNames: any;
  isInternal: boolean;
  linkPropertyNames: boolean;
  loaderConfig: any;
  manifestVersion: number;
  preconfiguredEntries: {
    description: { default: string };
    group: { default: string };
    groupId: string;
    iconImageUrl: string;
    officeFabricIconFontName: string;
    properties: TypedHash<any>;
    title: { default: string };

  }[];
  preloadComponents: any | null;
  requiredCapabilities: any | null;
  searchablePropertyNames: any | null;
  supportsFullBleed: boolean;
  version: string;
}

export interface ServerProcessedContent {
  searchablePlainTexts: TypedHash<string>;
  imageSources: TypedHash<string>;
  links: TypedHash<string>;
}

export interface ClientSideControlPosition {
  controlIndex?: number;
  sectionFactor: CanvasColumnFactorType;
  sectionIndex: number;
  zoneIndex: number;
}

export interface ClientSideControlData {
  controlType?: number;
  id?: string;
  editorType?: string;
  position: ClientSideControlPosition;
  webPartId?: string;
  displayMode?: number;
}

export interface ClientSideWebpartData {
  dataVersion: string;
  description: string;
  id: string;
  instanceId: string;
  properties: any;
  title: string;
  serverProcessedContent?: ServerProcessedContent;
}

export module ClientSideWebpartPropertyTypes {

  /**
   * Propereties for Embed (component id: 490d7c76-1824-45b2-9de3-676421c997fa)
   */
  export interface Embed {
    embedCode: string;
    cachedEmbedCode?: string;
    shouldScaleWidth?: boolean;
    tempState?: any;
  }

  /**
   * Properties for Bing Map (component id: e377ea37-9047-43b9-8cdb-a761be2f8e09)
   */
  export interface BingMap {
    center: {
      altitude?: number;
      altitudeReference?: number;
      latitude: number;
      longitude: number;
    };
    mapType: "aerial" | "birdseye" | "road" | "streetside";
    maxNumberOfPushPins?: number;
    pushPins?: {
      location: {
        latitude: number;
        longitude: number;
        altitude?: number;
        altitudeReference?: number;
      };
      address?: string;
      defaultAddress?: string;
      defaultTitle?: string;
      title?: string;
    }[];
    shouldShowPushPinTitle?: boolean;
    zoomLevel?: number;
  }
}