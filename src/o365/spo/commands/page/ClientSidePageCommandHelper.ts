import * as request from 'request-promise-native';
import {
	ClientSidePage,
	ClientSideWebpart,
	ClientSidePageComponent,
	CanvasSection,
	CanvasColumn,
	ClientSideText
} from './clientsidepages';
import Utils from '../../../../Utils';
import { StandardWebPart, StandardWebPartUtils } from '../../common/StandardWebpartTypes';

export interface IRequestContext {
	accessToken: string;
	requestDigest: string;
}

export interface ICommandContext {
	requestContext: IRequestContext;
	webUrl: string;
	pageName: string;
	debug: boolean;
	verbose: boolean;
	log: (message: any) => void;
}

const debug = (ctx: ICommandContext, message: any) => {
	if (ctx.debug && ctx.log) {
		ctx.log(message);
	}
};

const say = (ctx: ICommandContext, message: any) => {
	if (ctx.verbose && ctx.log) {
		ctx.log(message);
	}
};

class ClientSidePageCommandHelper {
	public static saveClientSidePage(context: ICommandContext, clientSidePage: ClientSidePage): request.RequestPromise {
		const serverRelativeSiteUrl: string = `${context.webUrl.substr(
			context.webUrl.indexOf('/', 8)
		)}/sitepages/${context.pageName}`;

		const updatedContent = clientSidePage.toHtml();

		debug(context, 'Updated canvas content: ');
		debug(context, updatedContent);
		debug(context, '');

		const requestOptions: any = {
			url: `${context.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeSiteUrl}')/ListItemAllFields`,
			headers: Utils.getRequestHeaders({
				authorization: `Bearer ${context.requestContext.accessToken}`,
				'X-RequestDigest': context.requestContext.requestDigest,
				'content-type': 'application/json;odata=verbose;charset=utf-8',
				'X-HTTP-Method': 'MERGE',
				'IF-MATCH': '*',
				accept: 'application/json;odata=nometadata'
			}),
			body: {
				__metadata: { type: 'SP.Data.SitePagesItem' },
				CanvasContent1: updatedContent
			},
			json: true
		};

		debug(context, 'Executing web request...');
		debug(context, requestOptions);
		debug(context, '');

		return request.post(requestOptions);
	}

	public static getClientSidePage(context: ICommandContext): Promise<ClientSidePage> {
		return new Promise<ClientSidePage>((resolve, reject) => {
			say(context, `Retrieving information about the page...`);

			const requestOptions: any = {
				url: `${context.webUrl}/_api/web/getfilebyserverrelativeurl('${context.webUrl.substr(
					context.webUrl.indexOf('/', 8)
				)}/sitepages/${encodeURIComponent(
					context.pageName
				)}')?$expand=ListItemAllFields/ClientSideApplicationId`,
				headers: Utils.getRequestHeaders({
					authorization: `Bearer ${context.requestContext.accessToken}`,
					'content-type': 'application/json;charset=utf-8',
					accept: 'application/json;odata=nometadata'
				}),
				json: true
			};

			debug(context, 'Executing web request...');
			debug(context, requestOptions);
			debug(context, '');

			request.get(requestOptions).then((res) => {
				debug(context, 'Response:');
				debug(context, res);
				debug(context, '');

				if (res.ListItemAllFields.ClientSideApplicationId !== 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec') {
					reject(new Error(`Page ${context.pageName} is not a modern page.`));
					return;
				}

				const clientSidePage: ClientSidePage = ClientSidePage.fromHtml(res.ListItemAllFields.CanvasContent1);

				resolve(clientSidePage);
			});
		});
	}

	public static getStandardWebPartInstance(
		context: ICommandContext,
		name: StandardWebPart
	): Promise<ClientSideWebpart> {
		return new Promise<ClientSideWebpart>((resolve, reject) => {
			const webPartId = StandardWebPartUtils.getWebPartId(name);

			if (!webPartId) {
				reject(new Error(`There is no available standard WebPart with name '${name}'.`));
				return;
			}

			debug(context, `WebPart Name: ${name}`);
			debug(context, `WebPartId: ${webPartId}`);

			const requestOptions: any = {
				url: `${context.webUrl}/_api/web/getclientsidewebparts()`,
				headers: Utils.getRequestHeaders({
					authorization: `Bearer ${context.requestContext.accessToken}`,
					'content-type': 'application/json;charset=utf-8',
					accept: 'application/json;odata=nometadata'
				}),
				json: true
			};

			debug(context, 'Executing web request...');
			debug(context, requestOptions);
			debug(context, '');

			request.get(requestOptions).then((res: { value: ClientSidePageComponent[] }) => {
				debug(context, 'Response:');
				debug(context, res);
				debug(context, '');

				const webPartDefinition = res.value.filter((c) => c.Id === webPartId);
				if (webPartDefinition.length == 0) {
					reject(new Error(`There is no available WebPart with Id '${webPartId}'.`));
					return;
				}

				debug(context, 'WebPart definition:');
				debug(context, webPartDefinition);
				debug(context, '');

				say(context, `Creating instance from definition of WebPart ${webPartId}...`);

				const webPart = ClientSideWebpart.fromComponentDef(webPartDefinition[0]);
				resolve(webPart);
			});
		});
	}

	private static getTargetCanvasColumn(
		context: ICommandContext,
		clientSidePage: ClientSidePage,
		section: number | null,
		column: number | null
	): CanvasColumn {
    debug(context, `Get target column of page section ${section} column ${column}`);
		let canvasSection: CanvasSection | null = null;
		let canvasColumn: CanvasColumn | null = null;
		// If the page has no section at all, add a default column and section
		if (clientSidePage.sections.length == 0) {
      debug(context, `The page has no sections. A default one will be created`);
			canvasSection = clientSidePage.addSection();
			canvasColumn = canvasSection.defaultColumn;
		} else {
			// Otherwise, get the section and column according to command arguments
			let actualSectionIndex: number | null = section;
			// If the section arg is not specified, set to first section
			if (!actualSectionIndex && actualSectionIndex != 0) {
				debug(context, `No section argument specified, The component will be added to default section`);
				actualSectionIndex = 1;
			}
			// Make sure the section is in the range, other add a new section
			if (actualSectionIndex < 1 || actualSectionIndex > clientSidePage.sections.length) {
				throw new Error(`Invalid Section '${section}'`);
			}

			const foundSections = clientSidePage.sections.filter((s) => s.order == actualSectionIndex);
			canvasSection = foundSections[0];
      
			let actualColumnIndex: number | null = column;
			// If the column arg is not specified, set to first column
			if (!actualColumnIndex && actualColumnIndex != 0) {
				debug(context, `No column argument specified, The component will be added to default column`);
				actualColumnIndex = 1;
			}
			// Make sure the column is in the range of the current section
			if (actualColumnIndex < 1 || actualColumnIndex > canvasSection.columns.length) {
				throw new Error(`Invalid Column '${column}'`);
			}

      debug(context, `found columns length ${canvasSection.columns.length}`);

			// Get the target column
      const foundColumns = canvasSection.columns.filter(c => c.order == actualColumnIndex);
      canvasColumn = foundColumns[0];
    }
    
    return canvasColumn;
	}

	public static getWebPartInstance(context: ICommandContext, webPartId: string): Promise<ClientSideWebpart> {
		return new Promise<ClientSideWebpart>((resolve, reject) => {
			if (!webPartId) {
				reject(new Error(`The WebPart id argument is missing.`));
				return;
			}

			if (!Utils.isValidGuid(webPartId)) {
				reject(new Error(`The specified WebPart id argument '${webPartId}' is not a valid GUID.`));
				return;
			}

			debug(context, `WebPartId: ${webPartId}`);

			const requestOptions: any = {
				url: `${context.webUrl}/_api/web/getclientsidewebparts()`,
				headers: Utils.getRequestHeaders({
					authorization: `Bearer ${context.requestContext.accessToken}`,
					'content-type': 'application/json;charset=utf-8',
					accept: 'application/json;odata=nometadata'
				}),
				json: true
			};

			debug(context, 'Executing web request...');
			debug(context, requestOptions);
			debug(context, '');

			request.get(requestOptions).then((res: { value: ClientSidePageComponent[] }) => {
				debug(context, 'Response:');
				debug(context, res);
				debug(context, '');

				const webPartDefinition = res.value.filter((c) => c.Id === webPartId);
				if (webPartDefinition.length == 0) {
					reject(new Error(`There is no available WebPart with Id '${webPartId}'.`));
					return;
				}

				debug(context, 'WebPart definition:');
				debug(context, webPartDefinition);
				debug(context, '');

				say(context, `Creating instance from definition of WebPart ${webPartId}...`);

				const webPart = ClientSideWebpart.fromComponentDef(webPartDefinition[0]);
				resolve(webPart);
			});
		});
	}

	public static addWebPartToPage(
		clientSidePage: ClientSidePage,
		webPart: ClientSideWebpart,
		context: ICommandContext,
		section: number = 1,
		column: number = 1,
		order?: number,
		propertiesJSON: any = null
	) {
    // Get the appropriate column according to page current layout and specified arguments
    const canvasColumn = ClientSidePageCommandHelper.getTargetCanvasColumn(context, clientSidePage, section, column);

		// Set the WebPart order
		if (order) {
			webPart.order = order;
			debug(context, 'WebPart order: ');
			debug(context, webPart.order);
		}

		if (propertiesJSON) {
			debug(context, 'WebPart properties: ');
			debug(context, propertiesJSON);
			debug(context, '');

			try {
				let properties = JSON.parse(propertiesJSON);
				webPart.setProperties(properties);
			} catch (error) {
				throw new Error('WebPart properties cannot be set');
			}
		}

		// Insert at specific order if specified
		if (order) {
			debug(context, 'WebPart order: ');
			debug(context, order);
			canvasColumn.insertControl(webPart, order - 1);
		} else {
			// Add the WebPart at the end of the cell contents
			canvasColumn.addControl(webPart);
		}
	}

	public static addTextToPage(
		clientSidePage: ClientSidePage,
		text: string,
		context: ICommandContext,
		section: number = 1,
		column: number = 1,
    order?: number
	) {		
    // Get the appropriate column according to page current layout and specified arguments
    const canvasColumn = ClientSidePageCommandHelper.getTargetCanvasColumn(context, clientSidePage, section, column);

		// Instantiate a new ClientSideText object
		const csText = new ClientSideText(text);

		// Insert at specific order if specified
		if (order) {
			debug(context, 'Text order: ');
			debug(context, order);
			canvasColumn.insertControl(csText, order - 1);
		} else {
			// Add the control at the end of the cell contents
			canvasColumn.addControl(csText);
		}
	}
}

export default ClientSidePageCommandHelper;
