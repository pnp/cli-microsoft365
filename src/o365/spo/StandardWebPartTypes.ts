export type StandardWebPart =
	| 'ContentRollup'
	| 'BingMap'
	| 'ContentEmbed'
	| 'DocumentEmbed'
	| 'Image'
	| 'ImageGallery'
	| 'LinkPreview'
	| 'NewsFeed'
	| 'NewsReel'
	| 'PowerBIReportEmbed'
	| 'QuickChart'
	| 'SiteActivity'
	| 'VideoEmbed'
	| 'YammerEmbed'
	| 'Events'
	| 'GroupCalendar'
	| 'Hero'
	| 'List'
	| 'PageTitle'
	| 'People'
	| 'QuickLinks'
	| 'CustomMessageRegion'
	| 'Divider'
	| 'MicrosoftForms'
	| 'Spacer';

const StandardWebParts = [
	{ name: 'ContentRollup', id: 'daf0b71c-6de8-4ef7-b511-faae7c388708' },
	{ name: 'BingMap', id: 'e377ea37-9047-43b9-8cdb-a761be2f8e09' },
	{ name: 'ContentEmbed', id: '490d7c76-1824-45b2-9de3-676421c997fa' },
	{ name: 'DocumentEmbed', id: 'b7dd04e1-19ce-4b24-9132-b60a1c2b910d' },
	{ name: 'Image', id: 'd1d91016-032f-456d-98a4-721247c305e8' },
	{ name: 'ImageGallery', id: 'af8be689-990e-492a-81f7-ba3e4cd3ed9c' },
	{ name: 'LinkPreview', id: '6410b3b6-d440-4663-8744-378976dc041e' },
	{ name: 'NewsFeed', id: '0ef418ba-5d19-4ade-9db0-b339873291d0' },
	{ name: 'NewsReel', id: 'a5df8fdf-b508-4b66-98a6-d83bc2597f63' },
	// Seems like we've been having 2 guids to identify this web part...
	{ name: 'NewsReel', id: '8c88f208-6c77-4bdb-86a0-0c47b4316588' },
	{ name: 'PowerBIReportEmbed', id: '58fcd18b-e1af-4b0a-b23b-422c2c52d5a2' },
	{ name: 'QuickChart', id: '91a50c94-865f-4f5c-8b4e-e49659e69772' },
	{ name: 'SiteActivity', id: 'eb95c819-ab8f-4689-bd03-0c2d65d47b1f' },
	{ name: 'VideoEmbed', id: '275c0095-a77e-4f6d-a2a0-6a7626911518' },
	{ name: 'YammerEmbed', id: '31e9537e-f9dc-40a4-8834-0e3b7df418bc' },
	{ name: 'Events', id: '20745d7d-8581-4a6c-bf26-68279bc123fc' },
	{ name: 'GroupCalendar', id: '6676088b-e28e-4a90-b9cb-d0d0303cd2eb' },
	{ name: 'Hero', id: 'c4bd7b2f-7b6e-4599-8485-16504575f590' },
	{ name: 'List', id: 'f92bf067-bc19-489e-a556-7fe95f508720' },
	{ name: 'PageTitle', id: 'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788' },
	{ name: 'People', id: '7f718435-ee4d-431c-bdbf-9c4ff326f46e' },
	{ name: 'QuickLinks', id: 'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd' },
	{ name: 'CustomMessageRegion', id: '71c19a43-d08c-4178-8218-4df8554c0b0e' },
	{ name: 'Divider', id: '2161a1c6-db61-4731-b97c-3cdb303f7cbb' },
	{ name: 'MicrosoftForms', id: 'b19b3b9e-8d13-4fec-a93c-401a091c0707' },
	{ name: 'Spacer', id: '8654b779-4886-46d4-8ffb-b5ed960ee986' },
	{ name: 'ClientWebPart', id: '243166f5-4dc3-4fe2-9df2-a7971b546a0a' }
];

export class StandardWebPartUtils {
	public static getWebPartId(type: StandardWebPart): string | null {
		let foundWebParts = StandardWebParts.filter((wp) => wp.name == type);
		return foundWebParts.length > 0 ? foundWebParts[0].id : null;
	}
  
  public static isValidStandardWebPartType(type: string) : boolean {
    return StandardWebPartUtils.getWebPartId(type as StandardWebPart) != null;
  }
}


