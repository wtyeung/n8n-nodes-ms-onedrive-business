import type {
	IDataObject,
	IExecuteFunctions,
	IHookFunctions,
	IHttpRequestMethods,
	IHttpRequestOptions,
	ILoadOptionsFunctions,
	IWebhookFunctions,
	JsonObject,
} from 'n8n-workflow';
import { NodeApiError } from 'n8n-workflow';

export async function microsoftApiRequest(
	this: IExecuteFunctions | ILoadOptionsFunctions | IHookFunctions | IWebhookFunctions,
	method: IHttpRequestMethods,
	resource: string,
	body: IDataObject | string = {},
	qs: IDataObject = {},
	uri?: string,
	headers: IDataObject = {},
	option: IDataObject = {},
// eslint-disable-next-line @typescript-eslint/no-explicit-any
): Promise<any> {
	let options: IHttpRequestOptions = {
		headers: {
			'Content-Type': 'application/json',
			...headers,
		},
		method,
		body,
		qs,
		url: uri || `https://graph.microsoft.com/v1.0${resource}`,
		json: true,
	};

	options = Object.assign({}, options, option);

	try {
		if (Object.keys(body).length === 0 && method !== 'GET') {
			delete options.body;
		}

		return await this.helpers.httpRequestWithAuthentication.call(
			this,
			'microsoftOneDriveBusinessOAuth2Api',
			options,
		);
	} catch (error) {
		throw new NodeApiError(this.getNode(), error as JsonObject);
	}
}

export async function microsoftApiRequestAllItems(
	this: IExecuteFunctions | ILoadOptionsFunctions | IHookFunctions | IWebhookFunctions,
	propertyName: string,
	method: IHttpRequestMethods,
	endpoint: string,
	body: IDataObject = {},
	query: IDataObject = {},
// eslint-disable-next-line @typescript-eslint/no-explicit-any
): Promise<any> {
	const returnData: IDataObject[] = [];

	let responseData;
	let uri: string | undefined;

	do {
		responseData = await microsoftApiRequest.call(this, method, endpoint, body, query, uri);
		uri = responseData['@odata.nextLink'];
		if (uri?.includes('$skiptoken=')) {
			delete query.$skiptoken;
		}
		returnData.push.apply(returnData, responseData[propertyName] as IDataObject[]);
	} while (responseData['@odata.nextLink'] !== undefined);

	return returnData;
}

export function getFileExtension(fileName: string): string {
	const lastDot = fileName.lastIndexOf('.');
	if (lastDot === -1) return '';
	return fileName.substring(lastDot + 1).toLowerCase();
}

export function getMimeType(fileName: string): string {
	const extension = getFileExtension(fileName);
	const mimeTypes: { [key: string]: string } = {
		pdf: 'application/pdf',
		doc: 'application/msword',
		docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		xls: 'application/vnd.ms-excel',
		xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		ppt: 'application/vnd.ms-powerpoint',
		pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		txt: 'text/plain',
		csv: 'text/csv',
		json: 'application/json',
		xml: 'application/xml',
		zip: 'application/zip',
		jpg: 'image/jpeg',
		jpeg: 'image/jpeg',
		png: 'image/png',
		gif: 'image/gif',
		svg: 'image/svg+xml',
		mp4: 'video/mp4',
		mp3: 'audio/mpeg',
	};
	return mimeTypes[extension] || 'application/octet-stream';
}
