import type {
	IDataObject,
	IExecuteFunctions,
	ILoadOptionsFunctions,
	INodeExecutionData,
	INodeListSearchResult,
	INodeType,
	INodeTypeDescription,
	JsonObject,
} from 'n8n-workflow';
import { NodeApiError, NodeOperationError } from 'n8n-workflow';

import { fileFields, fileOperations } from './FileDescription';
import { folderFields, folderOperations } from './FolderDescription';
import { getMimeType, microsoftApiRequest, microsoftApiRequestAllItems } from './GenericFunctions';

export class MicrosoftOneDriveBusiness implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Microsoft OneDrive Business',
		name: 'microsoftOneDriveBusiness',
		icon: 'file:onedrive.svg',
		group: ['input'],
		version: 1,
		subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
		description: 'Access Microsoft OneDrive for Business and SharePoint',
		defaults: {
			name: 'Microsoft OneDrive Business',
		},
		inputs: ['main'],
		outputs: ['main'],
		credentials: [
			{
				name: 'microsoftOneDriveBusinessOAuth2Api',
				required: true,
			},
		],
		properties: [
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'File',
						value: 'file',
					},
					{
						name: 'Folder',
						value: 'folder',
					},
				],
				default: 'file',
			},
			...fileOperations,
			...fileFields,
			...folderOperations,
			...folderFields,
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];
		const length = items.length;
		let responseData;
		const resource = this.getNodeParameter('resource', 0);
		const operation = this.getNodeParameter('operation', 0);

		for (let i = 0; i < length; i++) {
			try {
				const driveType = this.getNodeParameter('driveType', i) as string;
				let driveEndpoint = '/me/drive';

				if (driveType === 'user') {
					const userId = this.getNodeParameter('userId', i) as string;
					if (userId) {
						driveEndpoint = `/users/${userId}/drive`;
					}
				} else if (driveType === 'site') {
					const siteIdParam = this.getNodeParameter('siteId', i) as IDataObject | string;
					let siteId: string;

					// Handle resourceLocator format
					if (typeof siteIdParam === 'object' && siteIdParam.mode) {
						const mode = siteIdParam.mode as string;
						const value = siteIdParam.value as string;

						if (mode === 'list' || mode === 'id') {
							// Value is already a site ID
							siteId = value;
						} else if (mode === 'url') {
							// Convert URL to site ID
							// Extract hostname and path from URL
							const url = new URL(value);
							const hostname = url.hostname;
							const path = url.pathname;

							// Get site by URL
							const siteData = await microsoftApiRequest.call(
								this,
								'GET',
								`/sites/${hostname}:${path}`,
							);
							siteId = siteData.id as string;
						} else {
							siteId = value;
						}
					} else {
						// Legacy string format
						siteId = siteIdParam as string;
					}

					driveEndpoint = `/sites/${siteId}/drive`;
				}

				// Helper function to get file ID from path or ID
				const getFileId = async (driveEndpoint: string): Promise<string> => {
					const fileSelection = this.getNodeParameter('fileSelection', i, 'path') as string;
					
					if (fileSelection === 'id') {
						return this.getNodeParameter('fileId', i) as string;
					} else {
						// Get file by path (resourceLocator)
						const filePathParam = this.getNodeParameter('filePath', i) as IDataObject | string;
						
						// Handle resourceLocator format
						if (typeof filePathParam === 'object' && filePathParam.mode) {
							const mode = filePathParam.mode as string;
							const value = filePathParam.value as string;
							
							if (mode === 'list' || mode === 'id') {
								// Value is already an ID
								return value;
							} else {
								// mode === 'path', value is a path
								let filePath = value;
								// Remove leading slash if present
								filePath = filePath.replace(/^\/+/, '');
								
								const encodedPath = filePath.split('/').map(encodeURIComponent).join('/');
								const endpoint = `${driveEndpoint}/root:/${encodedPath}`;
								
								const fileMetadata = await microsoftApiRequest.call(
									this,
									'GET',
									endpoint,
								);
								
								return fileMetadata.id as string;
							}
						} else {
							// Legacy string format
							let filePath = filePathParam as string;
							filePath = filePath.replace(/^\/+/, '');
							
							const encodedPath = filePath.split('/').map(encodeURIComponent).join('/');
							const endpoint = `${driveEndpoint}/root:/${encodedPath}`;
							
							const fileMetadata = await microsoftApiRequest.call(
								this,
								'GET',
								endpoint,
							);
							
							return fileMetadata.id as string;
						}
					}
				};

				if (resource === 'file') {
					if (operation === 'delete') {
						const fileId = await getFileId(driveEndpoint);
						responseData = await microsoftApiRequest.call(
							this,
							'DELETE',
							`${driveEndpoint}/items/${fileId}`,
						);
						responseData = { success: true };
					}

					if (operation === 'download') {
						const fileId = await getFileId(driveEndpoint);
						const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i) as string;

						const fileMetadata = await microsoftApiRequest.call(
							this,
							'GET',
							`${driveEndpoint}/items/${fileId}`,
						);

						const fileName = fileMetadata.name as string;
						const downloadUrl = fileMetadata['@microsoft.graph.downloadUrl'] as string;

						if (!fileMetadata.file) {
							throw new NodeApiError(this.getNode(), fileMetadata as JsonObject, {
								message: 'The ID you provided does not belong to a file.',
							});
						}

						let mimeType = (fileMetadata.file as IDataObject).mimeType as string || 'application/octet-stream';

						// Download file content using the download URL (more reliable)
						let fileBuffer: Buffer;
						
						if (downloadUrl) {
							// Use the direct download URL
							const response = await this.helpers.httpRequest({
								method: 'GET',
								url: downloadUrl,
								encoding: 'arraybuffer',
								json: false,
							});
							fileBuffer = Buffer.from(response as ArrayBuffer);
						} else {
							// Fallback to content endpoint
							const response = await microsoftApiRequest.call(
								this,
								'GET',
								`${driveEndpoint}/items/${fileId}/content`,
								{},
								{},
								undefined,
								{},
								{ encoding: null, resolveWithFullResponse: false },
							);
							fileBuffer = Buffer.isBuffer(response) ? response : Buffer.from(response as ArrayBuffer);
						}

						const newItem: INodeExecutionData = {
							json: fileMetadata,
							binary: {},
						};

						if (items[i].binary !== undefined) {
							Object.assign(newItem.binary!, items[i].binary);
						}

						const data = fileBuffer;

						newItem.binary![binaryPropertyName] = await this.helpers.prepareBinaryData(
							data,
							fileName as string,
							mimeType,
						);

						returnData.push(newItem);
						continue;
					}

					if (operation === 'get') {
						const fileId = await getFileId(driveEndpoint);
						responseData = await microsoftApiRequest.call(
							this,
							'GET',
							`${driveEndpoint}/items/${fileId}`,
						);
					}

					if (operation === 'rename') {
						const fileId = await getFileId(driveEndpoint);
						const newName = this.getNodeParameter('newName', i) as string;
						const body = {
							name: newName,
						};
						responseData = await microsoftApiRequest.call(
							this,
							'PATCH',
							`${driveEndpoint}/items/${fileId}`,
							body,
						);
					}

					if (operation === 'search') {
						const query = this.getNodeParameter('query', i) as string;
						responseData = await microsoftApiRequestAllItems.call(
							this,
							'value',
							'GET',
							`${driveEndpoint}/root/search(q='${query}')`,
						);
						responseData = responseData.filter((item: IDataObject) => item.file);
					}

					if (operation === 'share') {
						const fileId = await getFileId(driveEndpoint);
						const linkType = this.getNodeParameter('linkType', i) as string;
						const linkScope = this.getNodeParameter('linkScope', i) as string;
						const body = {
							type: linkType,
							scope: linkScope,
						};
						responseData = await microsoftApiRequest.call(
							this,
							'POST',
							`${driveEndpoint}/items/${fileId}/createLink`,
							body,
						);
					}

					if (operation === 'upload') {
						// Get parent folder ID (supports resourceLocator)
						const parentIdParam = this.getNodeParameter('parentId', i) as IDataObject | string;
						let parentId: string;

						if (typeof parentIdParam === 'object' && parentIdParam.mode) {
							const mode = parentIdParam.mode as string;
							const value = parentIdParam.value as string;

							if (mode === 'list' || mode === 'id') {
								parentId = value || 'root';
							} else {
								// mode === 'path'
								if (!value || value === '/') {
									parentId = 'root';
								} else {
									// Resolve path to folder ID
									let folderPath = value.replace(/^\/+/, '');
									const encodedPath = folderPath.split('/').map(encodeURIComponent).join('/');
									const folderMetadata = await microsoftApiRequest.call(
										this,
										'GET',
										`${driveEndpoint}/root:/${encodedPath}`,
									);
									parentId = folderMetadata.id as string;
								}
							}
						} else {
							parentId = parentIdParam as string || 'root';
						}

						const fileName = this.getNodeParameter('fileName', i) as string;
						const binaryData = this.getNodeParameter('binaryData', i) as boolean;

						if (binaryData) {
							const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i) as string;
							const binaryDataInfo = this.helpers.assertBinaryData(i, binaryPropertyName);
							const body = await this.helpers.getBinaryDataBuffer(i, binaryPropertyName);

							const uploadFileName = fileName || binaryDataInfo.fileName || 'file';
							const encodedFilename = encodeURIComponent(uploadFileName);
							const mimeType = binaryDataInfo.mimeType || getMimeType(uploadFileName);

							const endpoint =
								parentId === 'root'
									? `${driveEndpoint}/root:/${encodedFilename}:/content`
									: `${driveEndpoint}/items/${parentId}:/${encodedFilename}:/content`;

							responseData = await microsoftApiRequest.call(
								this,
								'PUT',
								endpoint,
								body as unknown as IDataObject,
								{},
								undefined,
								{ 'Content-Type': mimeType, 'Content-Length': body.length.toString() },
								{},
							);

							if (typeof responseData === 'string') {
								responseData = JSON.parse(responseData);
							}
						} else {
							const fileContent = this.getNodeParameter('fileContent', i) as string;
							if (!fileName) {
								throw new NodeOperationError(this.getNode(), 'File name must be set!', {
									itemIndex: i,
								});
							}

							const encodedFilename = encodeURIComponent(fileName);
							const endpoint =
								parentId === 'root'
									? `${driveEndpoint}/root:/${encodedFilename}:/content`
									: `${driveEndpoint}/items/${parentId}:/${encodedFilename}:/content`;

							responseData = await microsoftApiRequest.call(
								this,
								'PUT',
								endpoint,
								fileContent,
								{},
								undefined,
								{ 'Content-Type': 'text/plain' },
							);
						}
					}
				}

				// Helper function to get folder ID from path or ID
				const getFolderId = async (driveEndpoint: string): Promise<string> => {
					const folderSelection = this.getNodeParameter('folderSelection', i, 'path') as string;
					
					if (folderSelection === 'id') {
						return this.getNodeParameter('folderId', i) as string;
					} else {
						// Get folder by path (resourceLocator)
						const folderPathParam = this.getNodeParameter('folderPath', i) as IDataObject | string;
						
						// Handle resourceLocator format
						if (typeof folderPathParam === 'object' && folderPathParam.mode) {
							const mode = folderPathParam.mode as string;
							const value = folderPathParam.value as string;
							
							if (mode === 'list' || mode === 'id') {
								// Value is already an ID
								return value || 'root';
							} else {
								// mode === 'path', value is a path
								let folderPath = value;
								
								// Handle root folder
								if (!folderPath || folderPath === '/') {
									return 'root';
								}
								
								// Remove leading slash if present
								folderPath = folderPath.replace(/^\/+/, '');
								
								const encodedPath = folderPath.split('/').map(encodeURIComponent).join('/');
								const endpoint = `${driveEndpoint}/root:/${encodedPath}`;
								
								const folderMetadata = await microsoftApiRequest.call(
									this,
									'GET',
									endpoint,
								);
								
								return folderMetadata.id as string;
							}
						} else {
							// Legacy string format
							let folderPath = folderPathParam as string;
							
							// Handle root folder
							if (!folderPath || folderPath === '/') {
								return 'root';
							}
							
							// Remove leading slash if present
							folderPath = folderPath.replace(/^\/+/, '');
							
							const encodedPath = folderPath.split('/').map(encodeURIComponent).join('/');
							const endpoint = `${driveEndpoint}/root:/${encodedPath}`;
							
							const folderMetadata = await microsoftApiRequest.call(
								this,
								'GET',
								endpoint,
							);
							
							return folderMetadata.id as string;
						}
					}
				};

				if (resource === 'folder') {
					if (operation === 'create') {
						// Get parent folder ID (supports resourceLocator)
						const parentIdParam = this.getNodeParameter('parentId', i) as IDataObject | string;
						let parentId: string;

						if (typeof parentIdParam === 'object' && parentIdParam.mode) {
							const mode = parentIdParam.mode as string;
							const value = parentIdParam.value as string;

							if (mode === 'list' || mode === 'id') {
								parentId = value || 'root';
							} else {
								// mode === 'path'
								if (!value || value === '/') {
									parentId = 'root';
								} else {
									// Resolve path to folder ID
									let folderPath = value.replace(/^\/+/, '');
									const encodedPath = folderPath.split('/').map(encodeURIComponent).join('/');
									const folderMetadata = await microsoftApiRequest.call(
										this,
										'GET',
										`${driveEndpoint}/root:/${encodedPath}`,
									);
									parentId = folderMetadata.id as string;
								}
							}
						} else {
							parentId = parentIdParam as string || 'root';
						}

						const folderName = this.getNodeParameter('folderName', i) as string;
						const body = {
							name: folderName,
							folder: {},
							'@microsoft.graph.conflictBehavior': 'rename',
						};

						const endpoint =
							parentId === 'root'
								? `${driveEndpoint}/root/children`
								: `${driveEndpoint}/items/${parentId}/children`;

						responseData = await microsoftApiRequest.call(this, 'POST', endpoint, body);
					}

					if (operation === 'delete') {
						const folderId = await getFolderId(driveEndpoint);
						responseData = await microsoftApiRequest.call(
							this,
							'DELETE',
							`${driveEndpoint}/items/${folderId}`,
						);
						responseData = { success: true };
					}

					if (operation === 'getItems') {
						const folderId = await getFolderId(driveEndpoint);
						const endpoint =
							folderId === 'root'
								? `${driveEndpoint}/root/children`
								: `${driveEndpoint}/items/${folderId}/children`;

						responseData = await microsoftApiRequestAllItems.call(this, 'value', 'GET', endpoint);
					}

					if (operation === 'rename') {
						const folderId = await getFolderId(driveEndpoint);
						const newName = this.getNodeParameter('newName', i) as string;
						const body = {
							name: newName,
						};
						responseData = await microsoftApiRequest.call(
							this,
							'PATCH',
							`${driveEndpoint}/items/${folderId}`,
							body,
						);
					}

					if (operation === 'search') {
						const query = this.getNodeParameter('query', i) as string;
						responseData = await microsoftApiRequestAllItems.call(
							this,
							'value',
							'GET',
							`${driveEndpoint}/root/search(q='${query}')`,
						);
						responseData = responseData.filter((item: IDataObject) => item.folder);
					}

					if (operation === 'share') {
						const folderId = await getFolderId(driveEndpoint);
						const linkType = this.getNodeParameter('linkType', i) as string;
						const linkScope = this.getNodeParameter('linkScope', i) as string;
						const body = {
							type: linkType,
							scope: linkScope,
						};
						responseData = await microsoftApiRequest.call(
							this,
							'POST',
							`${driveEndpoint}/items/${folderId}/createLink`,
							body,
						);
					}
				}

				if (Array.isArray(responseData)) {
					returnData.push.apply(returnData, responseData as INodeExecutionData[]);
				} else if (responseData !== undefined) {
					returnData.push({ json: responseData as IDataObject });
				}
			} catch (error) {
				if (this.continueOnFail()) {
					returnData.push({ json: { error: error.message } });
					continue;
				}
				throw error;
			}
		}

		return [returnData];
	}

	methods = {
		listSearch: {
			async searchFiles(
				this: ILoadOptionsFunctions,
				filter?: string,
			): Promise<INodeListSearchResult> {
				const driveType = this.getNodeParameter('driveType', 0) as string;
				let driveEndpoint = '/me/drive';

				if (driveType === 'user') {
					const userId = this.getNodeParameter('userId', 0) as string;
					if (userId) {
						driveEndpoint = `/users/${userId}/drive`;
					}
				} else if (driveType === 'site') {
					const siteIdParam = this.getNodeParameter('siteId', 0) as IDataObject | string;
					let siteId: string;

					if (typeof siteIdParam === 'object' && siteIdParam.value) {
						siteId = siteIdParam.value as string;
					} else {
						siteId = siteIdParam as string;
					}

					if (siteId) {
						driveEndpoint = `/sites/${siteId}/drive`;
					}
				}

				let items: IDataObject[] = [];

				if (filter) {
					// Search for files matching the filter
					items = await microsoftApiRequestAllItems.call(
						this,
						'value',
						'GET',
						`${driveEndpoint}/root/search(q='${filter}')`,
					);
					// Filter to only files
					items = items.filter((item: IDataObject) => item.file);
				} else {
					// List recent files
					const response = await microsoftApiRequest.call(
						this,
						'GET',
						`${driveEndpoint}/recent`,
					) as IDataObject;
					items = ((response.value as IDataObject[]) || []).filter((item: IDataObject) => item.file);
				}

				return {
					results: items.slice(0, 100).map((item: IDataObject) => ({
						name: item.name as string,
						value: item.id as string,
						url: item.webUrl as string,
					})),
				};
			},

			async searchFolders(
				this: ILoadOptionsFunctions,
				filter?: string,
			): Promise<INodeListSearchResult> {
				const driveType = this.getNodeParameter('driveType', 0) as string;
				let driveEndpoint = '/me/drive';

				if (driveType === 'user') {
					const userId = this.getNodeParameter('userId', 0) as string;
					if (userId) {
						driveEndpoint = `/users/${userId}/drive`;
					}
				} else if (driveType === 'site') {
					const siteIdParam = this.getNodeParameter('siteId', 0) as IDataObject | string;
					let siteId: string;

					if (typeof siteIdParam === 'object' && siteIdParam.value) {
						siteId = siteIdParam.value as string;
					} else {
						siteId = siteIdParam as string;
					}

					if (siteId) {
						driveEndpoint = `/sites/${siteId}/drive`;
					}
				}

				let items: IDataObject[] = [];

				if (filter) {
					// Search for folders matching the filter
					items = await microsoftApiRequestAllItems.call(
						this,
						'value',
						'GET',
						`${driveEndpoint}/root/search(q='${filter}')`,
					);
					// Filter to only folders
					items = items.filter((item: IDataObject) => item.folder);
				} else {
					// List root folder items
					items = await microsoftApiRequestAllItems.call(
						this,
						'value',
						'GET',
						`${driveEndpoint}/root/children`,
					);
					// Filter to only folders
					items = items.filter((item: IDataObject) => item.folder);
				}

				// Always include root folder as an option
				const results = [
					{
						name: '/ (Root)',
						value: 'root',
						url: '',
					},
					...items.slice(0, 100).map((item: IDataObject) => ({
						name: item.name as string,
						value: item.id as string,
						url: item.webUrl as string,
					})),
				];

				return { results };
			},

			async searchSites(
				this: ILoadOptionsFunctions,
				filter?: string,
			): Promise<INodeListSearchResult> {
				let items: IDataObject[] = [];

				if (filter) {
					// Search for sites matching the filter
					const response = await microsoftApiRequest.call(
						this,
						'GET',
						`/sites?search=${encodeURIComponent(filter)}`,
					) as IDataObject;
					items = (response.value as IDataObject[]) || [];
				} else {
					// List sites the user has access to
					const response = await microsoftApiRequest.call(
						this,
						'GET',
						'/sites?search=*',
					) as IDataObject;
					items = ((response.value as IDataObject[]) || []).slice(0, 50);
				}

				return {
					results: items.map((item: IDataObject) => ({
						name: `${item.displayName || item.name} (${item.webUrl})` as string,
						value: item.id as string,
						url: item.webUrl as string,
					})),
				};
			},
		},
	};
}
