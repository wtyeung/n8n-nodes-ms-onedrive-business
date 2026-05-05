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
import { excelFields, excelOperations } from './ExcelDescription';
import { getMimeType, microsoftApiRequest, microsoftApiRequestAllItems } from './GenericFunctions';

async function getDriveEndpointForLoadOptions(this: ILoadOptionsFunctions): Promise<string> {
	const driveType = this.getNodeParameter('driveType', 0) as string;
	let driveEndpoint = '/me/drive';

	if (driveType === 'user') {
		const userId = this.getNodeParameter('userId', 0) as string;
		if (userId) driveEndpoint = `/users/${userId}/drive`;
	} else if (driveType === 'site') {
		const siteIdParam = this.getNodeParameter('siteId', 0) as IDataObject | string;
		let siteId: string;

		if (typeof siteIdParam === 'object' && siteIdParam.mode) {
			const mode = siteIdParam.mode as string;
			const value = siteIdParam.value as string;
			if (mode === 'url') {
				const url = new URL(value);
				const siteData = await microsoftApiRequest.call(this, 'GET', `/sites/${url.hostname}:${url.pathname}`) as IDataObject;
				siteId = siteData.id as string;
			} else {
				siteId = value;
			}
		} else {
			siteId = siteIdParam as string;
		}

		if (!siteId) {
			throw new NodeOperationError(this.getNode(), 'Please select a SharePoint Site before loading folders.');
		}
		driveEndpoint = `/sites/${siteId}/drive`;
	}

	return driveEndpoint;
}

export class MicrosoftOneDriveBusiness implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'MS OneDrive Business',
		name: 'microsoftOneDriveBusiness',
		icon: 'file:onedrive.svg',
		group: ['input'],
		version: 1,
		subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
		description: 'Access MS OneDrive for Business and SharePoint',
		defaults: {
			name: 'MS OneDrive Business',
		},
		usableAsTool: true,
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
						name: 'Excel',
						value: 'excel',
					},
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
			...excelOperations,
			...excelFields,
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
					const fileSelection = this.getNodeParameter('fileSelection', i, 'browse') as string;
					
					if (fileSelection === 'browse') {
						// Find first level where user selected a file (value starts with 'file:')
						const levels = ['browseFolder1', 'browseFolder2', 'browseFolder3', 'browseFolder4', 'browseFolder5'];
						for (const level of levels) {
							const val = this.getNodeParameter(level, i, '') as string;
							if (val.startsWith('file:')) {
								return val.replace('file:', '');
							}
						}
						throw new NodeOperationError(this.getNode(), 'No file selected. Please select a 📄 file in one of the browse levels.');

					} else if (fileSelection === 'id') {
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

				// Helper function to get folder ID from path or ID
				const getFolderId = async (driveEndpoint: string): Promise<string> => {
					const folderSelection = this.getNodeParameter('folderSelection', i, 'browse') as string;

					if (folderSelection === 'browse') {
						// Walk levels F1→F5; last level with a folder: value is the effective folder
						const levels = ['browseFolderF1', 'browseFolderF2', 'browseFolderF3', 'browseFolderF4', 'browseFolderF5'];
						let folderId = '';
						for (const level of levels) {
							const val = this.getNodeParameter(level, i, '') as string;
							if (val && val.startsWith('folder:')) {
								folderId = val.replace('folder:', '');
							}
						}
						if (!folderId) {
							throw new NodeOperationError(this.getNode(), 'No folder selected. Please select a folder in the browse levels.');
						}
						return folderId;
					}

					if (folderSelection === 'id') {
						return this.getNodeParameter('folderId', i) as string;
					} else {
						// Get folder by path (resourceLocator)
						const folderPathParam = this.getNodeParameter('folderPath', i) as IDataObject | string;

						if (typeof folderPathParam === 'object' && folderPathParam.mode) {
							const mode = folderPathParam.mode as string;
							const value = folderPathParam.value as string;

							if (mode === 'list' || mode === 'id') {
								return value || 'root';
							} else {
								let folderPath = value;
								if (!folderPath || folderPath === '/') return 'root';
								folderPath = folderPath.replace(/^\/+/, '');
								const encodedPath = folderPath.split('/').map(encodeURIComponent).join('/');
								const folderMetadata = await microsoftApiRequest.call(this, 'GET', `${driveEndpoint}/root:/${encodedPath}`);
								return folderMetadata.id as string;
							}
						} else {
							let folderPath = folderPathParam as string;
							if (!folderPath || folderPath === '/') return 'root';
							folderPath = folderPath.replace(/^\/+/, '');
							const encodedPath = folderPath.split('/').map(encodeURIComponent).join('/');
							const folderMetadata = await microsoftApiRequest.call(this, 'GET', `${driveEndpoint}/root:/${encodedPath}`);
							return folderMetadata.id as string;
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

						const mimeType = (fileMetadata.file as IDataObject).mimeType as string || 'application/octet-stream';

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
						const uploadFolderSelection = this.getNodeParameter('folderSelection', i, 'browse') as string;
						let parentId: string;

						if (uploadFolderSelection === 'browse') {
							parentId = await getFolderId(driveEndpoint);
						} else {
							// Get parent folder ID via path or ID (resourceLocator)
							const parentIdParam = this.getNodeParameter('parentId', i) as IDataObject | string;

							if (typeof parentIdParam === 'object' && parentIdParam.mode) {
								const mode = parentIdParam.mode as string;
								const value = parentIdParam.value as string;

								if (mode === 'id') {
									parentId = value || 'root';
								} else {
									// mode === 'path'
									if (!value || value === '/') {
										parentId = 'root';
									} else {
										const folderPath = value.replace(/^\/+/, '');
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

				if (resource === 'folder') {
					if (operation === 'create') {
						const folderSelectionCreate = this.getNodeParameter('folderSelection', i, 'browse') as string;
						let parentId: string;

						if (folderSelectionCreate === 'browse') {
							parentId = await getFolderId(driveEndpoint);
						} else {
							// Get parent folder ID via path or ID (resourceLocator)
							const parentIdParam = this.getNodeParameter('parentId', i) as IDataObject | string;

							if (typeof parentIdParam === 'object' && parentIdParam.mode) {
								const mode = parentIdParam.mode as string;
								const value = parentIdParam.value as string;

								if (mode === 'id') {
									parentId = value || 'root';
								} else {
									// mode === 'path'
									if (!value || value === '/') {
										parentId = 'root';
									} else {
										const folderPath = value.replace(/^\/+/, '');
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

				if (resource === 'excel') {
					const workbookId = await getFileId(driveEndpoint);
					const worksheet = this.getNodeParameter('worksheet', i) as string;

					if (operation === 'readRows') {
						const useRange = this.getNodeParameter('useRange', i, false) as boolean;
						const options = this.getNodeParameter('options', i, {}) as IDataObject;
						const rawData = (options.rawData as boolean) || false;
						let endpoint: string;
						if (useRange) {
							const range = this.getNodeParameter('range', i, '') as string;
							endpoint = `${driveEndpoint}/items/${workbookId}/workbook/worksheets/${worksheet}/range(address='${range}')`;
						} else {
							endpoint = `${driveEndpoint}/items/${workbookId}/workbook/worksheets/${worksheet}/usedRange`;
						}
						responseData = await microsoftApiRequest.call(this, 'GET', endpoint);
						if (!rawData && responseData?.values) {
							const keyRow = this.getNodeParameter('keyRow', i, 0) as number;
							const dataStartRow = this.getNodeParameter('dataStartRow', i, 1) as number;
							const values = responseData.values as (string | number | boolean)[][];
							const headers = values[keyRow] || [];
							const rows = values.slice(dataStartRow).map((row) => {
								const obj: IDataObject = {};
								headers.forEach((header, idx) => { obj[String(header)] = row[idx]; });
								return { json: obj };
							});
							returnData.push(...rows);
							continue;
						}

					} else if (operation === 'appendOrUpdate') {
						const dataMode = this.getNodeParameter('dataMode', i) as string;
						const columnToMatchOn = this.getNodeParameter('columnToMatchOn', i, '') as string;
						const usedRangeData = await microsoftApiRequest.call(this, 'GET', `${driveEndpoint}/items/${workbookId}/workbook/worksheets/${worksheet}/usedRange`);
						const values = usedRangeData.values as (string | number | boolean)[][];
						const headers = values[0] || [];
						let newRow: (string | number | boolean)[];
						if (dataMode === 'autoMap') {
							const inputData = items[i].json;
							newRow = headers.map((h) => (inputData[String(h)] !== undefined ? String(inputData[String(h)]) : ''));
						} else {
							const fieldValues = ((this.getNodeParameter('fieldsUi', i, {}) as IDataObject).fieldValues as IDataObject[]) || [];
							newRow = headers.map((h) => { const f = fieldValues.find((x) => x.column === String(h)); return f ? String(f.fieldValue) : ''; });
						}
						let rowToUpdate = -1;
						if (columnToMatchOn) {
							const valueToMatchOn = this.getNodeParameter('valueToMatchOn', i) as string;
							const colIdx = headers.findIndex((h) => String(h) === columnToMatchOn);
							if (colIdx !== -1) {
								for (let r = 1; r < values.length; r++) {
									if (String(values[r][colIdx]) === valueToMatchOn) { rowToUpdate = r; break; }
								}
							}
						}
						const rowNum = rowToUpdate !== -1 ? rowToUpdate + 1 : values.length + 1;
						const colLetter = String.fromCharCode(65 + headers.length - 1);
						const range = `A${rowNum}:${colLetter}${rowNum}`;
						responseData = await microsoftApiRequest.call(this, 'PATCH', `${driveEndpoint}/items/${workbookId}/workbook/worksheets/${worksheet}/range(address='${range}')`, { values: [newRow] });

					} else if (operation === 'deleteRows') {
						const deleteRange = this.getNodeParameter('deleteRange', i) as string;
						responseData = await microsoftApiRequest.call(this, 'POST', `${driveEndpoint}/items/${workbookId}/workbook/worksheets/${worksheet}/range(address='${deleteRange}')/delete`, { shift: 'Up' });
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
		loadOptions: {
			async getBrowseLevel1(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				const driveEndpoint = await getDriveEndpointForLoadOptions.call(this);

				const allItems = await microsoftApiRequestAllItems.call(this, 'value', 'GET',
					`${driveEndpoint}/root/children?$select=id,name,folder,file`,
				);

				const results: Array<{ name: string; value: string }> = [];
				for (const item of allItems as IDataObject[]) {
					if (item.folder) {
						results.push({ name: `▶ ${item.name as string}`, value: `folder:${item.id as string}` });
					} else if (item.file) {
						results.push({ name: `📄 ${item.name as string}`, value: `file:${item.id as string}` });
					}
				}
				return results;
			},

			async getBrowseLevel2(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				const parentVal = this.getNodeParameter('browseFolder1', 0) as string;
				if (!parentVal || !parentVal.startsWith('folder:')) {
					return [{ name: '— Select a Folder at Level 1 First —', value: '__done__' }];
				}
				const parentId = parentVal.replace('folder:', '');
				const driveEndpoint = await getDriveEndpointForLoadOptions.call(this);

				const allItems = await microsoftApiRequestAllItems.call(this, 'value', 'GET',
					`${driveEndpoint}/items/${parentId}/children?$select=id,name,folder,file`,
				);
				const results: Array<{ name: string; value: string }> = [];
				for (const item of allItems as IDataObject[]) {
					if (item.folder) results.push({ name: `▶ ${item.name as string}`, value: `folder:${item.id as string}` });
					else if (item.file) results.push({ name: `📄 ${item.name as string}`, value: `file:${item.id as string}` });
				}
				return results;
			},

			async getBrowseLevel3(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				const parentVal = this.getNodeParameter('browseFolder2', 0) as string;
				if (!parentVal || !parentVal.startsWith('folder:')) {
					return [{ name: '— Select a Folder at Level 2 First —', value: '__done__' }];
				}
				const parentId = parentVal.replace('folder:', '');
				const driveEndpoint = await getDriveEndpointForLoadOptions.call(this);

				const allItems = await microsoftApiRequestAllItems.call(this, 'value', 'GET',
					`${driveEndpoint}/items/${parentId}/children?$select=id,name,folder,file`,
				);
				const results: Array<{ name: string; value: string }> = [];
				for (const item of allItems as IDataObject[]) {
					if (item.folder) results.push({ name: `▶ ${item.name as string}`, value: `folder:${item.id as string}` });
					else if (item.file) results.push({ name: `📄 ${item.name as string}`, value: `file:${item.id as string}` });
				}
				return results;
			},

			async getBrowseLevel4(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				const parentVal = this.getNodeParameter('browseFolder3', 0) as string;
				if (!parentVal || !parentVal.startsWith('folder:')) {
					return [{ name: '— Select a Folder at Level 3 First —', value: '__done__' }];
				}
				const parentId = parentVal.replace('folder:', '');
				const driveEndpoint = await getDriveEndpointForLoadOptions.call(this);

				const allItems = await microsoftApiRequestAllItems.call(this, 'value', 'GET',
					`${driveEndpoint}/items/${parentId}/children?$select=id,name,folder,file`,
				);
				const results: Array<{ name: string; value: string }> = [];
				for (const item of allItems as IDataObject[]) {
					if (item.folder) results.push({ name: `▶ ${item.name as string}`, value: `folder:${item.id as string}` });
					else if (item.file) results.push({ name: `📄 ${item.name as string}`, value: `file:${item.id as string}` });
				}
				return results;
			},

			async getBrowseLevel5(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				const parentVal = this.getNodeParameter('browseFolder4', 0) as string;
				if (!parentVal || !parentVal.startsWith('folder:')) {
					return [{ name: '— Select a Folder at Level 4 First —', value: '__done__' }];
				}
				const parentId = parentVal.replace('folder:', '');
				const driveEndpoint = await getDriveEndpointForLoadOptions.call(this);

				const allItems = await microsoftApiRequestAllItems.call(this, 'value', 'GET',
					`${driveEndpoint}/items/${parentId}/children?$select=id,name,folder,file`,
				);
				const results: Array<{ name: string; value: string }> = [];
				for (const item of allItems as IDataObject[]) {
					if (item.folder) results.push({ name: `▶ ${item.name as string}`, value: `folder:${item.id as string}` });
					else if (item.file) results.push({ name: `📄 ${item.name as string}`, value: `file:${item.id as string}` });
				}
				return results;
			},

			async getBrowseFolderLevel1(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				const driveEndpoint = await getDriveEndpointForLoadOptions.call(this);
				const allItems = await microsoftApiRequestAllItems.call(this, 'value', 'GET',
					`${driveEndpoint}/root/children?$select=id,name,folder`,
				);
				return [
					{ name: '▶ / (Root)', value: 'folder:root' },
					...(allItems as IDataObject[])
						.filter((item) => item.folder)
						.map((item) => ({ name: `▶ ${item.name as string}`, value: `folder:${item.id as string}` })),
				];
			},

			async getBrowseFolderLevel2(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				const parentVal = this.getNodeParameter('browseFolderF1', 0) as string;
				if (!parentVal || !parentVal.startsWith('folder:')) {
					return [{ name: '— Select a Folder at Level 1 First —', value: '__stop__' }];
				}
				const parentId = parentVal.replace('folder:', '');
				const driveEndpoint = await getDriveEndpointForLoadOptions.call(this);
				const endpoint = parentId === 'root'
					? `${driveEndpoint}/root/children?$select=id,name,folder`
					: `${driveEndpoint}/items/${parentId}/children?$select=id,name,folder`;
				const allItems = await microsoftApiRequestAllItems.call(this, 'value', 'GET', endpoint);
				return [
					{ name: '— Use Level 1 Folder —', value: '__stop__' },
					...(allItems as IDataObject[])
						.filter((item) => item.folder)
						.map((item) => ({ name: `▶ ${item.name as string}`, value: `folder:${item.id as string}` })),
				];
			},

			async getBrowseFolderLevel3(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				const parentVal = this.getNodeParameter('browseFolderF2', 0) as string;
				if (!parentVal || !parentVal.startsWith('folder:')) {
					return [{ name: '— Select a Folder at Level 2 First —', value: '__stop__' }];
				}
				const parentId = parentVal.replace('folder:', '');
				const driveEndpoint = await getDriveEndpointForLoadOptions.call(this);
				const endpoint = parentId === 'root'
					? `${driveEndpoint}/root/children?$select=id,name,folder`
					: `${driveEndpoint}/items/${parentId}/children?$select=id,name,folder`;
				const allItems = await microsoftApiRequestAllItems.call(this, 'value', 'GET', endpoint);
				return [
					{ name: '— Use Level 2 Folder —', value: '__stop__' },
					...(allItems as IDataObject[])
						.filter((item) => item.folder)
						.map((item) => ({ name: `▶ ${item.name as string}`, value: `folder:${item.id as string}` })),
				];
			},

			async getBrowseFolderLevel4(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				const parentVal = this.getNodeParameter('browseFolderF3', 0) as string;
				if (!parentVal || !parentVal.startsWith('folder:')) {
					return [{ name: '— Select a Folder at Level 3 First —', value: '__stop__' }];
				}
				const parentId = parentVal.replace('folder:', '');
				const driveEndpoint = await getDriveEndpointForLoadOptions.call(this);
				const endpoint = parentId === 'root'
					? `${driveEndpoint}/root/children?$select=id,name,folder`
					: `${driveEndpoint}/items/${parentId}/children?$select=id,name,folder`;
				const allItems = await microsoftApiRequestAllItems.call(this, 'value', 'GET', endpoint);
				return [
					{ name: '— Use Level 3 Folder —', value: '__stop__' },
					...(allItems as IDataObject[])
						.filter((item) => item.folder)
						.map((item) => ({ name: `▶ ${item.name as string}`, value: `folder:${item.id as string}` })),
				];
			},

			async getBrowseFolderLevel5(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				const parentVal = this.getNodeParameter('browseFolderF4', 0) as string;
				if (!parentVal || !parentVal.startsWith('folder:')) {
					return [{ name: '— Select a Folder at Level 4 First —', value: '__stop__' }];
				}
				const parentId = parentVal.replace('folder:', '');
				const driveEndpoint = await getDriveEndpointForLoadOptions.call(this);
				const endpoint = parentId === 'root'
					? `${driveEndpoint}/root/children?$select=id,name,folder`
					: `${driveEndpoint}/items/${parentId}/children?$select=id,name,folder`;
				const allItems = await microsoftApiRequestAllItems.call(this, 'value', 'GET', endpoint);
				return [
					{ name: '— Use Level 4 Folder —', value: '__stop__' },
					...(allItems as IDataObject[])
						.filter((item) => item.folder)
						.map((item) => ({ name: `▶ ${item.name as string}`, value: `folder:${item.id as string}` })),
				];
			},

			async getWorksheets(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				try {
					const driveEndpoint = await getDriveEndpointForLoadOptions.call(this);
					const fileSelection = this.getNodeParameter('fileSelection', 0) as string || 'browse';
					let workbookId = '';

					if (fileSelection === 'browse') {
						for (const level of ['browseFolder1', 'browseFolder2', 'browseFolder3', 'browseFolder4', 'browseFolder5']) {
							try {
								const val = this.getNodeParameter(level, 0) as string;
								if (val && val.startsWith('file:')) {
									workbookId = val.replace('file:', '');
									break;
								}
							} catch { continue; }
						}
					} else if (fileSelection === 'id') {
						workbookId = this.getNodeParameter('fileId', 0) as string;
					} else {
						const filePathParam = this.getNodeParameter('filePath', 0) as IDataObject;
						const mode = filePathParam.mode as string;
						const value = filePathParam.value as string;
						if (mode === 'list' || mode === 'id') {
							workbookId = value;
						} else {
							const filePath = (value as string).replace(/^\/+/, '');
							const encodedPath = filePath.split('/').map(encodeURIComponent).join('/');
							const meta = await microsoftApiRequest.call(this, 'GET', `${driveEndpoint}/root:/${encodedPath}`);
							workbookId = meta.id as string;
						}
					}

					if (!workbookId) return [{ name: '— Select an Excel File First —', value: '' }];

					const response = await microsoftApiRequest.call(this, 'GET', `${driveEndpoint}/items/${workbookId}/workbook/worksheets`);
					if (response?.value) {
						return (response.value as IDataObject[]).map((ws) => ({ name: ws.name as string, value: ws.name as string }));
					}
					return [{ name: '— No Worksheets Found —', value: '' }];
				} catch (error) {
					return [{ name: `— Error: ${error instanceof Error ? error.message : 'Unknown'} —`, value: '' }];
				}
			},

			async getColumns(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				try {
					const driveEndpoint = await getDriveEndpointForLoadOptions.call(this);
					const fileSelection = this.getNodeParameter('fileSelection', 0) as string || 'browse';
					let workbookId = '';

					if (fileSelection === 'browse') {
						for (const level of ['browseFolder1', 'browseFolder2', 'browseFolder3', 'browseFolder4', 'browseFolder5']) {
							try {
								const val = this.getNodeParameter(level, 0) as string;
								if (val && val.startsWith('file:')) {
									workbookId = val.replace('file:', '');
									break;
								}
							} catch { continue; }
						}
					} else if (fileSelection === 'id') {
						workbookId = this.getNodeParameter('fileId', 0) as string;
					} else {
						const filePathParam = this.getNodeParameter('filePath', 0) as IDataObject;
						const mode = filePathParam.mode as string;
						const value = filePathParam.value as string;
						if (mode === 'list' || mode === 'id') {
							workbookId = value;
						} else {
							const filePath = (value as string).replace(/^\/+/, '');
							const encodedPath = filePath.split('/').map(encodeURIComponent).join('/');
							const meta = await microsoftApiRequest.call(this, 'GET', `${driveEndpoint}/root:/${encodedPath}`);
							workbookId = meta.id as string;
						}
					}

					if (!workbookId) return [{ name: '— Select an Excel File First —', value: '' }];

					const worksheet = this.getNodeParameter('worksheet', 0) as string;
					if (!worksheet) return [{ name: '— Select a Worksheet First —', value: '' }];

					const usedRange = await microsoftApiRequest.call(this, 'GET', `${driveEndpoint}/items/${workbookId}/workbook/worksheets/${worksheet}/usedRange`);
					if (usedRange?.values?.length > 0) {
						return (usedRange.values[0] as (string | number | boolean)[]).map((h) => ({ name: String(h), value: String(h) }));
					}
					return [{ name: '— No Columns Found —', value: '' }];
				} catch (error) {
					return [{ name: `— Error: ${error instanceof Error ? error.message : 'Unknown'} —`, value: '' }];
				}
			},
		},

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
