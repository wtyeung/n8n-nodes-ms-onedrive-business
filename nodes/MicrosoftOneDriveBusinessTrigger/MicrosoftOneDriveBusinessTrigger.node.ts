import type {
	IDataObject,
	IHookFunctions,
	ILoadOptionsFunctions,
	INodeType,
	INodeTypeDescription,
	IWebhookFunctions,
	IWebhookResponseData,
} from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';

import {
	microsoftApiRequest,
	microsoftApiRequestAllItems,
} from '../MicrosoftOneDriveBusiness/GenericFunctions';

interface IStateData {
	deltaLink?: string;
	processedVersions: { [key: string]: boolean };
	lastKnownItems: { [key: string]: { eTag: string; lastModifiedDateTime: string } };
}

export class MicrosoftOneDriveBusinessTrigger implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'MS OneDrive Business Trigger',
		name: 'microsoftOneDriveBusinessTrigger',
		icon: 'file:onedrive.svg',
		group: ['trigger'],
		version: 1,
		subtitle: '={{$parameter["event"]}}',
		description: 'Triggers when files or folders are created or updated in OneDrive Business',
		defaults: {
			name: 'OneDrive Business Trigger',
		},
		inputs: [],
		outputs: ['main'],
		credentials: [
			{
				name: 'microsoftOneDriveBusinessOAuth2Api',
				required: true,
			},
		],
		webhooks: [
			{
				name: 'default',
				httpMethod: 'POST',
				responseMode: 'onReceived',
				path: 'webhook',
			},
		],
		usableAsTool: true,
		properties: [
			{
				displayName: 'Event',
				name: 'event',
				type: 'options',
				options: [
					{
						name: 'File Created',
						value: 'file.created',
						description: 'Fires when a new file is uploaded to the watched folder. Select a folder below.',
					},
					{
						name: 'File Updated',
						value: 'file.updated',
						description: 'Fires when a file is renamed or its content changes. Select a folder (all files inside) or a specific file below.',
					},
					{
						name: 'Folder Created',
						value: 'folder.created',
						description: 'Fires when a new subfolder is created inside the watched folder. Select the parent folder below.',
					},
					{
						name: 'Folder Updated',
						value: 'folder.updated',
						description: 'Fires when a subfolder is renamed or moved. Does NOT fire for file changes inside the folder — use File Updated for that.',
					},
				],
				default: 'file.created',
				required: true,
			},
			{
				displayName: 'Drive Type',
				name: 'driveType',
				type: 'options',
				options: [
					{
						name: 'User Drive',
						value: 'user',
						description: 'Monitor user OneDrive',
					},
					{
						name: 'SharePoint Site Drive',
						value: 'site',
						description: 'Monitor SharePoint site drive',
					},
				],
				default: 'user',
			},
			{
				displayName: 'User ID',
				name: 'userId',
				type: 'string',
				displayOptions: {
					show: {
						driveType: ['user'],
					},
				},
				default: '',
				placeholder: 'user@domain.com',
				description: 'User email or ID. Leave empty to use the authenticated user.',
			},
			{
				displayName: 'Site ID',
				name: 'siteId',
				type: 'string',
				displayOptions: {
					show: {
						driveType: ['site'],
					},
				},
				default: '',
				required: true,
				placeholder:
					'contoso.sharepoint.com,da60e844-ba1d-49bc-b4d4-d5e36bae9019,712a596e-90a1-49e3-9b48-bfa80bee8740',
				description: 'SharePoint site ID',
			},
			{
				displayName: 'Folder to Watch',
				name: 'folderSelection',
				type: 'options',
				options: [
					{
						name: 'Browse',
						value: 'browse',
						description: 'Navigate folders step-by-step to select the folder to watch',
					},
					{
						name: 'Entire Drive (Root)',
						value: 'root',
						description: 'Watch all changes across the entire drive',
					},
					{
						name: 'By Folder ID',
						value: 'id',
						description: 'Specify the folder by its OneDrive item ID',
					},
				],
				default: 'root',
				description: 'Select the folder to watch. For File/Folder Created events, select the parent folder where new items will appear. For File Updated, select a folder (watches all files inside) or a specific file. For Folder Updated, select the parent folder containing the subfolders to monitor.',
			},
			{
				displayName: 'Level 1 Name or ID',
				name: 'watchFolder1',
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getWatchLevel1',
					loadOptionsDependsOn: ['driveType', 'userId', 'siteId', 'event'],
				},
				displayOptions: {
					show: { folderSelection: ['browse'] },
				},
				default: '',
				required: true,
				description: 'Select a ▶ folder to go deeper, or select the target folder directly. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
			},
			{
				displayName: 'Level 2 Name or ID',
				name: 'watchFolder2',
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getWatchLevel2',
					loadOptionsDependsOn: ['driveType', 'userId', 'siteId', 'event', 'watchFolder1'],
				},
				displayOptions: {
					show: { folderSelection: ['browse'] },
				},
				default: '__stop__',
				description: 'Select a subfolder to go deeper, or leave as is to use Level 1 as the target. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
			},
			{
				displayName: 'Level 3 Name or ID',
				name: 'watchFolder3',
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getWatchLevel3',
					loadOptionsDependsOn: ['driveType', 'userId', 'siteId', 'event', 'watchFolder1', 'watchFolder2'],
				},
				displayOptions: {
					show: { folderSelection: ['browse'] },
				},
				default: '__stop__',
				description: 'Select a subfolder to go deeper, or leave as is to use Level 2 as the target. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
			},
			{
				displayName: 'Level 4 Name or ID',
				name: 'watchFolder4',
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getWatchLevel4',
					loadOptionsDependsOn: ['driveType', 'userId', 'siteId', 'event', 'watchFolder1', 'watchFolder2', 'watchFolder3'],
				},
				displayOptions: {
					show: { folderSelection: ['browse'] },
				},
				default: '__stop__',
				description: 'Select a subfolder to go deeper, or leave as is to use Level 3 as the target. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
			},
			{
				displayName: 'Level 5 Name or ID',
				name: 'watchFolder5',
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getWatchLevel5',
					loadOptionsDependsOn: ['driveType', 'userId', 'siteId', 'event', 'watchFolder1', 'watchFolder2', 'watchFolder3', 'watchFolder4'],
				},
				displayOptions: {
					show: { folderSelection: ['browse'] },
				},
				default: '__stop__',
				description: 'Select the deepest target folder, or leave as is to use Level 4 as the target. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
			},
			{
				displayName: 'Folder ID',
				name: 'watchFolderId',
				type: 'string',
				displayOptions: {
					show: { folderSelection: ['id'] },
				},
				default: '',
				required: true,
				placeholder: 'ABC123DEF456...',
				description: 'The unique OneDrive item ID of the folder to watch',
			},
		],
	};

	methods = {
		loadOptions: {
			async getWatchLevel1(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				const event = this.getNodeParameter('event', 0) as string;
				const showFiles = event === 'file.updated';
				const driveType = this.getNodeParameter('driveType') as string;
				let driveEndpoint = '/me/drive';
				if (driveType === 'user') {
					const userId = this.getNodeParameter('userId', 0) as string;
					if (userId) driveEndpoint = `/users/${userId}/drive`;
				} else if (driveType === 'site') {
					const siteId = this.getNodeParameter('siteId', 0) as string;
					driveEndpoint = `/sites/${siteId}/drive`;
				}
				const allItems = await microsoftApiRequestAllItems.call(this, 'value', 'GET',
					`${driveEndpoint}/root/children?$select=id,name,folder,file`,
				);
				return (allItems as IDataObject[])
					.filter((item) => showFiles || !!item.folder)
					.map((item) =>
						item.folder
							? { name: `▶ ${item.name as string}`, value: `folder:${item.id as string}` }
							: { name: item.name as string, value: `file:${item.id as string}` },
					);
			},

			async getWatchLevel2(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				const parentVal = this.getNodeParameter('watchFolder1', 0) as string;
				if (!parentVal || parentVal === '__stop__') return [{ name: '— Select a Folder at Level 1 First —', value: '__stop__' }];
				if (parentVal.startsWith('file:')) return [{ name: '— File Selected at Level 1 — No Subfolders —', value: '__stop__' }];
				const event = this.getNodeParameter('event', 0) as string;
				const showFiles = event === 'file.updated';
				const driveType = this.getNodeParameter('driveType') as string;
				let driveEndpoint = '/me/drive';
				if (driveType === 'user') {
					const userId = this.getNodeParameter('userId', 0) as string;
					if (userId) driveEndpoint = `/users/${userId}/drive`;
				} else if (driveType === 'site') {
					const siteId = this.getNodeParameter('siteId', 0) as string;
					driveEndpoint = `/sites/${siteId}/drive`;
				}
				const parentId = parentVal.replace('folder:', '');
				const allItems = await microsoftApiRequestAllItems.call(this, 'value', 'GET',
					`${driveEndpoint}/items/${parentId}/children?$select=id,name,folder,file`,
				);
				return [
					{ name: '— Use Level 1 Folder —', value: '__stop__' },
					...(allItems as IDataObject[])
						.filter((item) => showFiles || !!item.folder)
						.map((item) =>
							item.folder
								? { name: `▶ ${item.name as string}`, value: `folder:${item.id as string}` }
								: { name: item.name as string, value: `file:${item.id as string}` },
						),
				];
			},

			async getWatchLevel3(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				const parentVal = this.getNodeParameter('watchFolder2', 0) as string;
				if (!parentVal || parentVal === '__stop__') return [{ name: '— Select a Folder at Level 2 First —', value: '__stop__' }];
				if (parentVal.startsWith('file:')) return [{ name: '— File Selected at Level 2 — No Subfolders —', value: '__stop__' }];
				const event = this.getNodeParameter('event', 0) as string;
				const showFiles = event === 'file.updated';
				const driveType = this.getNodeParameter('driveType') as string;
				let driveEndpoint = '/me/drive';
				if (driveType === 'user') {
					const userId = this.getNodeParameter('userId', 0) as string;
					if (userId) driveEndpoint = `/users/${userId}/drive`;
				} else if (driveType === 'site') {
					const siteId = this.getNodeParameter('siteId', 0) as string;
					driveEndpoint = `/sites/${siteId}/drive`;
				}
				const parentId = parentVal.replace('folder:', '');
				const allItems = await microsoftApiRequestAllItems.call(this, 'value', 'GET',
					`${driveEndpoint}/items/${parentId}/children?$select=id,name,folder,file`,
				);
				return [
					{ name: '— Use Level 2 Folder —', value: '__stop__' },
					...(allItems as IDataObject[])
						.filter((item) => showFiles || !!item.folder)
						.map((item) =>
							item.folder
								? { name: `▶ ${item.name as string}`, value: `folder:${item.id as string}` }
								: { name: item.name as string, value: `file:${item.id as string}` },
						),
				];
			},

			async getWatchLevel4(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				const parentVal = this.getNodeParameter('watchFolder3', 0) as string;
				if (!parentVal || parentVal === '__stop__') return [{ name: '— Select a Folder at Level 3 First —', value: '__stop__' }];
				if (parentVal.startsWith('file:')) return [{ name: '— File Selected at Level 3 — No Subfolders —', value: '__stop__' }];
				const event = this.getNodeParameter('event', 0) as string;
				const showFiles = event === 'file.updated';
				const driveType = this.getNodeParameter('driveType') as string;
				let driveEndpoint = '/me/drive';
				if (driveType === 'user') {
					const userId = this.getNodeParameter('userId', 0) as string;
					if (userId) driveEndpoint = `/users/${userId}/drive`;
				} else if (driveType === 'site') {
					const siteId = this.getNodeParameter('siteId', 0) as string;
					driveEndpoint = `/sites/${siteId}/drive`;
				}
				const parentId = parentVal.replace('folder:', '');
				const allItems = await microsoftApiRequestAllItems.call(this, 'value', 'GET',
					`${driveEndpoint}/items/${parentId}/children?$select=id,name,folder,file`,
				);
				return [
					{ name: '— Use Level 3 Folder —', value: '__stop__' },
					...(allItems as IDataObject[])
						.filter((item) => showFiles || !!item.folder)
						.map((item) =>
							item.folder
								? { name: `▶ ${item.name as string}`, value: `folder:${item.id as string}` }
								: { name: item.name as string, value: `file:${item.id as string}` },
						),
				];
			},

			async getWatchLevel5(this: ILoadOptionsFunctions): Promise<Array<{ name: string; value: string }>> {
				const parentVal = this.getNodeParameter('watchFolder4', 0) as string;
				if (!parentVal || parentVal === '__stop__') return [{ name: '— Select a Folder at Level 4 First —', value: '__stop__' }];
				if (parentVal.startsWith('file:')) return [{ name: '— File Selected at Level 4 — No Subfolders —', value: '__stop__' }];
				const event = this.getNodeParameter('event', 0) as string;
				const showFiles = event === 'file.updated';
				const driveType = this.getNodeParameter('driveType') as string;
				let driveEndpoint = '/me/drive';
				if (driveType === 'user') {
					const userId = this.getNodeParameter('userId', 0) as string;
					if (userId) driveEndpoint = `/users/${userId}/drive`;
				} else if (driveType === 'site') {
					const siteId = this.getNodeParameter('siteId', 0) as string;
					driveEndpoint = `/sites/${siteId}/drive`;
				}
				const parentId = parentVal.replace('folder:', '');
				const allItems = await microsoftApiRequestAllItems.call(this, 'value', 'GET',
					`${driveEndpoint}/items/${parentId}/children?$select=id,name,folder,file`,
				);
				return [
					{ name: '— Use Level 4 Folder —', value: '__stop__' },
					...(allItems as IDataObject[])
						.filter((item) => showFiles || !!item.folder)
						.map((item) =>
							item.folder
								? { name: `▶ ${item.name as string}`, value: `folder:${item.id as string}` }
								: { name: item.name as string, value: `file:${item.id as string}` },
						),
				];
			},
		},
	};

	webhookMethods = {
		default: {
			async checkExists(this: IHookFunctions): Promise<boolean> {
				const webhookData = this.getWorkflowStaticData('node');
				return !!webhookData.subscriptionId;
			},
			async create(this: IHookFunctions): Promise<boolean> {
				const webhookUrl = this.getNodeWebhookUrl('default');
				const driveType = this.getNodeParameter('driveType') as string;
				const watchTarget = resolveWatchTarget(this);

				let driveEndpoint = '/me/drive';
				if (driveType === 'user') {
					const userId = this.getNodeParameter('userId') as string;
					if (userId) {
						driveEndpoint = `/users/${userId}/drive`;
					}
				} else if (driveType === 'site') {
					const siteId = this.getNodeParameter('siteId') as string;
					driveEndpoint = `/sites/${siteId}/drive`;
				}

				// Graph change notifications only support drive root — folder/file filtering is done via delta
				const resource = `${driveEndpoint}/root`.replace(/^\//, '');

				const body = {
					changeType: 'updated',
					notificationUrl: webhookUrl,
					resource,
					expirationDateTime: new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toISOString(),
					clientState: 'n8n-onedrive-business',
				};

				let responseData: IDataObject;
				try {
					responseData = await microsoftApiRequest.call(this, 'POST', '/subscriptions', body) as IDataObject;
				} catch (error) {
					const errMsg = (error as Error).message || String(error);
					throw new NodeOperationError(
						this.getNode(),
						`Graph subscription failed — resource: "${resource}", webhookUrl: "${webhookUrl}". Error: ${errMsg}`,
					);
				}

				const webhookData = this.getWorkflowStaticData('node');
				webhookData.subscriptionId = responseData.id as string;
				webhookData.subscriptionExpiration = responseData.expirationDateTime as string;

				// Initialize the delta link now so first real notification only returns NEW changes
				const deltaEndpoint =
					watchTarget.scope === 'root'
						? `${driveEndpoint}/root/delta`
						: `${driveEndpoint}/items/${watchTarget.scope}/delta`;

				try {
					const initialKnownItems: { [key: string]: { eTag: string; lastModifiedDateTime: string } } = {};
					let deltaUrl: string | undefined = deltaEndpoint;
					while (deltaUrl) {
						const deltaResp = (deltaUrl.startsWith('https://')
							? await microsoftApiRequest.call(this, 'GET', '', {}, {}, deltaUrl)
							: await microsoftApiRequest.call(this, 'GET', deltaUrl)) as IDataObject;
						// Record current state of each item so later changes are classified as 'updated', not 'created'
						if (Array.isArray(deltaResp.value)) {
							for (const item of deltaResp.value as IDataObject[]) {
								if (item.id && item.eTag && !item.deleted) {
									initialKnownItems[item.id as string] = {
										eTag: item.eTag as string,
										lastModifiedDateTime: item.lastModifiedDateTime as string,
									};
								}
							}
						}
						if (deltaResp['@odata.deltaLink']) {
							webhookData.deltaLink = deltaResp['@odata.deltaLink'] as string;
							deltaUrl = undefined;
						} else {
							deltaUrl = deltaResp['@odata.nextLink'] as string | undefined;
						}
					}
					webhookData.lastKnownItems = initialKnownItems;
				} catch {
					// Non-fatal: without delta init, first trigger may emit existing items
				}

				return true;
			},
			async delete(this: IHookFunctions): Promise<boolean> {
				const webhookData = this.getWorkflowStaticData('node');
				if (webhookData.subscriptionId) {
					try {
						await microsoftApiRequest.call(
							this,
							'DELETE',
							`/subscriptions/${webhookData.subscriptionId}`,
						);
					} catch {
						// Ignore errors on delete — subscription may have already expired
					}
					delete webhookData.subscriptionId;
					delete webhookData.subscriptionExpiration;
				}
				return true;
			},
		},
	};

	async webhook(this: IWebhookFunctions): Promise<IWebhookResponseData> {
		const query = this.getQueryData() as IDataObject;

		if (query.validationToken) {
			return {
				webhookResponse: query.validationToken as string,
			};
		}

		const event = this.getNodeParameter('event') as string;
		const driveType = this.getNodeParameter('driveType') as string;
		const watchTarget = resolveWatchTarget(this);

		let driveEndpoint = '/me/drive';
		if (driveType === 'user') {
			const userId = this.getNodeParameter('userId') as string;
			if (userId) {
				driveEndpoint = `/users/${userId}/drive`;
			}
		} else if (driveType === 'site') {
			const siteId = this.getNodeParameter('siteId') as string;
			driveEndpoint = `/sites/${siteId}/drive`;
		}

		const webhookData = this.getWorkflowStaticData('node');
		const state: IStateData = {
			deltaLink: (webhookData.deltaLink as string) || undefined,
			processedVersions: (webhookData.processedVersions as { [key: string]: boolean }) || {},
			lastKnownItems:
				(webhookData.lastKnownItems as {
					[key: string]: { eTag: string; lastModifiedDateTime: string };
				}) || {},
		};

		const deltaEndpoint =
			watchTarget.scope === 'root'
				? `${driveEndpoint}/root/delta`
				: `${driveEndpoint}/items/${watchTarget.scope}/delta`;

		let deltaUrl = state.deltaLink || deltaEndpoint;
		const changes: IDataObject[] = [];

		try {
			let hasMore = true;
			while (hasMore) {
				const response = deltaUrl.startsWith('https://')
					? await microsoftApiRequest.call(this, 'GET', '', {}, {}, deltaUrl)
					: await microsoftApiRequest.call(this, 'GET', deltaUrl);

				if (response.value && Array.isArray(response.value)) {
					for (const item of response.value as IDataObject[]) {
						if (item.deleted) {
							continue;
						}

						// If watching a specific file, skip all other items
						if (watchTarget.fileId && item.id !== watchTarget.fileId) {
							continue;
						}

						const isFile = !!item.file;

						if (!isStableItem(item)) {
							continue;
						}

						const versionKey = `${item.id as string}_${item.eTag as string}`;
						if (state.processedVersions[versionKey]) {
							continue;
						}

						let itemEvent = '';
						if (!state.lastKnownItems[item.id as string]) {
							itemEvent = isFile ? 'file.created' : 'folder.created';
						} else if (item.eTag !== state.lastKnownItems[item.id as string].eTag) {
							itemEvent = isFile ? 'file.updated' : 'folder.updated';
						}

						if (itemEvent === event) {
							changes.push(item);
							state.processedVersions[versionKey] = true;
						}

						state.lastKnownItems[item.id as string] = {
							eTag: item.eTag as string,
							lastModifiedDateTime: item.lastModifiedDateTime as string,
						};
					}
				}

				if (response['@odata.deltaLink']) {
					state.deltaLink = response['@odata.deltaLink'] as string;
					hasMore = false;
				} else if (response['@odata.nextLink']) {
					deltaUrl = response['@odata.nextLink'] as string;
				} else {
					hasMore = false;
				}
			}

			webhookData.deltaLink = state.deltaLink;
			webhookData.processedVersions = state.processedVersions;
			webhookData.lastKnownItems = state.lastKnownItems;

			if (changes.length === 0) {
				return { workflowData: [] };
			}

			return {
				workflowData: [changes.map((item) => ({ json: item }))],
			};
		} catch (error) {
			throw new NodeOperationError(this.getNode(), error as Error);
		}
	}
}

interface IWatchTarget {
	scope: string;
	fileId?: string;
}

function resolveWatchTarget(ctx: IHookFunctions | IWebhookFunctions): IWatchTarget {
	const folderSelection = ctx.getNodeParameter('folderSelection') as string;
	if (folderSelection === 'root') return { scope: 'root' };
	if (folderSelection === 'id') return { scope: ctx.getNodeParameter('watchFolderId') as string };
	// browse mode — walk levels tracking last folder scope and any final file selection
	const levels = ['watchFolder1', 'watchFolder2', 'watchFolder3', 'watchFolder4', 'watchFolder5'];
	let scopeId = 'root';
	let fileId: string | undefined;
	for (const level of levels) {
		const val = ctx.getNodeParameter(level, '') as string;
		if (!val || val === '__stop__') break;
		if (val.startsWith('folder:')) {
			scopeId = val.replace('folder:', '');
			fileId = undefined;
		} else if (val.startsWith('file:')) {
			fileId = val.replace('file:', '');
			break;
		}
	}
	return { scope: scopeId, fileId };
}

function isStableItem(item: IDataObject): boolean {
	if (!item.file) {
		return true;
	}

	// A file must have a non-zero size and hash to be considered fully uploaded
	const hasSize = item.size && (item.size as number) > 0;
	const hasHash = item.file && (item.file as IDataObject).hashes;

	if (!hasSize || !hasHash) {
		return false;
	}

	return true;
}
