import type {
	IDataObject,
	IHookFunctions,
	INodeType,
	INodeTypeDescription,
	IWebhookFunctions,
	IWebhookResponseData,
} from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';

import { microsoftApiRequest } from '../MicrosoftOneDriveBusiness/GenericFunctions';

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
						description: 'Trigger when a new file is created',
					},
					{
						name: 'File Updated',
						value: 'file.updated',
						description: 'Trigger when a file is updated',
					},
					{
						name: 'Folder Created',
						value: 'folder.created',
						description: 'Trigger when a new folder is created',
					},
					{
						name: 'Folder Updated',
						value: 'folder.updated',
						description: 'Trigger when a folder is updated',
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
				name: 'folderToWatch',
				type: 'string',
				default: 'root',
				description: 'Folder ID to monitor. Use "root" to monitor the entire drive.',
			},
		],
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
				const folderToWatch = this.getNodeParameter('folderToWatch') as string;

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

				const resource =
					folderToWatch === 'root'
						? `${driveEndpoint}/root`
						: `${driveEndpoint}/items/${folderToWatch}`;

				const body = {
					changeType: 'updated',
					notificationUrl: webhookUrl,
					resource,
					expirationDateTime: new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toISOString(),
					clientState: 'n8n-onedrive-business',
				};

				try {
					const responseData = await microsoftApiRequest.call(this, 'POST', '/subscriptions', body);
					const webhookData = this.getWorkflowStaticData('node');
					webhookData.subscriptionId = responseData.id as string;
					webhookData.subscriptionExpiration = responseData.expirationDateTime as string;
					return true;
				} catch {
					return false;
				}
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
		const folderToWatch = this.getNodeParameter('folderToWatch') as string;

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
			folderToWatch === 'root'
				? `${driveEndpoint}/root/delta`
				: `${driveEndpoint}/items/${folderToWatch}/delta`;

		let deltaUrl = state.deltaLink || deltaEndpoint;
		const changes: IDataObject[] = [];

		try {
			let hasMore = true;
			while (hasMore) {
				const response = await microsoftApiRequest.call(this, 'GET', '', {}, {}, deltaUrl);

				if (response.value && Array.isArray(response.value)) {
					for (const item of response.value as IDataObject[]) {
						if (item.deleted) {
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

function isStableItem(item: IDataObject): boolean {
	if (!item.file) {
		return true;
	}

	const hasSize = item.size && (item.size as number) > 0;
	const hasHash = item.file && (item.file as IDataObject).hashes;
	const timestampsMatch =
		item.lastModifiedDateTime ===
		(item.fileSystemInfo as IDataObject)?.lastModifiedDateTime;

	if (!hasSize || !hasHash || !timestampsMatch) {
		return false;
	}

	const lastModified = new Date(item.lastModifiedDateTime as string).getTime();
	const stabilityWindow = 15000;
	if (Date.now() - lastModified < stabilityWindow) {
		return false;
	}

	return true;
}
