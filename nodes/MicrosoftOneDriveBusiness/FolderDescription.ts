import type { INodeProperties } from 'n8n-workflow';

export const folderOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['folder'],
			},
		},
		options: [
			{
				name: 'Create',
				value: 'create',
				description: 'Create a folder',
				action: 'Create a folder',
			},
			{
				name: 'Delete',
				value: 'delete',
				description: 'Delete a folder',
				action: 'Delete a folder',
			},
			{
				name: 'Get Items',
				value: 'getItems',
				description: 'Get items in a folder',
				action: 'Get folder items',
			},
			{
				name: 'Rename',
				value: 'rename',
				description: 'Rename a folder',
				action: 'Rename a folder',
			},
			{
				name: 'Search',
				value: 'search',
				description: 'Search for folders',
				action: 'Search folders',
			},
			{
				name: 'Share',
				value: 'share',
				description: 'Create a sharing link for a folder',
				action: 'Share a folder',
			},
		],
		default: 'create',
	},
];

export const folderFields: INodeProperties[] = [
	{
		displayName: 'Drive Type',
		name: 'driveType',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['folder'],
			},
		},
		options: [
			{
				name: 'User Drive',
				value: 'user',
				description: 'Access user OneDrive',
			},
			{
				name: 'SharePoint Site Drive',
				value: 'site',
				description: 'Access SharePoint site drive',
			},
		],
		default: 'user',
		description: 'Type of drive to access',
	},
	{
		displayName: 'User ID',
		name: 'userId',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['folder'],
				driveType: ['user'],
			},
		},
		default: '',
		placeholder: 'user@domain.com or leave empty',
		hint: 'Leave empty to use the authenticated user from your credential',
		description: 'User email or ID. If empty, uses the authenticated user.',
	},
	{
		displayName: 'SharePoint Site',
		name: 'siteId',
		type: 'resourceLocator',
		displayOptions: {
			show: {
				resource: ['folder'],
				driveType: ['site'],
			},
		},
		default: { mode: 'list', value: '' },
		required: true,
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				placeholder: 'Select a site...',
				typeOptions: {
					searchListMethod: 'searchSites',
					searchable: true,
					searchFilterRequired: false,
				},
			},
			{
				displayName: 'By ID',
				name: 'id',
				type: 'string',
				placeholder: 'contoso.sharepoint.com,da60e844-ba1d-49bc-b4d4-d5e36bae9019,712a596e-90a1-49e3-9b48-bfa80bee8740',
				hint: 'Full SharePoint site ID',
			},
			{
				displayName: 'By URL',
				name: 'url',
				type: 'string',
				placeholder: 'https://contoso.sharepoint.com/sites/TeamSite',
				hint: 'SharePoint site URL',
			},
		],
		description: 'The SharePoint site to access',
	},
	{
		displayName: 'Create In Folder',
		name: 'parentId',
		type: 'resourceLocator',
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['create'],
			},
		},
		default: { mode: 'list', value: 'root' },
		required: true,
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				placeholder: 'Select a folder...',
				typeOptions: {
					searchListMethod: 'searchFolders',
					searchable: true,
					searchFilterRequired: false,
				},
			},
			{
				displayName: 'By Path',
				name: 'path',
				type: 'string',
				placeholder: '/Documents/MyFolder',
				hint: 'Enter folder path (use / for root)',
			},
			{
				displayName: 'By ID',
				name: 'id',
				type: 'string',
				placeholder: 'Folder ID or "root"',
			},
		],
		description: 'The parent folder where the new folder will be created',
	},
	{
		displayName: 'Folder Name',
		name: 'folderName',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['create'],
			},
		},
		default: '',
		required: true,
		description: 'Name of the folder to create',
	},
	{
		displayName: 'Folder Selection',
		name: 'folderSelection',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['delete', 'getItems', 'rename', 'share'],
			},
		},
		options: [
			{
				name: 'By Path',
				value: 'path',
				description: 'Specify folder by path (e.g., /Documents/MyFolder)',
			},
			{
				name: 'By ID',
				value: 'id',
				description: 'Specify folder by its unique ID',
			},
		],
		default: 'path',
		description: 'How to specify the folder',
	},
	{
		displayName: 'Folder Path',
		name: 'folderPath',
		type: 'resourceLocator',
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['delete', 'getItems', 'rename', 'share'],
				folderSelection: ['path'],
			},
		},
		default: { mode: 'list', value: '' },
		required: true,
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				placeholder: 'Select a folder...',
				typeOptions: {
					searchListMethod: 'searchFolders',
					searchable: true,
					searchFilterRequired: false,
				},
			},
			{
				displayName: 'By Path',
				name: 'path',
				type: 'string',
				placeholder: '/Documents/MyFolder',
				hint: 'Enter the full path from root (use / for root)',
			},
			{
				displayName: 'By ID',
				name: 'id',
				type: 'string',
				placeholder: 'Folder ID or "root"',
			},
		],
		description: 'The folder to operate on',
	},
	{
		displayName: 'Folder ID',
		name: 'folderId',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['delete', 'getItems', 'rename', 'share'],
				folderSelection: ['id'],
			},
		},
		default: '',
		required: true,
		description: 'The unique ID of the folder. Use "root" for the root folder.',
	},
	{
		displayName: 'New Name',
		name: 'newName',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['rename'],
			},
		},
		default: '',
		required: true,
		description: 'The new name for the folder',
	},
	{
		displayName: 'Search Query',
		name: 'query',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['search'],
			},
		},
		default: '',
		required: true,
		description: 'Search query string',
	},
	{
		displayName: 'Link Type',
		name: 'linkType',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['share'],
			},
		},
		options: [
			{
				name: 'View',
				value: 'view',
				description: 'Read-only link',
			},
			{
				name: 'Edit',
				value: 'edit',
				description: 'Read-write link',
			},
		],
		default: 'view',
		description: 'Type of sharing link to create',
	},
	{
		displayName: 'Link Scope',
		name: 'linkScope',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['share'],
			},
		},
		options: [
			{
				name: 'Anonymous',
				value: 'anonymous',
				description: 'Anyone with the link can access',
			},
			{
				name: 'Organization',
				value: 'organization',
				description: 'Only people in your organization can access',
			},
		],
		default: 'organization',
		description: 'Scope of the sharing link',
	},
];
