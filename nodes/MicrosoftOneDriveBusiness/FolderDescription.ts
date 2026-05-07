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
			{
				name: 'Shared Folder (Link)',
				value: 'sharedLink',
				description: 'Browse a folder shared via OneDrive/SharePoint sharing link',
			},
		],
		default: 'user',
		description: 'Type of drive to access',
	},
	{
		displayName: 'Shared Folder URL',
		name: 'sharedLinkUrl',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['folder'],
				driveType: ['sharedLink'],
			},
		},
		default: '',
		placeholder: 'https://1drv.ms/... or https://contoso.sharepoint.com/:f:/...',
		description: 'Paste the sharing link of the shared folder (right-click → Share → Copy link in OneDrive/SharePoint)',
		required: true,
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
		displayName: 'Create In Folder — Selection',
		name: 'folderSelection',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['create'],
			},
		},
		options: [
			{
				name: 'Browse',
				value: 'browse',
				description: 'Navigate folder by folder to select the parent',
			},
			{
				name: 'By Path',
				value: 'path',
				description: 'Specify parent folder by path',
			},
			{
				name: 'By ID',
				value: 'id',
				description: 'Specify parent folder by its unique ID',
			},
			{
				name: 'By Sharing Link',
				value: 'link',
				description: 'Paste an OneDrive/SharePoint sharing link ("Copy link") to access any shared folder',
			},
		],
		default: 'browse',
		description: 'How to specify the parent folder',
	},
	{
		displayName: 'Sharing Link',
		name: 'folderSharingUrl',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['folder'],
				folderSelection: ['link'],
			},
		},
		default: '',
		placeholder: 'https://1drv.ms/... or https://contoso.sharepoint.com/:f:/...',
		description: 'Paste the sharing link copied from OneDrive or SharePoint ("Share → Copy link")',
		required: true,
	},
	{
		displayName: 'Create In Folder',
		name: 'parentId',
		type: 'resourceLocator',
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['create'],
				folderSelection: ['path', 'id'],
			},
		},
		default: { mode: 'path', value: '/' },
		required: true,
		modes: [
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
				name: 'Browse',
				value: 'browse',
				description: 'Navigate folder by folder to select the target',
			},
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
		default: 'browse',
		description: 'How to specify the folder',
	},
	{
		displayName: 'Level 1 Name or ID',
		name: 'browseFolderF1',
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getBrowseFolderLevel1',
			loadOptionsDependsOn: ['driveType', 'userId', 'siteId'],
		},
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['create', 'delete', 'getItems', 'rename', 'share'],
				folderSelection: ['browse'],
			},
		},
		default: '',
		required: true,
		description: 'Select a ▶ folder to go deeper, or select the target folder directly. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
	},
	{
		displayName: 'Level 2 Name or ID',
		name: 'browseFolderF2',
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getBrowseFolderLevel2',
			loadOptionsDependsOn: ['driveType', 'userId', 'siteId', 'browseFolderF1'],
		},
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['create', 'delete', 'getItems', 'rename', 'share'],
				folderSelection: ['browse'],
			},
		},
		default: '__stop__',
		description: 'Select a subfolder to go deeper, or leave as is to use Level 1 as the target. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
	},
	{
		displayName: 'Level 3 Name or ID',
		name: 'browseFolderF3',
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getBrowseFolderLevel3',
			loadOptionsDependsOn: ['driveType', 'userId', 'siteId', 'browseFolderF1', 'browseFolderF2'],
		},
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['create', 'delete', 'getItems', 'rename', 'share'],
				folderSelection: ['browse'],
			},
		},
		default: '__stop__',
		description: 'Select a subfolder to go deeper, or leave as is to use Level 2 as the target. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
	},
	{
		displayName: 'Level 4 Name or ID',
		name: 'browseFolderF4',
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getBrowseFolderLevel4',
			loadOptionsDependsOn: ['driveType', 'userId', 'siteId', 'browseFolderF1', 'browseFolderF2', 'browseFolderF3'],
		},
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['create', 'delete', 'getItems', 'rename', 'share'],
				folderSelection: ['browse'],
			},
		},
		default: '__stop__',
		description: 'Select a subfolder to go deeper, or leave as is to use Level 3 as the target. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
	},
	{
		displayName: 'Level 5 Name or ID',
		name: 'browseFolderF5',
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getBrowseFolderLevel5',
			loadOptionsDependsOn: ['driveType', 'userId', 'siteId', 'browseFolderF1', 'browseFolderF2', 'browseFolderF3', 'browseFolderF4'],
		},
		displayOptions: {
			show: {
				resource: ['folder'],
				operation: ['create', 'delete', 'getItems', 'rename', 'share'],
				folderSelection: ['browse'],
			},
		},
		default: '__stop__',
		description: 'Select the deepest target folder here, or leave as is to use Level 4 as the target. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
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
