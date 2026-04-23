import type { INodeProperties } from 'n8n-workflow';

export const fileOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['file'],
			},
		},
		options: [
			{
				name: 'Delete',
				value: 'delete',
				description: 'Delete a file',
				action: 'Delete a file',
			},
			{
				name: 'Download',
				value: 'download',
				description: 'Download a file',
				action: 'Download a file',
			},
			{
				name: 'Get',
				value: 'get',
				description: 'Get file metadata',
				action: 'Get a file',
			},
			{
				name: 'Rename',
				value: 'rename',
				description: 'Rename a file',
				action: 'Rename a file',
			},
			{
				name: 'Search',
				value: 'search',
				description: 'Search for files',
				action: 'Search files',
			},
			{
				name: 'Share',
				value: 'share',
				description: 'Create a sharing link for a file',
				action: 'Share a file',
			},
			{
				name: 'Upload',
				value: 'upload',
				description: 'Upload a file',
				action: 'Upload a file',
			},
		],
		default: 'upload',
	},
];

export const fileFields: INodeProperties[] = [
	{
		displayName: 'Drive Type',
		name: 'driveType',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['file'],
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
				resource: ['file'],
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
				resource: ['file'],
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
		displayName: 'File Selection',
		name: 'fileSelection',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['file'],
				operation: ['delete', 'download', 'get', 'rename', 'share'],
			},
		},
		options: [
			{
				name: 'Browse',
				value: 'browse',
				description: 'Browse folders step-by-step to find your file',
			},
			{
				name: 'By Path',
				value: 'path',
				description: 'Specify file by path (e.g., /Documents/file.pdf)',
			},
			{
				name: 'By ID',
				value: 'id',
				description: 'Specify file by its unique ID',
			},
		],
		default: 'browse',
		description: 'How to specify the file',
	},
	{
		displayName: 'Level 1 Name or ID',
		name: 'browseFolder1',
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getBrowseLevel1',
			loadOptionsDependsOn: ['driveType', 'userId', 'siteId'],
		},
		displayOptions: {
			show: {
				resource: ['file'],
				operation: ['delete', 'download', 'get', 'rename', 'share'],
				fileSelection: ['browse'],
			},
		},
		default: '',
		required: true,
		description: 'Select a 📁 folder to go deeper, or a 📄 file to finish. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
	},
	{
		displayName: 'Level 2 Name or ID',
		name: 'browseFolder2',
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getBrowseLevel2',
			loadOptionsDependsOn: ['driveType', 'userId', 'siteId', 'browseFolder1'],
		},
		displayOptions: {
			show: {
				resource: ['file'],
				operation: ['delete', 'download', 'get', 'rename', 'share'],
				fileSelection: ['browse'],
			},
		},
		default: '',
		description: 'Select a 📁 folder to go deeper, or a 📄 file to finish. Only shown if Level 1 is a folder. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
	},
	{
		displayName: 'Level 3 Name or ID',
		name: 'browseFolder3',
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getBrowseLevel3',
			loadOptionsDependsOn: ['driveType', 'userId', 'siteId', 'browseFolder1', 'browseFolder2'],
		},
		displayOptions: {
			show: {
				resource: ['file'],
				operation: ['delete', 'download', 'get', 'rename', 'share'],
				fileSelection: ['browse'],
			},
		},
		default: '',
		description: 'Select a 📁 folder to go deeper, or a 📄 file to finish. Only shown if Level 2 is a folder. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
	},
	{
		displayName: 'Level 4 Name or ID',
		name: 'browseFolder4',
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getBrowseLevel4',
			loadOptionsDependsOn: ['driveType', 'userId', 'siteId', 'browseFolder1', 'browseFolder2', 'browseFolder3'],
		},
		displayOptions: {
			show: {
				resource: ['file'],
				operation: ['delete', 'download', 'get', 'rename', 'share'],
				fileSelection: ['browse'],
			},
		},
		default: '',
		description: 'Select a 📁 folder to go deeper, or a 📄 file to finish. Only shown if Level 3 is a folder. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
	},
	{
		displayName: 'Level 5 Name or ID',
		name: 'browseFolder5',
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getBrowseLevel5',
			loadOptionsDependsOn: ['driveType', 'userId', 'siteId', 'browseFolder1', 'browseFolder2', 'browseFolder3', 'browseFolder4'],
		},
		displayOptions: {
			show: {
				resource: ['file'],
				operation: ['delete', 'download', 'get', 'rename', 'share'],
				fileSelection: ['browse'],
			},
		},
		default: '',
		description: 'Select a 📄 file to finish. Only shown if Level 4 is a folder. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
	},
	{
		displayName: 'File Path',
		name: 'filePath',
		type: 'resourceLocator',
		displayOptions: {
			show: {
				resource: ['file'],
				operation: ['delete', 'download', 'get', 'rename', 'share'],
				fileSelection: ['path'],
			},
		},
		default: { mode: 'list', value: '' },
		required: true,
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				placeholder: 'Select a file...',
				typeOptions: {
					searchListMethod: 'searchFiles',
					searchable: true,
					searchFilterRequired: false,
				},
			},
			{
				displayName: 'By Path',
				name: 'path',
				type: 'string',
				placeholder: '/Documents/MyFile.pdf',
				hint: 'Enter the full path from root',
			},
			{
				displayName: 'By ID',
				name: 'id',
				type: 'string',
				placeholder: 'File ID',
				validation: [
					{
						type: 'regex',
						properties: {
							regex: '^[a-zA-Z0-9_-]+$',
							errorMessage: 'Not a valid file ID',
						},
					},
				],
			},
		],
		description: 'The file to operate on',
	},
	{
		displayName: 'File ID',
		name: 'fileId',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['file'],
				operation: ['delete', 'download', 'get', 'rename', 'share'],
				fileSelection: ['id'],
			},
		},
		default: '',
		required: true,
		description: 'The unique ID of the file',
	},
	{
		displayName: 'Binary Property',
		name: 'binaryPropertyName',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['file'],
				operation: ['download'],
			},
		},
		default: 'data',
		required: true,
		description: 'Name of the binary property to store the downloaded file',
	},
	{
		displayName: 'New Name',
		name: 'newName',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['file'],
				operation: ['rename'],
			},
		},
		default: '',
		required: true,
		description: 'The new name for the file',
	},
	{
		displayName: 'Search Query',
		name: 'query',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['file'],
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
				resource: ['file'],
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
				resource: ['file'],
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
	{
		displayName: 'Upload To Folder \u2014 Selection',
		name: 'folderSelection',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['file'],
				operation: ['upload'],
			},
		},
		options: [
			{
				name: 'Browse',
				value: 'browse',
				description: 'Navigate folder by folder to select the destination',
			},
			{
				name: 'By Path',
				value: 'path',
				description: 'Specify folder by path',
			},
			{
				name: 'By ID',
				value: 'id',
				description: 'Specify folder by its unique ID',
			},
		],
		default: 'browse',
		description: 'How to specify the upload destination folder',
	},
	{
		displayName: 'Upload To Folder',
		name: 'parentId',
		type: 'resourceLocator',
		displayOptions: {
			show: {
				resource: ['file'],
				operation: ['upload'],
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
		description: 'The folder where the file will be uploaded',
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
				resource: ['file'],
				operation: ['upload'],
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
				resource: ['file'],
				operation: ['upload'],
				folderSelection: ['browse'],
			},
		},
		default: '',
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
				resource: ['file'],
				operation: ['upload'],
				folderSelection: ['browse'],
			},
		},
		default: '',
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
				resource: ['file'],
				operation: ['upload'],
				folderSelection: ['browse'],
			},
		},
		default: '',
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
				resource: ['file'],
				operation: ['upload'],
				folderSelection: ['browse'],
			},
		},
		default: '',
		description: 'Select the deepest target folder here, or leave as is to use Level 4 as the target. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
	},
	{
		displayName: 'File Name',
		name: 'fileName',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['file'],
				operation: ['upload'],
			},
		},
		default: '',
		placeholder: 'Leave empty to use binary data filename',
		description: 'Name for the uploaded file. Leave empty to use the filename from the binary data.',
	},
	{
		displayName: 'Binary Data',
		name: 'binaryData',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['file'],
				operation: ['upload'],
			},
		},
		default: true,
		description: 'Whether the file content is in binary format',
	},
	{
		displayName: 'Binary Property',
		name: 'binaryPropertyName',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['file'],
				operation: ['upload'],
				binaryData: [true],
			},
		},
		default: 'data',
		required: true,
		description: 'Name of the binary property containing the file data',
	},
	{
		displayName: 'File Content',
		name: 'fileContent',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['file'],
				operation: ['upload'],
				binaryData: [false],
			},
		},
		default: '',
		description: 'Text content of the file',
	},
];
