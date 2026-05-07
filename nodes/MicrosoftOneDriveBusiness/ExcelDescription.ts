import type { INodeProperties } from 'n8n-workflow';

export const excelOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['excel'],
			},
		},
		options: [
			{
				name: 'Read Rows',
				value: 'readRows',
				description: 'Read rows from Excel worksheet',
				action: 'Read rows from excel worksheet',
			},
			{
				name: 'Append or Update Row',
				value: 'appendOrUpdate',
				description: 'Append a new row or update an existing one',
				action: 'Append or update row in excel worksheet',
			},
			{
				name: 'Delete Rows',
				value: 'deleteRows',
				description: 'Delete rows from Excel worksheet',
				action: 'Delete rows from excel worksheet',
			},
		],
		default: 'readRows',
	},
];

export const excelFields: INodeProperties[] = [
	// Drive Type
	{
		displayName: 'Drive Type',
		name: 'driveType',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['excel'],
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
				resource: ['excel'],
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
				resource: ['excel'],
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
				resource: ['excel'],
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
	// File Selection
	{
		displayName: 'File Selection',
		name: 'fileSelection',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['excel'],
			},
		},
		options: [
			{
				name: 'Browse',
				value: 'browse',
				description: 'Browse folders step-by-step to find your Excel file',
			},
			{
				name: 'By Path',
				value: 'path',
				description: 'Specify file by path (e.g., /Documents/file.xlsx)',
			},
			{
				name: 'By ID',
				value: 'id',
				description: 'Specify file by its unique ID',
			},
		],
		default: 'browse',
		description: 'How to specify the Excel file',
	},
	// Browse Levels
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
				resource: ['excel'],
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
			loadOptionsDependsOn: ['browseFolder1'],
		},
		displayOptions: {
			show: {
				resource: ['excel'],
				fileSelection: ['browse'],
			},
		},
		default: '',
		description: 'Select a 📁 folder to go deeper, or a 📄 file to finish. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
	},
	{
		displayName: 'Level 3 Name or ID',
		name: 'browseFolder3',
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getBrowseLevel3',
			loadOptionsDependsOn: ['browseFolder2'],
		},
		displayOptions: {
			show: {
				resource: ['excel'],
				fileSelection: ['browse'],
			},
		},
		default: '',
		description: 'Select a 📁 folder to go deeper, or a 📄 file to finish. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
	},
	{
		displayName: 'Level 4 Name or ID',
		name: 'browseFolder4',
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getBrowseLevel4',
			loadOptionsDependsOn: ['browseFolder3'],
		},
		displayOptions: {
			show: {
				resource: ['excel'],
				fileSelection: ['browse'],
			},
		},
		default: '',
		description: 'Select a 📁 folder to go deeper, or a 📄 file to finish. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
	},
	{
		displayName: 'Level 5 Name or ID',
		name: 'browseFolder5',
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getBrowseLevel5',
			loadOptionsDependsOn: ['browseFolder4'],
		},
		displayOptions: {
			show: {
				resource: ['excel'],
				fileSelection: ['browse'],
			},
		},
		default: '',
		description: 'Select a 📁 folder to go deeper, or a 📄 file to finish. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
	},
	// File Path (for path mode)
	{
		displayName: 'File Path',
		name: 'filePath',
		type: 'resourceLocator',
		displayOptions: {
			show: {
				resource: ['excel'],
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
					searchFilterRequired: true,
				},
			},
			{
				displayName: 'By Path',
				name: 'path',
				type: 'string',
				placeholder: '/Documents/report.xlsx',
				hint: 'Path relative to drive root',
			},
			{
				displayName: 'By ID',
				name: 'id',
				type: 'string',
				placeholder: '01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K',
				hint: 'File unique ID',
			},
		],
		description: 'The Excel file to access',
	},
	// File ID (for id mode)
	{
		displayName: 'File ID',
		name: 'fileId',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['excel'],
				fileSelection: ['id'],
			},
		},
		default: '',
		required: true,
		placeholder: '01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K',
		description: 'The unique ID of the Excel file',
	},
	// Worksheet Name
	{
		displayName: 'Worksheet Name or ID',
		name: 'worksheet',
		type: 'options',
		typeOptions: {
			loadOptionsMethod: 'getWorksheets',
			loadOptionsDependsOn: ['driveType', 'userId', 'siteId', 'fileSelection', 'filePath', 'fileId', 'browseFolder1', 'browseFolder2', 'browseFolder3', 'browseFolder4', 'browseFolder5'],
		},
		displayOptions: {
			show: {
				resource: ['excel'],
			},
		},
		default: '',
		required: true,
		description: 'Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>',
	},

	// Read Rows - Range Selection
	{
		displayName: 'Select a Range',
		name: 'useRange',
		type: 'boolean',
		default: false,
		displayOptions: {
			show: {
				resource: ['excel'],
				operation: ['readRows'],
			},
		},
	},
	{
		displayName: 'Range',
		name: 'range',
		type: 'string',
		placeholder: 'e.g. A1:B2',
		default: '',
		description: 'The sheet range to read the data from specified using a A1-style notation. Leave blank to return entire sheet.',
		hint: 'Leave blank to return entire sheet',
		displayOptions: {
			show: {
				resource: ['excel'],
				operation: ['readRows'],
				useRange: [true],
			},
		},
	},
	{
		displayName: 'Header Row',
		name: 'keyRow',
		type: 'number',
		typeOptions: {
			minValue: 0,
		},
		default: 0,
		hint: 'Index of the row which contains the column names',
		description: 'Relative to selected \'Range\', first row index is 0',
		displayOptions: {
			show: {
				resource: ['excel'],
				operation: ['readRows'],
				useRange: [true],
			},
		},
	},
	{
		displayName: 'First Data Row',
		name: 'dataStartRow',
		type: 'number',
		typeOptions: {
			minValue: 0,
		},
		default: 1,
		hint: 'Index of first row which contains the actual data',
		description: 'Relative to selected \'Range\', first row index is 0',
		displayOptions: {
			show: {
				resource: ['excel'],
				operation: ['readRows'],
				useRange: [true],
			},
		},
	},

	// Append or Update - Column to Match On
	{
		displayName: 'Match Record by Column Name or ID',
		name: 'columnToMatchOn',
		type: 'options',
		default: '',
		hint: 'Leave empty to always append a new row',
		description: 'The column used to find an existing row to update. If no match is found, a new row will be appended. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>.',
		typeOptions: {
			loadOptionsMethod: 'getColumns',
			loadOptionsDependsOn: ['driveType', 'userId', 'siteId', 'fileSelection', 'filePath', 'fileId', 'browseFolder1', 'browseFolder2', 'browseFolder3', 'browseFolder4', 'browseFolder5', 'worksheet'],
		},
		displayOptions: {
			show: {
				resource: ['excel'],
				operation: ['appendOrUpdate'],
			},
		},
	},

	// Append or Update - Value of Column to Match On
	{
		displayName: 'Matching Value',
		name: 'valueToMatchOn',
		type: 'string',
		default: '',
		description: 'The value to look for in the match column to identify the row to update',
		displayOptions: {
			show: {
				resource: ['excel'],
				operation: ['appendOrUpdate'],
			},
		},
	},

	// Append or Update - Data Mode
	{
		displayName: 'Data Mode',
		name: 'dataMode',
		type: 'options',
		options: [
			{
				name: 'Auto-Map Input Data to Columns',
				value: 'autoMap',
				description: 'Use when node input properties match destination column names',
			},
			{
				name: 'Map Each Column Manually',
				value: 'manual',
				description: 'Set the value for each destination column',
			},
		],
		default: 'autoMap',
		description: 'How data should be mapped to columns',
		displayOptions: {
			show: {
				resource: ['excel'],
				operation: ['appendOrUpdate'],
			},
		},
	},

	// Append or Update - Manual Mapping
	{
		displayName: 'Values to Send',
		name: 'fieldsUi',
		placeholder: 'Add Field',
		type: 'fixedCollection',
		typeOptions: {
			multipleValues: true,
		},
		displayOptions: {
			show: {
				resource: ['excel'],
				operation: ['appendOrUpdate'],
				dataMode: ['manual'],
			},
		},
		default: {},
		options: [
			{
				displayName: 'Field',
				name: 'fieldValues',
				values: [
					{
						displayName: 'Column Name or ID',
						name: 'column',
						type: 'options',
						typeOptions: {
							loadOptionsMethod: 'getColumns',
							loadOptionsDependsOn: ['driveType', 'userId', 'siteId', 'fileSelection', 'filePath', 'fileId', 'browseFolder1', 'browseFolder2', 'browseFolder3', 'browseFolder4', 'browseFolder5', 'worksheet'],
						},
						default: '',
						description: 'Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code/expressions/">expression</a>',
					},
					{
						displayName: 'Value',
						name: 'fieldValue',
						type: 'string',
						default: '',
					},
				],
			},
		],
	},

	// Delete Rows - Mode
	{
		displayName: 'Delete By',
		name: 'deleteMode',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['excel'],
				operation: ['deleteRows'],
			},
		},
		options: [
			{
				name: 'Row Number',
				value: 'rowNumber',
				description: 'Delete a single row using the _row_number field from Read Rows output',
			},
			{
				name: 'Row Range',
				value: 'rowRange',
				description: 'Delete a range of rows using a range address (e.g. 2:5)',
			},
		],
		default: 'rowNumber',
	},

	// Delete Rows - Row Number
	{
		displayName: 'Row Number',
		name: 'deleteRowNumber',
		type: 'number',
		default: 0,
		required: true,
		description: 'The row number to delete. Use the _row_number field from Read Rows output via an expression.',
		displayOptions: {
			show: {
				resource: ['excel'],
				operation: ['deleteRows'],
				deleteMode: ['rowNumber'],
			},
		},
	},

	// Delete Rows - Range
	{
		displayName: 'Range',
		name: 'deleteRange',
		type: 'string',
		placeholder: 'e.g. 2:5',
		default: '',
		required: true,
		description: 'The range of rows to delete (e.g., "2:5" to delete rows 2 through 5)',
		displayOptions: {
			show: {
				resource: ['excel'],
				operation: ['deleteRows'],
				deleteMode: ['rowRange'],
			},
		},
	},

	// Options
	{
		displayName: 'Options',
		name: 'options',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: {
			show: {
				resource: ['excel'],
			},
		},
		options: [
			{
				displayName: 'RAW Data',
				name: 'rawData',
				type: 'boolean',
				default: false,
				description: 'Whether to return the data RAW instead of parsed into keys according to their header',
				displayOptions: {
					show: {
						'/operation': ['readRows'],
					},
				},
			},
			{
				displayName: 'Data Property',
				name: 'dataProperty',
				type: 'string',
				default: 'data',
				required: true,
				displayOptions: {
					show: {
						'/operation': ['readRows'],
						rawData: [true],
					},
				},
				description: 'The name of the property into which to write the RAW data',
			},
		],
	},
];
