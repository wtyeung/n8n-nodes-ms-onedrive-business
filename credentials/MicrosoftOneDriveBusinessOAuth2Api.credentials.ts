import type { ICredentialType, INodeProperties } from 'n8n-workflow';

export class MicrosoftOneDriveBusinessOAuth2Api implements ICredentialType {
	name = 'microsoftOneDriveBusinessOAuth2Api';

	extends = ['microsoftOAuth2Api'];

	displayName = 'MS OneDrive Business OAuth2 API';

	icon = 'file:onedrive.svg' as const;

	documentationUrl = 'https://github.com/timyeung/n8n-nodes-ms-onedrive-business';

	test = {
		request: {
			baseURL: 'https://graph.microsoft.com/v1.0',
			url: '/me/drive',
		},
	};

	properties: INodeProperties[] = [
		{
			displayName: 'Scope',
			name: 'scope',
			type: 'hidden',
			default: 'Files.ReadWrite.All Sites.ReadWrite.All offline_access',
			description: 'Required permissions for OneDrive Business operations',
		},
	];
}
