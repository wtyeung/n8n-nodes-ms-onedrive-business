import type { ICredentialType, INodeProperties } from 'n8n-workflow';

export class MicrosoftOneDriveBusinessOAuth2Api implements ICredentialType {
	name = 'microsoftOneDriveBusinessOAuth2Api';

	extends = ['oAuth2Api'];

	displayName = 'MS OneDrive Business OAuth2 API';

	documentationUrl = 'https://github.com/timyeung/n8n-nodes-ms-onedrive-business';

	properties: INodeProperties[] = [
		{
			displayName: 'Grant Type',
			name: 'grantType',
			type: 'hidden',
			default: 'authorizationCode',
		},
		{
			displayName: 'Authorization URL',
			name: 'authUrl',
			type: 'string',
			default: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
			description: 'Azure AD OAuth2 authorization endpoint',
		},
		{
			displayName: 'Access Token URL',
			name: 'accessTokenUrl',
			type: 'string',
			default: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
			description: 'Azure AD OAuth2 token endpoint',
		},
		{
			displayName: 'Auth URI Query Parameters',
			name: 'authQueryParameters',
			type: 'hidden',
			default: 'response_mode=query',
		},
		{
			displayName: 'Scope',
			name: 'scope',
			type: 'hidden',
			default: 'Files.ReadWrite.All Sites.ReadWrite.All offline_access',
			description: 'Required permissions for OneDrive Business operations',
		},
		{
			displayName: 'Authentication',
			name: 'authentication',
			type: 'hidden',
			default: 'body',
		},
	];
}
