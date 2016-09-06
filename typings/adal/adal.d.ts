//*********************************************************
// Typescript Definitions for AdalJS v1.0.9 or above
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//*********************************************************

declare type CacheOptions = "localStorage" | "sessionStorage";

declare type Endpoints = { [resource: string]: string };

declare interface IUserInfo {
	isAuthenticated: boolean,
	userName: string,
	loginError: string,
	profile: any
}

declare interface IAdalRequestType {
	LOGIN: string;
	RENEW_TOKEN: string;
	UNKNOWN: string;
}

declare interface IAdalConstants {
	ACCESS_TOKEN: string;
	EXPIRES_IN: string;
	ID_TOKEN: string;
	ERROR_DESCRIPTION: string;
	SESSION_STATE: string;
	STORAGE: {
		TOKEN_KEYS: string;
		ACCESS_TOKEN_KEY: string;
		EXPIRATION_KEY: string;
		START_PAGE: string;
		START_PAGE_PARAMS: string;
		STATE_LOGIN: string;
		STATE_RENEW: string;
		STATE_RENEW_RESOURCE: string;
		NONCE_IDTOKEN: string;
		SESSION_STATE: string;
		USERNAME: string;
		IDTOKEN: string;
		ERROR: string;
		ERROR_DESCRIPTION: string;
		LOGIN_REQUEST: string;
		LOGIN_ERROR: string;
	},
	RESOURCE_DELIMETER: string;
	ERR_MESSAGES: {
		NO_TOKEN: string;
	},
	LOGGING_LEVEL: {
		ERROR: number;
		WARN: number;
		INFO: number;
		VERBOSE: number;
	},
	LEVEL_STRING_MAP: {
		0: string;
		1: string;
		2: string;
		3: string;
	}
}

declare interface IAuthenticationConfig {
	instance?: string;
	clientId: string;
	tenant: string;
	endpoints: Endpoints;
	redirectUri?: string;
	cacheLocation?: CacheOptions;
	postLogoutRedirectUri: string;
	loginResource?: string;

	state?: string;
	correlationId?: string;
	expireOffsetSeconds?: number;
	displayCall?: (navigateUrl: string) => any;
}

declare var Logging: {
	level: number;
	log: (string) => void;
};

declare class AuthenticationContext {
	constructor(config: IAuthenticationConfig);

	instance: string;
	config: IAuthenticationConfig;

	isCallback(windowHash: string): boolean;
	handleWindowCallback(): void;
	getLoginError(): string;
	getCachedUser(): IUserInfo;

	login(): void;
	loginInProgress(): boolean;
	logOut(): void;
	acquireToken(resource: string, callback: (error: string, token: string) => void);

	_getItem(key: string): string;
	_saveItem(key: string, obj: any): string;
	_getNavigateUrl(responseType: string, resource: string): string;
	_getHostFromUri(uri: string): string;

	getCachedToken(resource: string): string;
	getResourceForEndpoint(endpoint: string): string;
	clearCache(): void;
	clearCacheForResource(resource: string): void;
	info(message: string): void;
	verbose(message: string): void;

	log(level: number, message: string, error: Error): void;
	error(message: string, error: Error): void;
	warn(message: string): void;
	info(message: string): void;
	verbose(message: string): void;
	_libVersion(): string;

	CONSTANTS: IAdalConstants;
	REQUEST_TYPE: IAdalRequestType;
}