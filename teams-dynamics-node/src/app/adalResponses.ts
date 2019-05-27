export interface IErrorResponse {
    error: string;
    errorDescription: string;
}
export interface IUserCodeResponse {
    userCode: string;
    deviceCode: string;
    verificationUrl: string;
    expiresIn: number;
    interval: number;
    message: string;
}
export interface IDeviceCodeTokenResponse {
    tokenType: "Bearer";
    expiresIn: number;
    expiresOn: Date;
    resource: string;
    accessToken: string;
    refreshToken: string;
    userId: string;
    isUserDisplayable: boolean;
    familyName: string;
    givenName: string;
    oid: string;
    tenantId: string;
    isMRRT: boolean;
    _clientid: string;
    _authority: string;

}
export interface IAcquireTokenResoponse {
    tokenType: "Bearer";
    expiresIn: number;
    expiresOn: Date;
    resource: string;
    accessToken: string;
    refreshToken: string;
    userId: string;
    isUserDisplayable: boolean;
    familyName: string;
    givenName: string;
    oid: string;
    tenantId: string;
}

