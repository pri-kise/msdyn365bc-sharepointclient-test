// ------------------------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
// ------------------------------------------------------------------------------------------------
namespace System.Integration.Sharepoint;

codeunit 50101 "PTE SharepointClientCred." implements "SharePoint Authorization"
{
    Access = Internal;
    InherentPermissions = X;

    var
        ClientId: Text;
        Certificate: Text;
        AadTenantId: Text;
        Scopes: List of [Text];

    procedure SetParameters(NewAadTenantId: Text; NewClientId: Text; NewCertificate: Text; NewScopes: List of [Text])
    begin
        AadTenantId := NewAadTenantId;
        ClientId := NewClientId;
        Certificate := NewCertificate;
        Scopes := NewScopes;
    end;

    procedure Authorize(var HttpRequestMessage: HttpRequestMessage);
    var
        Headers: HttpHeaders;
        BearerTxt: Label 'Bearer %1', Comment = '%1 = Token', Locked = true;
    begin
        HttpRequestMessage.GetHeaders(Headers);
        Headers.Add('Authorization', SecretStrSubstNo(BearerTxt, GetToken()));
    end;

    local procedure GetToken(): SecretText
    var
        ErrorText: Text;
        AccessToken: SecretText;
    begin
        if not AcquireToken(AccessToken, ErrorText) then
            Error(ErrorText);
        exit(AccessToken);
    end;

    local procedure AcquireToken(var AccessToken: SecretText; var ErrorText: Text): Boolean
    var
        OAuth2: Codeunit System.Security.Authentication.OAuth2;
        FailedErr: Label 'Failed to retrieve an access token.';
        //TODO: Check Authority Url
        // ClientCredentialsTokenAuthorityUrlTxt: Label 'https://login.microsoftonline.com/%1/oauth2/v2.0/token', Comment = '%1 = AAD tenant ID', Locked = true;
        ClientCredentialsTokenAuthorityUrlTxt: Label 'https://login.microsoftonline.com/%1/oauth2/v2.0/authorize', Comment = '%1 = AAD tenant ID', Locked = true;
        IsSuccess: Boolean;
        AuthorityUrl: Text;
        IdToken: Text;
    begin
        AuthorityUrl := StrSubstNo(ClientCredentialsTokenAuthorityUrlTxt, AadTenantId);
        ClearLastError();
        if (not OAuth2.AcquireTokensFromCacheWithCertificate(ClientId, Certificate, '', AuthorityUrl, Scopes, AccessToken, IdToken)) or (AccessToken.IsEmpty()) then
            OAuth2.AcquireTokensWithCertificate(ClientId, Certificate, '', AuthorityUrl, Scopes, AccessToken, IdToken);

        IsSuccess := not AccessToken.IsEmpty();

        if not IsSuccess then begin
            ErrorText := GetLastErrorText();
            if ErrorText = '' then
                ErrorText := FailedErr;
        end;

        exit(IsSuccess);
    end;
}