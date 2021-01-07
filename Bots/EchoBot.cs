// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
//
// Generated with Bot Builder V4 SDK Template for Visual Studio EchoBot v4.6.2

using System.Collections.Generic;
using System.Threading;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema.Teams;
using System;
using System.IdentityModel.Tokens.Jwt;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.IdentityModel.JsonWebTokens;
using Microsoft.IdentityModel.Tokens;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using AdaptiveCards;
using System.IO;
//using System.Security.Cryptography.HashAlgorithm;

namespace HelloWorldBot.Bots
{
    public class EchoBot : TeamsActivityHandler
    {
        private const string microsoftTenantID = "f8cdef31-a31e-4b4a-93e4-5f571e91255a";
        private const string tokenRequestUrl = "https://login.microsoftonline.com/f8cdef31-a31e-4b4a-93e4-5f571e91255a/oauth2/v2.0/token";
        private const string clientId = "e5e15768-1702-474d-ba7b-904c7cad2bcf";
        private const string clientAssertion = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer";
        private const string clientCredentials = "client_credentials";

        protected override async Task<MessagingExtensionResponse> OnTeamsAppBasedLinkQueryAsync(ITurnContext<IInvokeActivity> turnContext, AppBasedLinkQuery query, CancellationToken cancellationToken)
        {
            if(turnContext != null && turnContext.Activity != null)
            {
                JObject valueObject = JObject.FromObject(turnContext.Activity.Value);
                
                if (valueObject["authentication"] != null)
                {
                    JObject authObj = JObject.FromObject(valueObject["authentication"]);
                    
                    string accessToken = GetPostTransformedPFTToken((authObj["token"]).ToString());
                    string actorToken = await getActorToken(query.url);

                    //string spMetadata = await getSharePointMetadata(accessToken, actorToken, "");

                   var previewCard = new HeroCard
                    {
                        Title = "Test Title",
                        Subtitle = "Test Subtitle",
                        Text = "Sample text",
                    };

                    try
                    {
                        var client = new HttpClient();
                        var cardTemplate = File.ReadAllText(@".\CardTemplate\adaptiveCardSample.json");
                        var parsedResult = AdaptiveCard.FromJson(cardTemplate);

                        MessagingExtensionAttachment attachment = new MessagingExtensionAttachment
                        {
                            ContentType = AdaptiveCard.ContentType,
                            Content = parsedResult.Card,
                            Preview = previewCard.ToAttachment()
                        };

                        var result = new MessagingExtensionResult("list", "result", new[] { attachment });
                        return new MessagingExtensionResponse(result);
                    }
                    catch (AdaptiveSerializationException e)
                    {
                        throw e;
                    }

                }
                else
                {
                    // Request SSO token if not present in turnContext
                    return new MessagingExtensionResponse
                    {
                        ComposeExtension = new MessagingExtensionResult
                        {
                            Type = "silentAuth"
                        }
                    };
                }
            }
            return null;
        }

        /*
         * Modufy the header in the PFT token so that it can be accepted by SharePoint service
         * Documentation: https://aadwiki.windows-int.net/index.php?title=Server-to-server_authentication
         */
        private string GetPostTransformedPFTToken(String preTransformedPFTToken)
        {
            JwtSecurityToken jwtSecurityToken = new JwtSecurityToken(preTransformedPFTToken);
            JObject headerObject = JObject.Parse(Base64UrlEncoder.Decode(jwtSecurityToken.RawHeader));

            string nonceIn = headerObject.GetValue("nonce", StringComparison.OrdinalIgnoreCase).ToString();
            string hashName = headerObject.GetValue("alg", StringComparison.OrdinalIgnoreCase).ToString();

            HashAlgorithm hashAlgorithm = EchoBot.GetHashFunction(hashName);
            string nonceOut = Base64UrlEncoder.Encode(hashAlgorithm.ComputeHash(Encoding.UTF8.GetBytes(nonceIn)));
            headerObject["nonce"] = nonceOut;
            string newHeaderString = Base64UrlEncoder.Encode(headerObject.ToString(Formatting.None));
            string postTransformedToken = $"{newHeaderString}.{jwtSecurityToken.RawPayload}.{jwtSecurityToken.RawSignature}";
            return postTransformedToken;
        }

        private static HashAlgorithm GetHashFunction(string hashAlgorithmName)
        {
            switch (hashAlgorithmName)
            {
                case "RS256":
                    return SHA256.Create();
                default:
                    string errorMessage = $"Algorithm [{hashAlgorithmName}] not supported for PFT at this time";
                    throw new NotSupportedException(errorMessage);
            }
        }

        private async Task<string> getActorToken(string spUrl)
        {
            string scope = (new Uri(spUrl)).Host + "/.default";
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri(tokenRequestUrl);

            string clientAssertion = GetSignedClientAssertion();

            var values = new Dictionary<string, string>
            {
                { "grant_type", clientCredentials },
                { "client_id", clientId },
                {"client_assertion", clientAssertion },
                {"scope", scope },
                {"client_assertion_type", clientAssertion }
            };

            client.DefaultRequestHeaders.Add("Accept", "application/json");

            var content = new FormUrlEncodedContent(values);
            var response = await client.PostAsync(tokenRequestUrl, content);
            var responseString = await response.Content.ReadAsStringAsync();
            JObject test = JObject.Parse(responseString);
            return ((test["access_token"]).ToString());
        }

        /*
         * Create self signed certificate using Azure Key Vault
         * The code in this method currently uses the local pfx file, later on this will be modified
         */
        private string GetSignedClientAssertion()
        {
            X509Certificate2 selfSignedCertificate = new X509Certificate2(@"C:\Users\riagarwa\Downloads\oct5keyvault-ProdCert2-20201211.pfx", "", X509KeyStorageFlags.EphemeralKeySet);

            // AAD Prod Tenant ID: f8cdef31-a31e-4b4a-93e4-5f571e91255a
            string aud = $"https://login.microsoftonline.com/f8cdef31-a31e-4b4a-93e4-5f571e91255a/v2.0";
            // no need to add exp, nbf as JsonWebTokenHandler will add them by default.
            var claims = new Dictionary<string, object>()
            {
                { "aud", aud },
                { "iss", clientId },
                { "jti", Guid.NewGuid().ToString() },
                { "sub", clientId }
            };

            var securityTokenDescriptor = new SecurityTokenDescriptor
            {
                Claims = claims,
                SigningCredentials = new X509SigningCredentials(selfSignedCertificate)
            };

            JsonWebTokenHandler handler = new JsonWebTokenHandler();
            string signedClientAssertion = handler.CreateToken(securityTokenDescriptor);
            return signedClientAssertion;
        }

        private async Task<string> getSharePointMetadata(string accessToken, string actorToken, string url)
        {
            HttpClient client = new HttpClient();
            string authorizationRequestHeaderValue = "MSAuth1.0 actortoken=Bearer " + actorToken + ", accesstoken=Bearer " + accessToken + ", type=PFAT";

            client.DefaultRequestHeaders.Add("Authorization", authorizationRequestHeaderValue);
            HttpResponseMessage response = await client.GetAsync(url);
            response.EnsureSuccessStatusCode();
            var resp = await response.Content.ReadAsStringAsync();

            return resp;
        }
    }
}
