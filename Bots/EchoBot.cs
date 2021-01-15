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
using AdaptiveCards.Templating;
using System.Text.RegularExpressions;
using System.Linq;

namespace HelloWorldBot.Bots
{
    public class EchoBot : TeamsActivityHandler
    {
        private const string aadLoginUrl = "https://login.microsoftonline.com/";
        private const string tokenRequestParam = "/oauth2/v2.0/token";
        private const string firstPartyAppID = "e5e15768-1702-474d-ba7b-904c7cad2bcf";
        private const string clientIdKey = "client_id";
        private const string clientId = "e5e15768-1702-474d-ba7b-904c7cad2bcf";
        private const string clientAssertionKey = "client_assertion";
        private const string clientAssertionTypeKey = "client_assertion_type";
        private const string clientAssertionTypeValue = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer";
        private const string clientCredentialsKey = "grant_type";
        private const string clientCredentialsValue = "client_credentials";
        private const string scopeKey = "scope";
        private const string authorizationKey = "Authorization";
        private const string acceptKey = "Accept";
        private const string acceptJsonVal = "application/json";
        private string tenantId = "";

        protected override async Task<MessagingExtensionResponse> OnTeamsAppBasedLinkQueryAsync(ITurnContext<IInvokeActivity> turnContext, AppBasedLinkQuery query, CancellationToken cancellationToken)
        {
            JObject tenant = JObject.FromObject(turnContext.Activity.ChannelData.tenant);
            if (tenant != null && tenant["id"] != null)
            {
                tenantId = tenant["id"].ToString();
            }

            if (turnContext != null && turnContext.Activity != null)
            {
                JObject valueObject = JObject.FromObject(turnContext.Activity.Value);
                if (valueObject["authentication"] != null)
                {
                    string accessToken = TransformPFTToken((valueObject["authentication"]["token"]).ToString());
                    Uri inputUrl = new Uri(valueObject["url"].ToString());

                    string actorToken = await GetActorToken(inputUrl);
                    JObject[] spMetadata = await GetSharePointMetadata(accessToken, actorToken, inputUrl);
                    return CreateAdaptiveCard(spMetadata);
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
         * */

        private string TransformPFTToken(String preTransformedPFTToken)
        {
            JwtSecurityToken jwtSecurityToken = new JwtSecurityToken(preTransformedPFTToken);
            JObject headerObject = JObject.Parse(Base64UrlEncoder.Decode(jwtSecurityToken.RawHeader));

            string nonceIn = headerObject.GetValue("nonce", StringComparison.OrdinalIgnoreCase).ToString();
            string hashName = headerObject.GetValue("alg", StringComparison.OrdinalIgnoreCase).ToString();

            HashAlgorithm hashAlgorithm = GetHashFunction(hashName);
            string nonceOut = Base64UrlEncoder.Encode(hashAlgorithm.ComputeHash(Encoding.UTF8.GetBytes(nonceIn)));
            headerObject["nonce"] = nonceOut;
            string newHeaderString = Base64UrlEncoder.Encode(headerObject.ToString(Formatting.None));
            string postTransformedToken = $"{newHeaderString}.{jwtSecurityToken.RawPayload}.{jwtSecurityToken.RawSignature}";
            return postTransformedToken;
        }

        private HashAlgorithm GetHashFunction(string hashAlgorithmName)
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

        private async Task<string> GetActorToken(Uri spUrl)
        {
            try
            {
                string scope = $"{spUrl.Scheme}://{spUrl.Host}/.default";
                string tokenRequestUrl = $"{aadLoginUrl}{tenantId}{tokenRequestParam}";

                HttpClient client = new HttpClient();
                client.BaseAddress = new Uri(tokenRequestUrl);
                string clientAssertionValue = GetSignedClientAssertion();

                Dictionary<string, string> values = new Dictionary<string, string>
                {
                    { clientCredentialsKey, clientCredentialsValue },
                    { clientIdKey, clientId },
                    {clientAssertionKey, clientAssertionValue },
                    {scopeKey, scope },
                    {clientAssertionTypeKey, clientAssertionTypeValue }
                };

                client.DefaultRequestHeaders.Add(acceptKey, acceptJsonVal);

                FormUrlEncodedContent content = new FormUrlEncodedContent(values);
                HttpResponseMessage responseMessage = await client.PostAsync(tokenRequestUrl, content);
                if (responseMessage.IsSuccessStatusCode)
                {
                    string responseString = await responseMessage.Content.ReadAsStringAsync();
                    JObject responseJson = JObject.Parse(responseString);
                    if (responseJson["access_token"] != null)
                    {
                        return ((responseJson["access_token"]).ToString());
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }

            return null;
        }

        /*
         * Create self signed certificate using Azure Key Vault
         * The code in this method currently uses the local pfx file, later on this will be modified
         */

        private string GetSignedClientAssertion()
        {
            string certThumbprint = "89F68251505D05923994B05B25AEB9E9954C2F68";
            bool validOnly = false;

            using (X509Store certStore = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                certStore.Open(OpenFlags.ReadOnly);

                X509Certificate2Collection certCollection = certStore.Certificates.Find(
                                            X509FindType.FindByThumbprint,
                                            // Replace below with your certificate's thumbprint
                                            certThumbprint,
                                            validOnly);
                // Get the first cert with the thumbprint
                X509Certificate2 selfSignedCertificate = certCollection.OfType<X509Certificate2>().FirstOrDefault();

                string aud = $"{aadLoginUrl}{tenantId}/v2.0";

                Dictionary<string, object> claims = new Dictionary<string, object>()
            {
                { "aud", aud },
                { "iss", clientId },
                { "jti", Guid.NewGuid().ToString() },
                { "sub", clientId }
            };

                SecurityTokenDescriptor securityTokenDescriptor = new SecurityTokenDescriptor
                {
                    Claims = claims,
                    SigningCredentials = new X509SigningCredentials(selfSignedCertificate)
                };

                JsonWebTokenHandler handler = new JsonWebTokenHandler();
                string signedClientAssertion = handler.CreateToken(securityTokenDescriptor);
                return signedClientAssertion;
            }
        }

        private async Task<JObject[]> GetSharePointMetadata(string accessToken, string actorToken, Uri spUrl)
        {
            var taskList = new[]
            {
                GetSiteContent(accessToken, actorToken, spUrl),
                GetSiteTitle(accessToken, actorToken, spUrl),
                //EchoBot.GetImageThumbnail(accessToken, actorToken, spUrl)
            };

            return await Task.WhenAll(taskList);
        }

        private async Task<JObject> GetSiteTitle(string accessToken, string actorToken, Uri url)
        {
            string teamSite = Regex.Split(url.ToString(), @"/sitepages/", RegexOptions.IgnoreCase)[0];
            string requestUrl = teamSite + "/_api/web/Title";

            HttpResponseMessage responseMessage = await MakePFTRequest(accessToken, actorToken, requestUrl);
            if (responseMessage.IsSuccessStatusCode)
            {
                string responseStr = await responseMessage.Content.ReadAsStringAsync();
                return JsonConvert.DeserializeObject<JObject>(responseStr);
            }
            return null;
        }

        private async Task<JObject> GetSiteContent(string accessToken, string actorToken, Uri url)
        {
            string teamSite = Regex.Split(url.ToString(), @"/sitepages/", RegexOptions.IgnoreCase)[0];
            string requestUrl = $"{teamSite}/_api/sitepages/pages/GetByUrl('{url.AbsolutePath}')";

            HttpResponseMessage responseMessage = await MakePFTRequest(accessToken, actorToken, requestUrl);
            if (responseMessage.IsSuccessStatusCode)
            {
                string responseStr = await responseMessage.Content.ReadAsStringAsync();
                return JsonConvert.DeserializeObject<JObject>(responseStr);
            }
            return null;
        }

        private async Task<JObject> GetImageThumbnail(string accessToken, string actorToken, Uri url)
        {
            string teamSite = Regex.Split(url.ToString(), @"/\/sitepages\//i", RegexOptions.IgnoreCase)[0];
            string requestUrl = $"{ url.Scheme}://{url.Host}/_api/v2.0/sharePoint:/{teamSite}:/driveItem/thumbnails/0/c71x40/content";

            HttpResponseMessage responseMessage = await MakePFTRequest(accessToken, actorToken, requestUrl);
            if (responseMessage.IsSuccessStatusCode)
            {
                byte[] responseStr = await responseMessage.Content.ReadAsByteArrayAsync();
                JObject obj = new JObject();
                obj["imageUrlInBase64"] = "data:image/jpeg;base64," + Convert.ToBase64String(responseStr);
                return obj;
            }
            return null;
        }

        private async Task<HttpResponseMessage> MakePFTRequest(string accessToken, string actorToken, string apiUrl)
        {
            HttpClient client = new HttpClient();
            string authorizationHeaderValue = "MSAuth1.0 actortoken=" + '"' +
                "Bearer " + actorToken + '"' + ", accesstoken=" + '"' + "Bearer " + accessToken + '"' + ", type=" + '"' + "PFAT" + '"';

            client.DefaultRequestHeaders.Add(authorizationKey, authorizationHeaderValue);
            client.DefaultRequestHeaders.Add(acceptKey, acceptJsonVal);

            return await client.GetAsync(apiUrl);
        }

        private MessagingExtensionResponse CreateAdaptiveCard(JObject[] spMetadata)
        {
            try
            {
                string cardTemplate = Path.Combine(".", "adaptiveCardSample.json");
                string cardContent = File.ReadAllText(cardTemplate);
                AdaptiveCardTemplate template = new AdaptiveCardTemplate(cardContent);

                var spData = new
                {
                    imageUrl = spMetadata[0]["BannerImageUrl"].ToString(),
                    siteName = spMetadata[0]["Title"].ToString(),
                    pageTitle = spMetadata[1]["value"].ToString(),
                    authorName = GetAuthorName(spMetadata[0]),
                    authorDate = spMetadata[0]["FirstPublished"].ToString()
                };

                string cardJson = template.Expand(spData);

                HeroCard previewCard = new HeroCard
                {
                    Title = spMetadata[1]["value"].ToString(),
                    Subtitle = spMetadata[0]["Title"].ToString(),
                    Text = "Sample text",
                };

                MessagingExtensionAttachment cardAttachment = CreateAdaptiveCardAttachment(cardJson, previewCard);

                MessagingExtensionResult result = new MessagingExtensionResult("list", "result", new[] { cardAttachment });
                return new MessagingExtensionResponse(result);
            }
            catch (AdaptiveSerializationException e)
            {
                throw e;
            }
        }

        private MessagingExtensionAttachment CreateAdaptiveCardAttachment(string cardJson, HeroCard previewCard)
        {
            MessagingExtensionAttachment adaptiveCardAttachment = new MessagingExtensionAttachment()
            {
                ContentType = "application/vnd.microsoft.card.adaptive",
                Content = JsonConvert.DeserializeObject(cardJson),
                Preview = previewCard.ToAttachment()
            };

            return adaptiveCardAttachment;
        }

        private String GetAuthorName(JObject spPageContent)
        {
            string authorName = "";
            try
            {
                if (spPageContent != null && spPageContent["LayoutWebpartsContent"] != null)
                {
                    JArray layoutwebpartsContentStr = JsonConvert.DeserializeObject<JArray>(spPageContent["LayoutWebpartsContent"].ToString());
                    if (layoutwebpartsContentStr != null && layoutwebpartsContentStr[0] != null && layoutwebpartsContentStr[0]["properties"] != null &&
                        layoutwebpartsContentStr[0]["properties"]["authors"] != null && layoutwebpartsContentStr[0]["properties"]["authors"][0] != null
                        && layoutwebpartsContentStr[0]["properties"]["authors"][0]["name"] != null)
                    {
                        authorName = layoutwebpartsContentStr[0]["properties"]["authors"][0]["name"].ToString();
                    }
                }
            }
            catch(Exception e)
            {
                return authorName;
            }
            return authorName;
        }
    }
}
