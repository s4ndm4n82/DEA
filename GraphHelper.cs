using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using TimeZoneConverter;

namespace DEA
{
    public class GraphHelper
    {
        private static DeviceCredentials tokenCredentials;
        private static GraphServiceClient graphClient;

        public static void Initialize(string clientID, string[] scopes,
                                      Func<DeviceCodeInfo, CancellationToken, Task> callBack)
        {
            tokenCredentials = new DeviceCredentials(callBack, clientID);
            graphClient = new GraphServiceClient(tokenCredentials, scopes);
        }

        public static async Task<string> GetAccessTokenAsync(string[] scopes)
        {
            var context = new TokenRequestContext(scopes);
            var response = await tokenCredential.GetTokenAsync(context);
            return response.Token;
        }
    }
}