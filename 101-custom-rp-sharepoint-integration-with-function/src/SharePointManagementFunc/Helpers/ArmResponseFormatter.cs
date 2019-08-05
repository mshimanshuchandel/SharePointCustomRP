using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointManager.Helpers
{
    public class ArmResource
    {
        [JsonProperty]
        public string name { get; set; }
        [JsonProperty]
        public string type { get; set; }
        [JsonProperty]
        public object properties { get; set; }

        public ArmResource(string resourceName, string type, string subscriptionId, string resourceGroupName, string providerName, string resourceType, object properties)
        {
            this.name = resourceName;
            this.type = type;
            this.properties = properties;
        }
    }

    public class ArmProperty
    {
        [JsonProperty]
        public string type { get; set; }
        [JsonProperty]
        public object properties { get; set; }

        public ArmProperty(object properties, string resourceType)
        {
            this.type = resourceType;
            this.properties = properties;
        }
    }
}
