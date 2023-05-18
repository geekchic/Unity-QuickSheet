///////////////////////////////////////////////////////////////////////////////
///
/// GoogleDataSettings.cs
/// 
/// (c)2013 Kim, Hyoun Woo
///
///////////////////////////////////////////////////////////////////////////////
using UnityEngine;
using UnityEditor;
using System.Collections;
using System.IO;
using System.Collections.Generic;
using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;

namespace UnityQuickSheet
{
    /// <summary>
    /// A class manages google account setting.
    /// </summary>
    [CreateAssetMenu(menuName = "QuickSheet/Setting/GoogleData Setting")]
    public class GoogleDataSettings : SingletonScriptableObject<GoogleDataSettings>
    {
        // A flag which indicates use local installed oauth2 json file for authentication or not.
        static public bool useOAuth2JsonFile = false;

        public string JsonFilePath
        {
            get { return jsonFilePath; }
            set
            {
                if (string.IsNullOrEmpty(value) == false)
                    jsonFilePath = value;
            }
        }
        private string jsonFilePath = string.Empty;

        /// <summary>
        /// A default path where .txt template files are.
        /// </summary>
        public string TemplatePath = "QuickSheet/GDataPlugin/Templates";

        /// <summary>
        /// A path where generated ScriptableObject derived class and its data class script files are to be put.
        /// </summary>
        public string RuntimePath = string.Empty;

        /// <summary>
        /// A path where generated editor script files are to be put.
        /// </summary>
        public string EditorPath = string.Empty;

        public string ScriptableObjectPath = string.Empty;

        [System.Serializable]
        public struct OAuth2JsonData
        {
            public string type;
            public string project_id;
            public string private_key_id;
            public string private_key;
            public string client_email;
            public string client_id;
            public string auth_uri;
            public string token_uri;
            public string auth_provider_x509_cert_url;
            public string client_x509_cert_url;
            public string universe_domain;
        };

        public OAuth2JsonData OAuth2Data;

        private SheetsService service;
        public SheetsService Service 
        {
            get 
            {
                if (service == null)
                {
                    Stream jsonCreds = (Stream)File.Open(JsonFilePath, FileMode.Open);

                    ServiceAccountCredential credential = ServiceAccountCredential.FromServiceAccountData(jsonCreds);

                    service = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                    });
                }
                
                return service;
            }
        }

        // enter Access Code after getting it from auth url
        //public string _AccessCode = "Paste AcecessCode here!";

        //// enter Auth 2.0 Refresh Token and AccessToken after succesfully authorizing with Access Code
        //public string _RefreshToken = "";

        //public string _AccessToken = "";

        /// <summary>
        /// Select currently exist account setting asset file.
        /// </summary>
        [MenuItem("Edit/QuickSheet/Select Google Data Setting")]
        public static void Edit()
        {
            Selection.activeObject = Instance;
            if (Selection.activeObject == null)
            {
                Debug.LogError("No GoogleDataSettings.asset file is found. Create setting file first.");
            }
        }
    }
}