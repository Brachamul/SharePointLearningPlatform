var __extends = this.__extends || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    __.prototype = b.prototype;
    d.prototype = new __();
};
var AzureDataPipe = (function () {
    function AzureDataPipe() {
        this.unittestFlag = false;
    }
    AzureDataPipe.prototype.getData = function (extensionType, hostConfig) {
        hostConfig.getDataCall(extensionType, SuiteExtensionsDataStore.dataPipeSuccessCallback, SuiteExtensionsDataStore.dataPipeFailCallback);
    };
    return AzureDataPipe;
})();
var AddInsFlights = (function () {
    function AddInsFlights() {
    }
    AddInsFlights.AzureDataPipe = 0;
    AddInsFlights.FileHandlerAddInPicker = 1;
    return AddInsFlights;
})();
var HostConfig = (function () {
    function HostConfig(userId, getDataCall) {
        this.enabledFlights = [];
        this.host = "Office365";
        this.logging = new Logging();
        this.cultureName = "en-us";
        this.resourceId = "http://office365.com/";
        this.set_logging(null);
        this.userId = userId;
        this.getDataCall = getDataCall;
    }
    HostConfig.prototype.set_logging = function (traceComponent) {
        switch (this.host) {
            case "SharePoint":
                this.logging = new SharePointLogging(traceComponent);
                break;
            case "OWA":
                this.logging = new OwaLogging(traceComponent);
                break;
            default:
        }
    };
    HostConfig.prototype.addFlight = function (flight) {
        this.enabledFlights.push(flight);
    };
    HostConfig.prototype.isFlightEnabled = function (flight) {
        return this.enabledFlights.indexOf(flight) >= 0;
    };
    return HostConfig;
})();

var Logging = (function () {
    function Logging() {
        this.prefix = null;
        this.traceComponent = null;
    }
    Logging.prototype.WriteDebugLog = function (logName, logLevel, message) {
    };
    Logging.prototype.WriteEngagementLog = function (logName, extensionType) {
    };
    Logging.prototype.WriteStart = function (startName) {
    };
    Logging.prototype.WriteSuccess = function (successName) {
    };
    Logging.prototype.WriteFailure = function (failureName) {
    };
    return Logging;
})();


var SharePointLogging = (function () {
    function SharePointLogging(traceComponent) {
        this.prefix = AddInType.FileHandler + "_";
        this.traceComponent = traceComponent;
    }
    SharePointLogging.prototype.WriteDebugLog = function (logName, logLevel, message) {
        WriteDebugLog(this.prefix + logName, logLevel, message);
    };
    SharePointLogging.prototype.WriteEngagementLog = function (logName, extensionType) {
        WriteEngagementLog(this.prefix + logName, extensionType);
    };
    SharePointLogging.prototype.WriteStart = function (startName) {
        WriteStart(this.prefix + startName);
    };
    SharePointLogging.prototype.WriteSuccess = function (successName) {
        WriteSuccess(this.prefix + successName);
    };
    SharePointLogging.prototype.WriteFailure = function (failureName) {
        WriteFailure(this.prefix + failureName);
    };
    return SharePointLogging;
})();

var OwaLogging = (function () {
    function OwaLogging(traceComponent) {
        this.prefix = AddInType.FileHandler + "_";
        this.traceComponent = traceComponent;
    }
    OwaLogging.prototype.WriteDebugLog = function (logName, logLevel, message) {
        if (logLevel == true) {
            _js.Trace.logError(this.traceComponent, "{0} {1}", this.prefix + logName, message);
        } else {
            _js.Trace.logWarning(this.traceComponent, "{0} {1}", this.prefix + logName, message);
        }
    };
    OwaLogging.prototype.WriteEngagementLog = function (logName, extensionType) {
        _js.Trace.logWarning(this.traceComponent, "{0} {1}", this.prefix + logName, extensionType);
    };
    OwaLogging.prototype.WriteStart = function (startName) {
        _js.Trace.logWarning(this.traceComponent, this.prefix + startName);
    };
    OwaLogging.prototype.WriteSuccess = function (successName) {
        _js.Trace.logWarning(this.traceComponent, this.prefix + successName);
    };
    OwaLogging.prototype.WriteFailure = function (failureName) {
        _js.Trace.logError(this.traceComponent, this.prefix + failureName);
    };
    return OwaLogging;
})();
function callInitializeIconControls() {
    if (typeof InitializeIconControls == 'function') {
        InitializeIconControls();
    } else
        setTimeout("callInitializeIconControls()", 300);
}
var ScriptOnDemand = (function () {
    function ScriptOnDemand() {
    }
    ScriptOnDemand.AddScriptToDocument = function (scriptPath, id) {
        var scriptElement = document.createElement('script');
        scriptElement.setAttribute("id", id);
        scriptElement.setAttribute("src", scriptPath);
        scriptElement.setAttribute("type", "text/javascript");
        document.getElementsByTagName('head')[0].appendChild(scriptElement);
        return scriptElement;
    };
    ScriptOnDemand.LoadControlsJavascript = function () {
        var dataStore = SuiteExtensionsDataStore.GetInstance();
        var controlsFileRef = document.getElementById("ControlScript");

        if (controlsFileRef == null && dataStore.hostConfig.host == "SharePoint") {
            controlsFileRef = ScriptOnDemand.AddScriptToDocument("/_layouts/15/online/scripts/SuiteExtensionsControls.js", "ControlScript");

            if (navigator.userAgent.indexOf('MSIE') !== -1 || navigator.appVersion.indexOf('Trident/') > 0) {
                controlsFileRef.onreadystatechange = function () {
                    if (this.readyState == 'complete' || this.readyState == 'loaded') {
                        InitializeIconControls();
                    }
                };
            } else {
                controlsFileRef.onload = function () {
                    InitializeIconControls();
                };
            }
        } else {
            callInitializeIconControls();
        }
    };
    ScriptOnDemand.LoadLocalizedStrings = function (hostConfig) {
        if (hostConfig != null && typeof hostConfig.localizedStringsPath === 'string' && hostConfig.localizedStringsPath != null && hostConfig.localizedStringsPath.length > 0 && (typeof Strings === 'undefined' || typeof Strings.CloudApps === 'undefined')) {
            var xmlHttp = new XMLHttpRequest();
            xmlHttp.onreadystatechange = function () {
                if (xmlHttp.readyState === 4) {
                    if (xmlHttp.status === 200) {
                        hostConfig.logging.WriteSuccess("SuccessLoadLocalizedStrings");
                        ScriptOnDemand.AddScriptToDocument(hostConfig.localizedStringsPath, "CloudAppsLocalizationScript");
                    } else {
                        hostConfig.logging.WriteFailure("FailedLoadLocalizedStrings");
                    }
                }
            };
            xmlHttp.open("GET", hostConfig.localizedStringsPath, true);
            xmlHttp.send("");
            hostConfig.logging.WriteStart("StartLoadLocalizedStrings");
        }
    };
    return ScriptOnDemand;
})();

var SuiteExtensionsLocalStorage = (function () {
    function SuiteExtensionsLocalStorage() {
        if (SuiteExtensionsLocalStorage.instance) {
            throw new Error("Error: Instantiation failed: Use SuiteExtensionsLocalStorage.GetInstance() instead of new.");
        }
        SuiteExtensionsLocalStorage.instance = this;
        var last = this.GetItemInternal("lastPage") || "";
        var current = location.href;
        if (last === current) {
            this.lastUpdatedTime = new Date().getTime() - SuiteExtensionsLocalStorage.MaximumExpiredPeriod - 1;
        }
        this.SetItemInternal("lastPage", current);
    }
    SuiteExtensionsLocalStorage.GetInstance = function () {
        if (!SuiteExtensionsLocalStorage.Supported()) {
            return null;
        }
        if (SuiteExtensionsLocalStorage.instance === null) {
            SuiteExtensionsLocalStorage.instance = new SuiteExtensionsLocalStorage();
        }
        return SuiteExtensionsLocalStorage.instance;
    };
    SuiteExtensionsLocalStorage.prototype.SetItem = function (extensionType, jsonString) {
        this.SetItemInternal(this.GetExtensionTypeKey(extensionType), jsonString);
    };
    SuiteExtensionsLocalStorage.prototype.GetItem = function (extensionType) {
        return this.GetItemInternal(this.GetExtensionTypeKey(extensionType));
    };
    SuiteExtensionsLocalStorage.prototype.SetItemObject = function (extensionType, jsonObject) {
        this.SetItem(extensionType, JSON.stringify(jsonObject));
    };
    SuiteExtensionsLocalStorage.prototype.GetItemObject = function (extensionType) {
        return JSON.parse(this.GetItem(extensionType));
    };
    SuiteExtensionsLocalStorage.prototype.GetExpirationKey = function (extensionType) {
        return this.GetExtensionTypeKey(extensionType) + ".LocalStorageSetTime";
    };

    SuiteExtensionsLocalStorage.prototype.IsExpired = function (extensionType) {
        var value = this.GetItemInternal(this.GetExpirationKey(extensionType));
        if (value == null || value == undefined) {
            return true;
        } else {
            this.lastUpdatedTime = +value;
            var currentTime = new Date().getTime();
            if ((currentTime - this.lastUpdatedTime) >= SuiteExtensionsLocalStorage.MaximumExpiredPeriod) {
                return true;
            } else {
                return false;
            }
        }
    };

    SuiteExtensionsLocalStorage.Supported = function () {
        try  {
            localStorage.setItem("Test", "2");
            localStorage.removeItem("Test");
            return true;
        } catch (e) {
            return false;
        }
    };
    SuiteExtensionsLocalStorage.prototype.GetExtensionTypeKey = function (extensionType) {
        return "Office365.AddIns." + extensionType + "." + SuiteExtensionsDataStore.GetInstance().hostConfig.userId;
    };

    SuiteExtensionsLocalStorage.prototype.SetItemInternal = function (key, value) {
        localStorage.setItem(key, value);
        var localStorageSetValue = new Date().getTime().toString();
        localStorage.setItem(key + ".LocalStorageSetTime", localStorageSetValue);
    };
    SuiteExtensionsLocalStorage.prototype.GetItemInternal = function (key) {
        return localStorage.getItem(key);
    };
    SuiteExtensionsLocalStorage.MaximumExpiredPeriod = 7200000;
    SuiteExtensionsLocalStorage.instance = null;
    return SuiteExtensionsLocalStorage;
})();

var TestDataPipe = (function () {
    function TestDataPipe() {
    }
    TestDataPipe.prototype.getData = function (extensionType, hostConfig) {
        var extType = extensionType;
        var jsonManifest = [
            {
                'type': 'FileHandler',
                'appId': '4463c52c-491f-4559-a2b1-8b688fca9eb9',
                'displayName': 'GPX App',
                'addInId': '12b4c4f2-d1c5-c12a-6245-1b5c3fdd85b1',
                'properties': {
                    'extension': 'gpx',
                    'fileIcon': 'https://gpxfileapp-int.azurewebsites.net/content/pinpoint-16x16.png',
                    'openUrl': 'https://gpxfileapp-int.azurewebsites.net/FileHandler/Open',
                    'previewUrl': 'https://gpxfileapp-int.azurewebsites.net/FileHandler/Preview'
                }
            },
            {
                'type': 'FileHandler',
                'appId': '6f7b9f0c-e0f9-41f9-9bff-bcf66d599c2b',
                'displayName': 'TEST or INI App',
                'addInId': '5f0b526b-f1e4-4def-9d43-3b0d28d4ff53',
                'properties': {
                    'extension': 'test',
                    'fileIcon': 'https://testorinifileapp-int.azurewebsites.net/content/test.png',
                    'openUrl': 'https://testorinifileapp-int.azurewebsites.net/FileHandler/Open'
                }
            },
            {
                'type': 'FileHandler',
                'appId': '6f7b9f0c-e0f9-41f9-9bff-bcf66d599c2b',
                'displayName': 'TEST or INI App',
                'addInId': 'e94ca6fe-66c3-4212-b4db-3c32c2b00d85',
                'properties': {
                    'extension': 'ini',
                    'previewUrl': 'https://testorinifileapp-int.azurewebsites.net/FileHandler/Preview'
                }
            },
            {
                'type': 'FileHandler',
                'appId': 'f02cf431-6830-4090-9e3e-f8526e676852',
                'displayName': 'DRW App',
                'addInId': 'b0497361-9dff-4fcf-a8ac-ae11b5f27f1c',
                'properties': {
                    'extension': 'drw',
                    'openUrl': 'https://drwfileapp-int.azurewebsites.net/FileHandler/Open',
                    'previewUrl': 'https://drwfileapp-int.azurewebsites.net/FileHandler/Preview'
                }
            },
            {
                'type': 'FileHandler',
                'appId': '523f34b4-200f-4239-aa3d-67682bf24bdc',
                'displayName': 'OFF App',
                'addInId': '55d5929b-e6d3-4518-8297-98078e4f0ee8',
                'properties': {
                    'extension': 'off',
                    'fileIcon': 'https://offfileapp-int.azurewebsites.net/content/off.png',
                    'openUrl': 'https://offfileapp-int.azurewebsites.net/FileHandler/Open'
                }
            },
            {
                'type': 'FileHandler',
                'appId': 'f266399a-ab8d-4407-9f2d-797a35a75d3f',
                'displayName': 'OFF or OFFICE App',
                'addInId': '5afb40fd-3604-4cbe-b991-82d01a66a29b',
                'properties': {
                    'extension': 'off',
                    'fileIcon': 'https://offorofficefileapp-int.azurewebsites.net/content/off-mt.png',
                    'previewUrl': 'https://offorofficefileapp-int.azurewebsites.net/FileHandler/Preview'
                }
            },
            {
                'type': 'FileHandler',
                'appId': 'f266399a-ab8d-4407-9f2d-797a35a75d3f',
                'displayName': 'OFF or OFFICE App',
                'addInId': '3a3ff70e-23c7-4dae-8bad-43c314a98b82',
                'properties': {
                    'extension': 'office',
                    'fileIcon': 'https://offorofficefileapp-int.azurewebsites.net/content/office.png',
                    'previewUrl': 'https://offorofficefileapp-int.azurewebsites.net/FileHandler/Preview',
                    'openUrl': 'https://offorofficefileapp-int.azurewebsites.net/FileHandler/Open'
                }
            },
            {
                'type': 'FileHandler',
                'appId': 'b565d800-242f-4eef-aa2e-1b5f09529a94',
                'displayName': 'EXT App',
                'addInId': '0c21f706-6071-4b2c-9273-59ea1f486eb9',
                'properties': {
                    'extension': 'ext1',
                    'previewUrl': 'https://extfileapp-int.azurewebsites.net/FileHandler/Preview',
                    'openUrl': 'https://extfileapp-int.azurewebsites.net/FileHandler/Open'
                }
            },
            {
                'type': 'FileHandler',
                'appId': 'b565d800-242f-4eef-aa2e-1b5f09529a94',
                'displayName': 'EXT App',
                'addInId': '88551be5-def0-4aeb-b2b6-6bf2e259f3ea',
                'properties': {
                    'extension': 'ext2',
                    'fileIcon': 'https://extfileapp-int.azurewebsites.net/content/ext2.png',
                    'openUrl': 'https://extfileapp-int.azurewebsites.net/FileHandler/Open'
                }
            },
            {
                'type': 'FileHandler',
                'appId': 'b565d800-242f-4eef-aa2e-1b5f09529a94',
                'displayName': 'EXT App',
                'addInId': 'da4b51ef-4b12-434a-b79d-0e930db6f133',
                'properties': {
                    'extension': 'ext3',
                    'fileIcon': 'https://extfileapp-int.azurewebsites.net/content/ext3.png',
                    'previewUrl': 'https://extfileapp-int.azurewebsites.net/FileHandler/Preview'
                }
            },
            {
                'type': 'FileHandler',
                'appId': 'b565d800-242f-4eef-aa2e-1b5f09529a94',
                'displayName': 'EXT App',
                'addInId': '92f9bcf2-aa0f-4a68-8714-c648af3fb715',
                'properties': {
                    'extension': 'ext4',
                    'fileIcon': 'https://extfileapp-int.azurewebsites.net/content/ext4.png'
                }
            },
            {
                'type': 'FileHandler',
                'appId': 'b565d800-242f-4eef-aa2e-1b5f09529a94',
                'displayName': 'EXT App',
                'addInId': '7fed2d59-55b3-490d-819c-474dfd772900',
                'properties': {
                    'extension': 'ext5',
                    'previewUrl': 'https://extfileapp-int.azurewebsites.net/FileHandler/Preview',
                    'openUrl': 'https://extfileapp-int.azurewebsites.net/FileHandler/Open'
                }
            },
            {
                'type': 'FileHandler',
                'addInId': '0436ffad-8131-4307-963b-9595ebb34e85',
                'displayName': 'Code View App',
                'appId': '41e0235f-7bc4-4b32-83ec-2f4a7951be84',
                'properties': {
                    'extension': 'txa;txb;txc',
                    'fileIcon': 'https://codeviewapp.azurewebsites.net/content/fileicon.png',
                    'openUrl': 'https://codeviewapp.azurewebsites.net/filehandler/open',
                    'previewUrl': 'https://codeviewapp.azurewebsites.net/filehandler/preview'
                }
            }
        ];
        var manifestString = JSON.stringify(jsonManifest);
        SuiteExtensionsDataStore.dataPipeSuccessCallback(extType, manifestString);
    };
    return TestDataPipe;
})();

var SuiteExtensionsDataStore = (function () {
    function SuiteExtensionsDataStore() {
        this.extensionsDictionary = {};
        if (SuiteExtensionsDataStore._instance) {
            throw new Error("Error: Instantiation failed: Use SuiteExtensionsDataStore.GetInstance() instead of new.");
        }
        SuiteExtensionsDataStore._instance = this;
        this.localStorage = SuiteExtensionsLocalStorage.GetInstance();
        this.useAzureDataPipe = false;
    }
    SuiteExtensionsDataStore.Initialize = function (hostConfig) {
        ScriptOnDemand.LoadLocalizedStrings(hostConfig);
        var useAzureDataPipe = hostConfig.isFlightEnabled(AddInsFlights.AzureDataPipe);
        var dataStoreInstance = SuiteExtensionsDataStore.GetInstance();
        dataStoreInstance.set_host(hostConfig);
        dataStoreInstance.SetDataPipe(useAzureDataPipe);
        dataStoreInstance.GetDataFromLocalStorage();
    };
    SuiteExtensionsDataStore.GetInstance = function () {
        if (SuiteExtensionsDataStore._instance === null) {
            SuiteExtensionsDataStore._instance = new SuiteExtensionsDataStore();
        }
        return SuiteExtensionsDataStore._instance;
    };
    SuiteExtensionsDataStore.dataPipeSuccessCallback = function (extensionType, extensionData) {
        var dataStore = SuiteExtensionsDataStore.GetInstance();
        try  {
            dataStore.hostConfig.logging.WriteDebugLog("StartSetLocalStorage", false, "Starting to set local storage with add-in data");
            dataStore.hostConfig.logging.WriteStart("StartSetLocalStorage");
            dataStore.localStorage.SetItem(extensionType, extensionData);
            dataStore.hostConfig.logging.WriteSuccess("SuccessSetLocalStorage");
            dataStore.hostConfig.logging.WriteDebugLog("SuccessSetLocalStorage", false, "Successfully set local storage with add-in data");
        } catch (e) {
            dataStore.hostConfig.logging.WriteFailure("FailedSetLocalStorage");
            dataStore.hostConfig.logging.WriteDebugLog("FailedSetLocalStorage", true, "Failed to set local storage with add-in data in this browser " + e.message);
        }
        dataStore.GetDataFromStorageAndLoadControlsJS(extensionType, "StartLoadControlsScriptOnNewData", "SuccessLoadControlsScriptOnNewData");
    };
    SuiteExtensionsDataStore.dataPipeFailCallback = function (e) {
        SuiteExtensionsDataStore.GetInstance().hostConfig.logging.WriteDebugLog("FailedAzureDataPipe", true, "Failed to get the add-in data from azure" + e.message);
    };
    SuiteExtensionsDataStore.prototype.SetDataPipe = function (useAzureDataPipe) {
        this.useAzureDataPipe = useAzureDataPipe;
        if (this.useAzureDataPipe === true) {
            this.dataPipe = new AzureDataPipe();
        } else {
            this.dataPipe = new TestDataPipe();
        }
    };
    SuiteExtensionsDataStore.prototype.set_host = function (hostConfig) {
        this.hostConfig = hostConfig;
    };
    SuiteExtensionsDataStore.prototype.GetDataFromLocalStorage = function () {
        var extensionType = AddInType.FileHandler;

        if (this.localStorage != null) {
            this.GetDataFromStorageAndLoadControlsJS(extensionType, "StartLoadControlsScript", "SuccessLoadControlsScript");

            if (this.localStorage.IsExpired(extensionType)) {
                this.dataPipe.getData(extensionType, this.hostConfig);
            }
        } else {
            this.hostConfig.logging.WriteEngagementLog("LocalStorageNotSupportedInBrowser", null);
            var browserInfo = navigator.appCodeName + navigator.appName + navigator.appVersion;
            this.hostConfig.logging.WriteDebugLog("LocalStorageNotSupportedInBrowser", true, "Local Storage is not supported in this browser: " + browserInfo);
        }
    };
    SuiteExtensionsDataStore.prototype.GetExtensions = function (extensionType) {
        if (this.extensionsDictionary[extensionType]) {
            return this.extensionsDictionary[extensionType];
        }
        return null;
    };
    SuiteExtensionsDataStore.prototype.GetDataFromStorageAndLoadControlsJS = function (extensionType, loadControlsJSLogStartMsg, loadControlsJSLogEndMsg) {
        if (this.localStorage != null) {
            try  {
                this.hostConfig.logging.WriteDebugLog("StartGetLocalStorageItem", false, "Trying to get add-ins from local storage");
                this.hostConfig.logging.WriteStart("StartGetLocalStorageItem");
                var extensionTypeValue = this.localStorage.GetItemObject(extensionType);
                this.hostConfig.logging.WriteSuccess("SuccessGetLocalStorageItem");
                this.hostConfig.logging.WriteDebugLog("SuccessGetLocalStorageItem", false, "Successfully retrieved add-ins from local storage");
            } catch (e) {
                this.hostConfig.logging.WriteFailure("FailedGetLocalStorageItem");
                this.hostConfig.logging.WriteDebugLog("FailedGetLocalStorageItem", true, "Failed to retrive add in from local storage: " + e.message);
            }
            if (extensionTypeValue != null && extensionTypeValue.length > 0) {
                this.extensionsDictionary[extensionType] = extensionTypeValue;
                this.hostConfig.logging.WriteDebugLog(loadControlsJSLogStartMsg, false, "Trying to load suite extensions controls javascript");
                this.hostConfig.logging.WriteStart(loadControlsJSLogStartMsg);
                ScriptOnDemand.LoadControlsJavascript();
                this.hostConfig.logging.WriteSuccess(loadControlsJSLogEndMsg);
                this.hostConfig.logging.WriteDebugLog(loadControlsJSLogEndMsg, false, "Successfully loaded suite extensions controls javascript");
            }
        }
    };
    SuiteExtensionsDataStore.prototype.getAddIns = function (addInType, addInFilter) {
        var addIns = [];
        var extensions = this.GetExtensions(addInType);
        var addInIndex = 0;
        if (extensions != null) {
            for (var i = 0; i < extensions.length; i++) {
                var extension = extensions[i];
                if (extension["type"] === addInType) {
                    var fileHandlerAddIn = new FileHandlerAddIn(extension);
                    if (addInFilter.IsMatch(fileHandlerAddIn)) {
                        addIns[addInIndex] = fileHandlerAddIn;
                        addInIndex++;
                    }
                }
            }
        }
        return addIns;
    };
    SuiteExtensionsDataStore._instance = null;
    return SuiteExtensionsDataStore;
})();
var ControlType = (function () {
    function ControlType() {
    }
    ControlType.Icon = 0;
    ControlType.Preview = 1;
    ControlType.Edit = 2;
    return ControlType;
})();
var ControlCreationProperty = (function () {
    function ControlCreationProperty() {
    }
    ControlCreationProperty.FileExtension = 'fileExtension';
    return ControlCreationProperty;
})();
var ControlFactory = (function () {
    function ControlFactory() {
    }
    ControlFactory.CreateControl = function (fileExtension, controlType) {
        var dataStore = SuiteExtensionsDataStore.GetInstance();
        var debugMessage = "Creating File Handler Control: [extension=" + fileExtension + "][control=" + controlType + "]";
        dataStore.hostConfig.logging.WriteDebugLog("CreatingControl", false, debugMessage);
        var propertyBag = new Object();
        propertyBag[ControlCreationProperty.FileExtension] = fileExtension;
        var addInFilter = new AddInFilter();
        addInFilter.add(FileHandlerAddInPropertyName.Extension, AddInFilterOperation.Contains, fileExtension);
        if (!dataStore.hostConfig.isFlightEnabled(AddInsFlights.FileHandlerAddInPicker)) {
            dataStore.hostConfig.logging.WriteDebugLog("CreatingControl", false, "Adding extra addin filters for property 'Not null'");
            switch (controlType) {
                case ControlType.Icon:
                    addInFilter.add(FileHandlerAddInPropertyName.FileIcon, AddInFilterOperation.NotNull, "");
                    break;
                case ControlType.Preview:
                    addInFilter.add(FileHandlerAddInPropertyName.PreviewUrl, AddInFilterOperation.NotNull, "");
                    break;
                case ControlType.Edit:
                    addInFilter.add(FileHandlerAddInPropertyName.OpenUrl, AddInFilterOperation.NotNull, "");
                    break;
                default:
                    break;
            }
        }
        var addIns = dataStore.getAddIns(AddInType.FileHandler, addInFilter);
        dataStore.hostConfig.logging.WriteDebugLog("CreatingControl", false, "Found " + (addIns == null ? 0 : addIns.length) + "AddIn");
        if (addIns != null && addIns.length > 0) {
            return ControlFactory.CreateControlCommon(addIns, controlType, propertyBag);
        } else {
            return null;
        }
    };

    ControlFactory.CreateControlCommon = function (addIns, controlType, propertyBag) {
        if (typeof SuiteExtensionsControl != 'undefined') {
            var dataStore = SuiteExtensionsDataStore.GetInstance();
            var debugMessage = "Creating Common Control:[control=" + controlType + "]";
            dataStore.hostConfig.logging.WriteDebugLog("CreatingCommonControl", false, debugMessage);
            if (addIns.length > 0 && addIns[0].type === AddInType.FileHandler) {
                var fileHandlerAddIns = addIns;
                switch (controlType) {
                    case ControlType.Icon:
                        var fileExtension = propertyBag[ControlCreationProperty.FileExtension];
                        return new IconControl(fileHandlerAddIns, fileExtension);
                    case ControlType.Preview:
                        var fileExtension = propertyBag[ControlCreationProperty.FileExtension];
                        var resourceId = dataStore.hostConfig.resourceId;
                        var cultureName = dataStore.hostConfig.cultureName;
                        var host = dataStore.hostConfig.host;
                        return new PreviewControl(fileHandlerAddIns, fileExtension, resourceId, cultureName, host);
                    case ControlType.Edit:
                        var fileExtension = propertyBag[ControlCreationProperty.FileExtension];
                        var resourceId = dataStore.hostConfig.resourceId;
                        var cultureName = dataStore.hostConfig.cultureName;
                        var host = dataStore.hostConfig.host;
                        return new EditControl(fileHandlerAddIns, fileExtension, resourceId, cultureName, host);
                    default:
                        dataStore.hostConfig.logging.WriteDebugLog("CreatingCommonControl", true, "Invalid Control Type");
                        return null;
                }
            }
        }
        return null;
    };
    return ControlFactory;
})();
var AddInType = (function () {
    function AddInType() {
    }
    AddInType.FileHandler = "FileHandler";
    return AddInType;
})();
var AddIn = (function () {
    function AddIn(extension) {
        this.addInId = extension[AddInPropertyName.AddInId];
        this.type = extension[AddInPropertyName.Type];
        this.appId = extension[AddInPropertyName.AppId];
        this.displayName = extension[AddInPropertyName.DisplayName];
        this.id = this.appId + "_" + this.addInId;
    }
    return AddIn;
})();
var AddInPropertyName = (function () {
    function AddInPropertyName() {
    }
    AddInPropertyName.AddInId = "addInId";
    AddInPropertyName.Type = "type";
    AddInPropertyName.AppId = "appId";
    AddInPropertyName.DisplayName = "displayName";
    return AddInPropertyName;
})();
var AddInFilterOperation = (function () {
    function AddInFilterOperation() {
    }
    AddInFilterOperation.Equals = 1;
    AddInFilterOperation.NotNull = 2;
    AddInFilterOperation.Contains = 3;
    return AddInFilterOperation;
})();
var FilterClause = (function () {
    function FilterClause(name, operation, value) {
        this.name = name;
        this.operation = operation;
        this.value = value;
    }
    FilterClause.prototype.IsMatch = function (addIn) {
        var testValue = addIn[this.name];
        switch (this.operation) {
            case AddInFilterOperation.Equals:
                if (testValue == undefined) {
                    return false;
                }
                if (typeof testValue == "object") {
                    throw new Error("Error: FilterClause: the property to test for Equals shouldn't be an object.");
                }
                return (testValue == this.value ? true : false);
            case AddInFilterOperation.NotNull:
                return (testValue != undefined ? true : false);
                break;
            case AddInFilterOperation.Contains:
                if (testValue == undefined) {
                    return false;
                }
                if (typeof testValue != "object" || typeof testValue.indexOf != "function") {
                    throw new Error("Error: FilterClause: the property to test for Contains should be an Array.");
                }
                return (testValue.indexOf(this.value) >= 0 ? true : false);
                break;
            default:
                return false;
        }
    };
    return FilterClause;
})();
var AddInFilter = (function () {
    function AddInFilter() {
        this.clauses = [];
    }
    AddInFilter.prototype.add = function (name, operation, value) {
        this.clauses.push(new FilterClause(name, operation, value));
    };
    AddInFilter.prototype.IsMatch = function (addIn) {
        var result = true;
        for (var index in this.clauses) {
            result = result && this.clauses[index].IsMatch(addIn);
        }
        return result;
    };
    return AddInFilter;
})();
var FileHandlerAddIn = (function (_super) {
    __extends(FileHandlerAddIn, _super);
    function FileHandlerAddIn(extension) {
        _super.call(this, extension);
        var extensionString = extension["properties"][FileHandlerAddInPropertyName.Extension];
        if (extensionString != undefined) {
            this.extension = extensionString.split(";").filter(function (e) {
                return e.length > 0;
            });
        } else {
            this.extension = undefined;
        }
        this.fileIcon = extension["properties"][FileHandlerAddInPropertyName.FileIcon];
        this.openUrl = extension["properties"][FileHandlerAddInPropertyName.OpenUrl];
        this.previewUrl = extension["properties"][FileHandlerAddInPropertyName.PreviewUrl];
    }
    return FileHandlerAddIn;
})(AddIn);
var FileHandlerAddInPropertyName = (function (_super) {
    __extends(FileHandlerAddInPropertyName, _super);
    function FileHandlerAddInPropertyName() {
        _super.apply(this, arguments);
    }
    FileHandlerAddInPropertyName.Extension = "extension";
    FileHandlerAddInPropertyName.FileIcon = "fileIcon";
    FileHandlerAddInPropertyName.OpenUrl = "openUrl";
    FileHandlerAddInPropertyName.PreviewUrl = "previewUrl";
    return FileHandlerAddInPropertyName;
})(AddInPropertyName);
