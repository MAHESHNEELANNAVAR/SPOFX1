var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './SharePointOnlineFrameworkSoln1WebPart.module.scss';
import * as strings from 'SharePointOnlineFrameworkSoln1WebPartStrings';
var sharePointOnlineFrameworkSoln1WebPart = /** @class */ (function (_super) {
    __extends(sharePointOnlineFrameworkSoln1WebPart, _super);
    function sharePointOnlineFrameworkSoln1WebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    sharePointOnlineFrameworkSoln1WebPart.prototype.render = function () {
        this.domElement.innerHTML = "\n    <h1>welcome to https://www.maheshneelannavar.pro</h1>\n      <section class=\"".concat(styles.sharePointOnlineFrameworkSoln1, " ").concat(!!this.context.sdks.microsoftTeams ? styles.teams : '', "\">\n        <div class=\"").concat(styles.welcome, "\">\n          <p class=\"").concat(styles.welcome, "\">Title: ").concat(escape(this.context.pageContext.web.title), "</p>\n           <p class=\"").concat(styles.welcome, "\">Display Name: ").concat(escape(this.context.pageContext.user.displayName), "</p>\n        </div>\n      </section>");
    };
    sharePointOnlineFrameworkSoln1WebPart.prototype.onInit = function () {
        return this._getEnvironmentMessage().then(function () { });
    };
    sharePointOnlineFrameworkSoln1WebPart.prototype._getEnvironmentMessage = function () {
        var _this = this;
        if (!!this.context.sdks.microsoftTeams) { // Check if running in Teams, Office.com, or Outlook
            return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
                .then(function (context) {
                var environmentMessage = '';
                switch (context.app.host.name) {
                    case 'Office':
                        environmentMessage = _this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
                        break;
                    case 'Outlook':
                        environmentMessage = _this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
                        break;
                    case 'Teams':
                    case 'TeamsModern':
                        environmentMessage = _this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
                        break;
                    default:
                        environmentMessage = strings.UnknownEnvironment;
                }
                return environmentMessage;
            })
                .then(function () { });
        }
        return Promise.resolve();
    };
    sharePointOnlineFrameworkSoln1WebPart.prototype.onThemeChanged = function (currentTheme) {
        if (!currentTheme) {
            return;
        }
        var semanticColors = currentTheme.semanticColors;
        if (semanticColors) {
            this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || '');
            this.domElement.style.setProperty('--link', semanticColors.link || '');
            this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || '');
        }
    };
    Object.defineProperty(sharePointOnlineFrameworkSoln1WebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    sharePointOnlineFrameworkSoln1WebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return sharePointOnlineFrameworkSoln1WebPart;
}(BaseClientSideWebPart));
export default sharePointOnlineFrameworkSoln1WebPart;
//# sourceMappingURL=SharePointOnlineFrameworkSoln1WebPart.js.map