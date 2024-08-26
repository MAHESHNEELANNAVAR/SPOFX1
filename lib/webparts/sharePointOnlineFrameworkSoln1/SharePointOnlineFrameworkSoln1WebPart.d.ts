import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
export interface IsharePointOnlineFrameworkSoln1WebPartProps {
    description: string;
}
export default class sharePointOnlineFrameworkSoln1WebPart extends BaseClientSideWebPart<IsharePointOnlineFrameworkSoln1WebPartProps> {
    render(): void;
    protected onInit(): Promise<void>;
    private _getEnvironmentMessage;
    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=SharePointOnlineFrameworkSoln1WebPart.d.ts.map