import { SPComponentLoader } from "@microsoft/sp-loader";
import { Guid } from "@microsoft/sp-core-library";

export default class ComponentHelper{
    public static loadSPComponentById(componentId: string) {
        return new Promise((resolve) => {
            SPComponentLoader.loadComponentById(componentId).then((component: any) => {
                resolve(component);
            });
        });
    }

    public static getTenantUrl(absoluteUrl: string): string {
        if (absoluteUrl.indexOf('/sites') === -1)
            return absoluteUrl;
        else return absoluteUrl.substring(0, absoluteUrl.indexOf('/sites'));
    }

    public static getThumbnailUrl(absoluteUrl: string, siteId: Guid, itemId: string, size?: number) {
        return `${ComponentHelper.getTenantUrl(absoluteUrl)}/_api/v2.1/sites/${siteId}/drive/items/${itemId}/thumbnails/0/c${size || 400}x99999/content?preferNoRedirect=true&clientType=modernWebPart`;
    };
}