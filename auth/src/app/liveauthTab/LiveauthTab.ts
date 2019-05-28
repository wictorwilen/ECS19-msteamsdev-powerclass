import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/liveauthTab/index.html")
@PreventIframe("/liveauthTab/config.html")
@PreventIframe("/liveauthTab/remove.html")
export class LiveauthTab {
}
