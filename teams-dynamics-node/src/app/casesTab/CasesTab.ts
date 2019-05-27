import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/casesTab/index.html")
@PreventIframe("/casesTab/config.html")
@PreventIframe("/casesTab/remove.html")
export class CasesTab {
}
