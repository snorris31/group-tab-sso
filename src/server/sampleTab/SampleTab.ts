import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/sampleTab/index.html")
@PreventIframe("/sampleTab/config.html")
@PreventIframe("/sampleTab/remove.html")
export class SampleTab {
}
