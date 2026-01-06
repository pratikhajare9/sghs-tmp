import { BaseApplicationCustomizer, PlaceholderName } from '@microsoft/sp-application-base';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";


export interface IGlobalAlertBannerProperties { }

export default class GlobalAlertBannerApplicationCustomizer
  extends BaseApplicationCustomizer<IGlobalAlertBannerProperties> {

  private sp : any;//spfi().using(SPFx(this.context));

  public async onInit(): Promise<void> {

    await super.onInit();
    this.sp = spfi().using(SPFx(this.context));

    const placeholder = this.context.placeholderProvider.tryCreateContent(
      PlaceholderName.Top
    );
    if (!placeholder) return;

    const alert = await this.getActiveAlert();
    if (!alert) return;

    // Dismiss per session + per alert
    if (sessionStorage.getItem(`global-alert-${alert.Id}`)) return;

    placeholder.domElement.innerHTML = this.renderBanner(alert);
    this.bindDismiss(alert.Id);

    return Promise.resolve();
  }

  private async getActiveAlert() {
    const now = new Date().toISOString();

    const items = await this.sp.web.lists
      .getByTitle("Global Alerts")
      .items
      .select("Id", "Title", "AlertBody", "Severity", "LearnMoreUrl")
      .filter(`
        IsActive eq 1
      `)
      .top(1)();

    return items.length ? items[0] : null;
  }

  private renderBanner(alert: any): string {

    const bgColor =
      alert.Severity === "Critical" ? "#9B1C1C" :
        alert.Severity === "Warning" ? "#B45309" :
          "#1D4ED8";

    return `<style>
      div#globalAlertBanner div, div#globalAlertBanner a {color:white!important; font-family:Georgia, 'Times New Roman', Times, serif}
      div#globalAlertBannerBody * {font-size: 18px !important;}
    </style>
      <div id="globalAlertBanner" style="
        background:${bgColor};
        color:white;
        padding:5px 20px;
      ">
        <div id="globalAlertBannerContent" style="
          width:100%;
          margin:0 auto;
          display:flex;
          gap:20px;
        ">
          <div style="flex:1">
            <div id="globalAlertBannerTitle" style="
              font-size:22px;
              margin:8px 0;
            ">
              <p style="margin:0">${alert.Title}</p>
            </div>

            <div id="globalAlertBannerBody" style="
              font-size:18px;              
              color:white !important;
            ">
              <p style="margin:0">${alert.AlertBody || ""}</p>
            </div>
          </div>

          <button id="dismissAlert" title="Dismiss"
            style="
              background:none;
              border:none;
              color:white;
              font-size:20px;
              cursor:pointer;
              align-self:flex-start;
            ">
            ✕
          </button>
        </div>
      </div>
    `;
  }

  private bindDismiss(alertId: number): void {
    document.getElementById("dismissAlert")?.addEventListener("click", () => {
      sessionStorage.setItem(`global-alert-${alertId}`, "true");
      location.reload();
    });
  }
}
