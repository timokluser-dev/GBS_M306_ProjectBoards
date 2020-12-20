import { Component, OnInit } from "@angular/core";
import { BroadcastService, MsalService } from "@azure/msal-angular";
import { Logger, CryptoUtils } from "msal";

@Component({
  selector: "app-root",
  templateUrl: "./app.component.html",
  styleUrls: ["./app.component.css"],
})
export class AppComponent implements OnInit {
  public title = "Project-Boards";
  public isIframe = false;
  public loggedIn = false;

  constructor(
    private broadcastService: BroadcastService,
    private authService: MsalService
  ) {}

  ngOnInit() {
    this.isIframe = window !== window.parent && !window.opener;

    this.checkAccount();

    this.broadcastService.subscribe("msal:loginSuccess", () => {
      this.checkAccount();
    });

    this.authService.handleRedirectCallback((authError, response) => {
      if (authError) {
        console.error("Redirect Error: ", authError.errorMessage);
        return;
      }

      console.log("Redirect Success: ", response.accessToken);
    });

    this.authService.setLogger(
      new Logger(
        (logLevel, message, piiEnabled) => {
          console.log("MSAL Logging: ", message);
        },
        {
          correlationId: CryptoUtils.createNewGuid(),
          piiLoggingEnabled: false,
        }
      )
    );
  }

  private checkAccount() {
    this.loggedIn = !!this.authService.getAccount();
  }

  public login() {
    this.authService.loginRedirect();
  }

  public logout() {
    this.authService.logout();
  }
}
