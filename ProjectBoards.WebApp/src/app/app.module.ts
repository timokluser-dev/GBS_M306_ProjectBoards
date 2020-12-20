import { BrowserModule } from "@angular/platform-browser";
import { BrowserAnimationsModule } from "@angular/platform-browser/animations";
import { NgModule } from "@angular/core";

import { MatButtonModule } from "@angular/material/button";
import { MatListModule } from "@angular/material/list";
import { MatToolbarModule } from "@angular/material/toolbar";
import { MatProgressBarModule } from "@angular/material/progress-bar";

import { AppRoutingModule } from "./app-routing.module";
import { AppComponent } from "./app.component";

import { NgxChartsModule } from "@swimlane/ngx-charts";

import { MsalModule, MsalInterceptor } from "@azure/msal-angular";
import { HTTP_INTERCEPTORS, HttpClientModule } from "@angular/common/http";
import { ProfileComponent } from "./profile/profile.component";
import { HomeComponent } from "./home/home.component";
import { PlannerComponent } from "./planner/planner.component";
import { environment } from "src/environments/environment";

const isIE =
  window.navigator.userAgent.indexOf("MSIE ") > -1 ||
  window.navigator.userAgent.indexOf("Trident/") > -1;

@NgModule({
  declarations: [
    AppComponent,
    ProfileComponent,
    HomeComponent,
    PlannerComponent,
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    BrowserAnimationsModule,
    NgxChartsModule,
    HttpClientModule,
    MatToolbarModule,
    MatButtonModule,
    MatListModule,
    MatProgressBarModule,
    AppRoutingModule,
    MsalModule.forRoot(
      {
        auth: {
          clientId: environment.clientId,
          authority: `https://login.microsoftonline.com/${environment.tenant}`,
          redirectUri: window.location.origin,
          postLogoutRedirectUri: window.location.origin,
        },
        cache: {
          cacheLocation: "sessionStorage",
          storeAuthStateInCookie: isIE,
        },
      },
      {
        popUp: false,
        consentScopes: ["user.read", "openid", "profile"],
        unprotectedResources: [],
        protectedResourceMap: [["https://graph.microsoft.com/", ["user.read"]]],
        extraQueryParameters: {},
      }
    ),
  ],
  providers: [
    {
      provide: HTTP_INTERCEPTORS,
      useClass: MsalInterceptor,
      multi: true,
    },
  ],
  bootstrap: [AppComponent],
})
export class AppModule {}
